# -*- coding: utf-8 -*-
"""
ERP Bank Reconciliation Tool (4 Phases + Multiplicity Control)
==============================================================
Phase 1 : ERP.reference found inside RLV.description     + same amount
Phase 2 : fuzzy(ERP.description, RLV.description) >= 80%  + same amount + |date gap| <= 35 d
Phase 3 : fuzzy(ERP.nom,         RLV.description) >= 80%  + same amount + |date gap| <= 35 d
Phase 4 : Amount-only matching on remaining rows:
          - 1-to-1 unique remaining amount -> Matched as Phase 4 (light red highlight for manual check)
          - 1-to-N / N-to-1 / N-to-N repeated amounts -> Grouped for manual inspection
"""

from __future__ import annotations

import io
import re
import unicodedata
import datetime
from collections import defaultdict

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from rapidfuzz import fuzz


# ============================================================================
#  Configuration
# ============================================================================
FUZZY_THRESHOLD = 80
MAX_DATE_GAP_DAYS = 35
REPORT_NAME = "Reconciliation_Final_Report.xlsx"

# الألوان المعتمدة في تقرير الإكسل
FILL_PHASE4_LIGHT_RED = PatternFill("solid", fgColor="ffd700")  # أحمر خفيف لـ Phase 4 (في الورقتين)
FILL_MATCHED_GREEN = PatternFill("solid", fgColor="349908")     # أخضر للمتطابق في Phase 1-3
FILL_UNMATCHED_RED = PatternFill("solid", fgColor="ff0100")     # أحمر لغير المتطابق في RLV
FILL_GROUP_HEADER = PatternFill("solid", fgColor="FEF9E7")      # أصفر خفيف لتنبيه مجموعات التكرار

ROLE_KEYWORDS = {
    "date":        ["date"],
    "credit":      ["credit", "cred", "avoir"],
    "debit":       ["debit", "deb"],
    "document":    ["document", "piece", "justif", "ndoc", "n doc", "num piece"],
    "reference":   ["reference", "ref"],
    "nom":         ["nom", "client", "tiers", "beneficiaire", "raison", "fournisseur", "donneur"],
    "description": ["description", "libelle", "libell", "operation", "detail",
                    "motif", "narration", "objet", "intitule", "desc"],
}


# ============================================================================
#  Generic helpers
# ============================================================================
def deaccent(value) -> str:
    s = unicodedata.normalize("NFD", str(value if value is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").strip().lower()


def clean_text(value) -> str:
    s = str(value).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def to_number(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    cleaned = (series.astype(str)
               .str.replace(r"\s", "", regex=True)
               .str.replace(",", ".", regex=False)
               .str.replace(r"[^0-9.\-]", "", regex=True))
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def parse_date(value):
    if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(value).normalize()
    s = str(value).strip()
    if not s:
        return pd.NaT
    parts = re.split(r"[-/.]", s.split()[0])
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        a, b, c = (int(p) for p in parts)
        try:
            return pd.Timestamp(year=a, month=b, day=c) if a > 31 else pd.Timestamp(year=c, month=b, day=a)
        except ValueError:
            return pd.NaT
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


# ============================================================================
#  Loading & column detection
# ============================================================================
def read_table(upload) -> pd.DataFrame:
    name = upload.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(upload)
    sheets = pd.read_excel(upload, sheet_name=None)
    for sheet_name in sheets:
        if deaccent(sheet_name) in ("erp", "rlv", "releve", "bank", "banque"):
            return sheets[sheet_name]
    return list(sheets.values())[0]


def detect_columns(df: pd.DataFrame) -> dict:
    headers = {c: deaccent(c) for c in df.columns}
    used, mapping = set(), {}
    for role, keywords in ROLE_KEYWORDS.items():
        chosen = None
        for col in df.columns:
            if col in used:
                continue
            if any(kw in headers[col] for kw in keywords):
                chosen = col
                break
        mapping[role] = chosen
        if chosen is not None:
            used.add(chosen)
    return mapping


def normalize(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    cmap = detect_columns(df)
    out = pd.DataFrame(index=df.index)

    out["date"] = df[cmap["date"]].map(parse_date) if cmap["date"] else pd.NaT

    credit = to_number(df[cmap["credit"]]).abs() if cmap["credit"] else pd.Series(0.0, index=df.index)
    debit = to_number(df[cmap["debit"]]).abs() if cmap["debit"] else pd.Series(0.0, index=df.index)
    out["debit"] = debit
    out["credit"] = credit
    out["amount"] = credit.where(credit > 0, debit)
    out["cents"] = (out["amount"] * 100).round().astype("int64")

    out["reference"] = df[cmap["reference"]].map(clean_text) if cmap["reference"] else ""
    out["nom"] = df[cmap["nom"]].map(clean_text) if cmap["nom"] else ""
    out["document"] = df[cmap["document"]].map(clean_text) if cmap["document"] else ""

    if cmap["description"]:
        out["description"] = df[cmap["description"]].map(clean_text)
    else:
        skip = {cmap[r] for r in cmap} | {cmap["date"]}
        text_cols = [c for c in df.columns if c not in skip
                     and not pd.api.types.is_numeric_dtype(df[c])
                     and not pd.api.types.is_datetime64_any_dtype(df[c])]
        out["description"] = (df[text_cols].astype(str).agg(" ".join, axis=1)
                              if text_cols else "")

    out["ref_norm"] = out["reference"].map(deaccent)
    out["nom_norm"] = out["nom"].map(deaccent)
    out["desc_norm"] = out["description"].map(deaccent)
    return out, cmap


# ============================================================================
#  Reconciliation engine (4 Phases + Multiplicity Control)
# ============================================================================
def reconcile(erp: pd.DataFrame, rlv: pd.DataFrame) -> tuple[list[dict], list[dict]]:
    rlv_by_amt: dict[int, list] = defaultdict(list)
    for rj in rlv.index:
        if rlv.at[rj, "cents"] > 0:
            rlv_by_amt[rlv.at[rj, "cents"]].append(rj)

    erp_done, rlv_done, matches = set(), set(), []

    def date_gap_ok(ed, rd):
        return not pd.isna(ed) and not pd.isna(rd) and abs((ed - rd).days) <= MAX_DATE_GAP_DAYS

    # ── Phase 1: Reference in RLV Description + Exact Amount ──
    for ei in erp.index:
        ref = erp.at[ei, "ref_norm"]
        cents = erp.at[ei, "cents"]
        if not ref or cents <= 0:
            continue
        for rj in rlv_by_amt.get(cents, []):
            if rj not in rlv_done and ref in rlv.at[rj, "desc_norm"]:
                erp_done.add(ei)
                rlv_done.add(rj)
                matches.append({"phase": "Phase 1", "score": 100.0, "ei": ei, "rj": rj})
                break

    # ── Phase 2 & 3: Fuzzy Description / Nom + Exact Amount + Date Window ──
    def fuzzy_phase(label, field):
        for ei in erp.index:
            if ei in erp_done:
                continue
            text = erp.at[ei, field]
            cents = erp.at[ei, "cents"]
            ed = erp.at[ei, "date"]
            if not text or cents <= 0:
                continue
            best_rj, best_score = None, FUZZY_THRESHOLD - 1
            for rj in rlv_by_amt.get(cents, []):
                if rj in rlv_done or not date_gap_ok(ed, rlv.at[rj, "date"]):
                    continue
                score = fuzz.token_set_ratio(text, rlv.at[rj, "desc_norm"])
                if score >= FUZZY_THRESHOLD and score > best_score:
                    best_score, best_rj = score, rj
            if best_rj is not None:
                erp_done.add(ei)
                rlv_done.add(best_rj)
                matches.append({"phase": label, "score": float(best_score), "ei": ei, "rj": best_rj})

    fuzzy_phase("Phase 2", "desc_norm")
    fuzzy_phase("Phase 3", "nom_norm")

    # ── Phase 4: Amount-Only Matching on Remaining Rows ──
    rem_erp_by_amt: dict[int, list] = defaultdict(list)
    for ei in erp.index:
        if ei not in erp_done and erp.at[ei, "cents"] > 0:
            rem_erp_by_amt[erp.at[ei, "cents"]].append(ei)

    rem_rlv_by_amt: dict[int, list] = defaultdict(list)
    for rj in rlv.index:
        if rj not in rlv_done and rlv.at[rj, "cents"] > 0:
            rem_rlv_by_amt[rlv.at[rj, "cents"]].append(rj)

    ambiguous_groups = []
    common_cents = sorted(set(rem_erp_by_amt.keys()) & set(rem_rlv_by_amt.keys()))

    for cents in common_cents:
        e_list = rem_erp_by_amt[cents]
        r_list = rem_rlv_by_amt[cents]

        if len(e_list) == 1 and len(r_list) == 1:
            # تطابق 1 مقابل 1 بالمبلغ فقط -> Phase 4 (يلون بالأحمر الخفيف للمراجعة اليدوية)
            ei, rj = e_list[0], r_list[0]
            erp_done.add(ei)
            rlv_done.add(rj)
            matches.append({"phase": "Phase 4", "score": 50.0, "ei": ei, "rj": rj})
        else:
            # تعدد المبالغ (مثلا 1 في ERP مقابل 3 في RLV أو 3 في ERP مقابل 1 في RLV)
            ambiguous_groups.append({
                "cents": cents,
                "amount": cents / 100.0,
                "erp_indices": e_list,
                "rlv_indices": r_list,
            })

    return matches, ambiguous_groups


# ============================================================================
#  Report generation (Styled Excel)
# ============================================================================
def style_header_and_numbers(ws):
    navy, white = "1F3864", "FFFFFF"
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.font = Font(name="Arial", size=10, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = 20
    for col in ws.iter_cols(min_row=2):
        for cell in col:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
    ws.freeze_panes = "A2"


def build_report(erp_raw: pd.DataFrame, rlv_raw: pd.DataFrame,
                 erp: pd.DataFrame, rlv: pd.DataFrame, rlv_cmap: dict,
                 matches: list[dict], ambiguous_groups: list[dict]):

    erp_done = {m["ei"] for m in matches}
    rlv_phase_map = {m["rj"]: m["phase"] for m in matches}

    # ── 1. Sheet 1: Matched_Results ──
    matched_rows = []
    for m in matches:
        e, r = erp.loc[m["ei"]], rlv.loc[m["rj"]]
        matched_rows.append({
            "Match Type": m["phase"],
            "Score (%)": round(m["score"], 1),
            "Montant": e["amount"],
            "ERP Date": e["date"],
            "ERP Référence": e["reference"],
            "ERP Nom": e["nom"],
            "ERP Document": e["document"],
            "ERP Description": e["description"],
            "ERP Débit": e["debit"],
            "ERP Crédit": e["credit"],
            "RLV Date": r["date"],
            "RLV Description": r["description"],
            "RLV Débit": r["debit"],
            "RLV Crédit": r["credit"],
        })

    matched_cols = [
        "Match Type", "Score (%)", "Montant", "ERP Date", "ERP Référence", "ERP Nom",
        "ERP Document", "ERP Description", "ERP Débit", "ERP Crédit",
        "RLV Date", "RLV Description", "RLV Débit", "RLV Crédit"
    ]
    matched_df = pd.DataFrame(matched_rows, columns=matched_cols)

    # ── 2. Sheet 2: RLV_Complet (نسخة طبق الأصل من ملف RLV) ──
    rlv_full_df = rlv_raw.copy().reset_index(drop=True)
    rlv_orig_indices = list(rlv.index)

    # ── 3. Sheet 3: Phase4_Multiples_Control (حالات 1 مقابل عدة أو عدة مقابل 1) ──
    multi_rows = []
    for grp in ambiguous_groups:
        e_list = grp["erp_indices"]
        r_list = grp["rlv_indices"]
        max_len = max(len(e_list), len(r_list))
        cas_label = f"ERP ({len(e_list)}) vs RLV ({len(r_list)})"

        for k in range(max_len):
            row_dict = {
                "Cas_Contrôle": cas_label,
                "Montant_Commun": grp["amount"],
                "Nb_ERP": len(e_list),
                "Nb_RLV": len(r_list),
            }
            if k < len(e_list):
                e = erp.loc[e_list[k]]
                row_dict.update({
                    "ERP Date": e["date"],
                    "ERP Référence": e["reference"],
                    "ERP Nom": e["nom"],
                    "ERP Description": e["description"],
                    "ERP Débit": e["debit"],
                    "ERP Crédit": e["credit"],
                })
            else:
                row_dict.update({
                    "ERP Date": "", "ERP Référence": "", "ERP Nom": "",
                    "ERP Description": "— (Même montant ERP ci-dessus) —",
                    "ERP Débit": None, "ERP Crédit": None,
                })

            if k < len(r_list):
                r = rlv.loc[r_list[k]]
                row_dict.update({
                    "RLV Date": r["date"],
                    "RLV Description": r["description"],
                    "RLV Débit": r["debit"],
                    "RLV Crédit": r["credit"],
                })
            else:
                row_dict.update({
                    "RLV Date": "",
                    "RLV Description": "— (Même montant RLV ci-dessus) —",
                    "RLV Débit": None, "RLV Crédit": None,
                })
            multi_rows.append(row_dict)

    multi_cols = [
        "Cas_Contrôle", "Montant_Commun", "Nb_ERP", "Nb_RLV",
        "ERP Date", "ERP Référence", "ERP Nom", "ERP Description", "ERP Débit", "ERP Crédit",
        "RLV Date", "RLV Description", "RLV Débit", "RLV Crédit"
    ]
    multi_df = pd.DataFrame(multi_rows, columns=multi_cols)

    # ── 4. Sheet 4: Unmatched_ERP ──
    unmatched_erp = erp_raw.loc[[i for i in erp.index if i not in erp_done]].reset_index(drop=True)

    # كتابة ملف الإكسل وتطبيق التنسيقات والألوان
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="dd/mm/yyyy") as xl:
        matched_df.to_excel(xl, sheet_name="Matched_Results", index=False)
        rlv_full_df.to_excel(xl, sheet_name="RLV_Complet", index=False)
        multi_df.to_excel(xl, sheet_name="Phase4_Multiples_Control", index=False)
        unmatched_erp.to_excel(xl, sheet_name="Unmatched_ERP", index=False)

    buf.seek(0)
    wb = load_workbook(buf)

    for ws in wb.worksheets:
        if ws.max_row >= 1:
            style_header_and_numbers(ws)

    # تلوين سطور Phase 4 بالأحمر الخفيف في الورقة الأولى (Matched_Results)
    ws_matched = wb["Matched_Results"]
    for row_idx in range(2, ws_matched.max_row + 1):
        match_type_val = ws_matched.cell(row_idx, 1).value
        if match_type_val == "Phase 4":
            for col_idx in range(1, ws_matched.max_column + 1):
                ws_matched.cell(row_idx, col_idx).fill = FILL_PHASE4_LIGHT_RED

    # تلوين خلية المبلغ فقط (Débit أو Crédit) في الورقة الثانية (RLV_Complet)
    ws_rlv = wb["RLV_Complet"]
    rlv_headers = [ws_rlv.cell(1, c).value for c in range(1, ws_rlv.max_column + 1)]

    debit_col_name = rlv_cmap.get("debit")
    credit_col_name = rlv_cmap.get("credit")

    debit_col_idx = (rlv_headers.index(debit_col_name) + 1) if debit_col_name in rlv_headers else None
    credit_col_idx = (rlv_headers.index(credit_col_name) + 1) if credit_col_name in rlv_headers else None

    for row_pos, orig_idx in enumerate(rlv_orig_indices):
        excel_row = row_pos + 2
        phase = rlv_phase_map.get(orig_idx)  # None إذا كان غير متطابق

        if phase in ("Phase 1", "Phase 2", "Phase 3"):
            target_fill = FILL_MATCHED_GREEN
        elif phase == "Phase 4":
            target_fill = FILL_PHASE4_LIGHT_RED  # نفس الأحمر الخفيف ديال الورقة الأولى
        else:
            target_fill = FILL_UNMATCHED_RED     # أحمر واضح لغير المتطابق

        # تحديد أي خلية تحتوي على المبلغ في هذا السطر (Débit أم Crédit)
        d_val = rlv.at[orig_idx, "debit"]
        c_val = rlv.at[orig_idx, "credit"]

        if credit_col_idx and c_val > 0:
            ws_rlv.cell(excel_row, credit_col_idx).fill = target_fill
        elif debit_col_idx and d_val > 0:
            ws_rlv.cell(excel_row, debit_col_idx).fill = target_fill
        else:
            # في حالة مبلغ صفر، نلون العمود المتوفر
            if credit_col_idx:
                ws_rlv.cell(excel_row, credit_col_idx).fill = target_fill
            elif debit_col_idx:
                ws_rlv.cell(excel_row, debit_col_idx).fill = target_fill

    # تلوين خفيف لورقة مراجعة التكرارات (Phase4_Multiples_Control)
    ws_multi = wb["Phase4_Multiples_Control"]
    for row_idx in range(2, ws_multi.max_row + 1):
        for col_idx in range(1, ws_multi.max_column + 1):
            ws_multi.cell(row_idx, col_idx).fill = FILL_GROUP_HEADER

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), matched_df, rlv_full_df, multi_df, unmatched_erp, rlv_phase_map


# ============================================================================
#  Interface Streamlit
# ============================================================================
st.markdown("""
    <style>
    .centered-header {
        text-align: center;
        color: #1e293b;
        font-size: 1.8rem;
        font-weight: 700;
        margin-top: 5px;
        margin-bottom: 25px;
    }
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1e293b;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 0.3rem;
        margin: 1.2rem 0 0.8rem 0;
    }
    div[data-testid="stFileUploader"] {
        border: none !important;
        padding: 0rem;
    }
    div[data-testid="stFileUploader"] > section {
        border: none !important;
    }
    div.stButton > button {
        background-color: #1e3a8a !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    div.stButton > button:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="centered-header">Rapprochement multi-banques (4 Phases)</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    st.markdown('<div class="section-header">Fichier ERP</div>', unsafe_allow_html=True)
    erp_file = st.file_uploader("Importer le fichier ERP (Excel/CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="p1234_erp")
with c2:
    st.markdown('<div class="section-header">Relevé bancaire (RLV)</div>', unsafe_allow_html=True)
    rlv_file = st.file_uploader("Importer le fichier RLV (Excel/CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="p1234_rlv")

st.markdown("")
run_col = st.columns([1, 2, 1])[1]
with run_col:
    run_btn = st.button("🚀 Lancer le rapprochement des données", use_container_width=True)

if run_btn:
    if erp_file and rlv_file:
        try:
            with st.spinner("Rapprochement en cours…"):
                erp_raw = read_table(erp_file)
                rlv_raw = read_table(rlv_file)
                erp, _ = normalize(erp_raw)
                rlv, rlv_cmap = normalize(rlv_raw)
                matches, ambiguous_groups = reconcile(erp, rlv)
                (report_bytes, matched_df, rlv_full_df,
                 multi_df, un_erp, rlv_phase_map) = build_report(
                    erp_raw, rlv_raw, erp, rlv, rlv_cmap, matches, ambiguous_groups
                )

            p1 = sum(m["phase"] == "Phase 1" for m in matches)
            p2 = sum(m["phase"] == "Phase 2" for m in matches)
            p3 = sum(m["phase"] == "Phase 3" for m in matches)
            p4 = sum(m["phase"] == "Phase 4" for m in matches)
            total = len(matches)
            unmatched_rlv_count = len(rlv_raw) - len(rlv_phase_map)

            st.subheader("📊 Indicateurs")
            r1 = st.columns(5)
            r1[0].metric("ERP Rows", len(erp_raw))
            r1[1].metric("RLV Rows", len(rlv_raw))
            r1[2].metric("Phase 1", p1)
            r1[3].metric("Phase 2", p2)
            r1[4].metric("Phase 3", p3)

            r2 = st.columns(5)
            r2[0].metric("Phase 4 (1↔1)", p4)
            r2[1].metric("Cas Multiples (1↔N)", len(ambiguous_groups))
            r2[2].metric("Total Matches", total)
            r2[3].metric("Unmatched ERP", len(un_erp))
            r2[4].metric("Unmatched RLV", unmatched_rlv_count)

            # عرض جدول Matched_Results مع تلوين Phase 4 بالأحمر الخفيف في الواجهة أيضا
            st.subheader("✅ Feuille 1 : Matched_Results")
            def highlight_p4_ui(row):
                if row["Match Type"] == "Phase 4":
                    return ["background-color: #FADBD8"] * len(row)
                return [""] * len(row)

            st.dataframe(
                matched_df.style.apply(highlight_p4_ui, axis=1),
                use_container_width=True,
                hide_index=True
            )

            # عرض حالات 1 مقابل عدة أو عدة مقابل 1 للمراقبة اليدوية
            if not multi_df.empty:
                st.subheader(f"🔍 Contrôle Manuel : Montants Multiples (1↔N / N↔1) ({len(ambiguous_groups)} cas)")
                st.dataframe(multi_df, use_container_width=True, hide_index=True)

            # عرض RLV الكامل في الواجهة
            with st.expander(f"📄 Feuille 2 : RLV Complet ({len(rlv_full_df)} lignes)"):
                st.dataframe(rlv_full_df, use_container_width=True, hide_index=True)

            with st.expander(f"⚠️ Unmatched ERP ({len(un_erp)})"):
                st.dataframe(un_erp, use_container_width=True, hide_index=True)

            st.download_button(
                "⬇️ Télécharger Reconciliation_Final_Report.xlsx",
                data=report_bytes,
                file_name=REPORT_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        except Exception as exc:
            st.error(f"Impossible de traiter les fichiers : {exc}")
    else:
        st.warning("⚠️ Veuillez d'abord importer les deux fichiers (ERP et RLV) !")
