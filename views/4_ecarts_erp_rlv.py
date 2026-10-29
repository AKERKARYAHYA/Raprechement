# -*- coding: utf-8 -*-
"""
Application Streamlit — Réconciliation ERP / RLV avec Matching Avancé
(Many-to-One Same-Day + Unlimited Cross-Date Matching)
---------------------------------------------------------------------------------------
Auteur : AI Treasury Partner
"""

import io
import datetime
import re
import unicodedata
import collections
import itertools

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# LOGIQUE MÉTIER & UTILITAIRES
# ============================================================

def deaccent(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").strip().lower()


def norm(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return datetime.date(v.year, v.month, v.day)
    if isinstance(v, str):
        p = re.split(r"[-/]", v.strip())
        if len(p) == 3 and all(x.isdigit() for x in p):
            dd, mm, yy = (int(x) for x in p)
            return datetime.date(yy, mm, dd)
    return None


def num(v):
    return v if isinstance(v, (int, float)) else 0.0


def pick_sheet(wb, name: str):
    for s in wb.sheetnames:
        if deaccent(s) == name:
            return wb[s]
    return wb[wb.sheetnames[0]]


def find_col(ws, *keywords) -> int:
    for c in range(1, ws.max_column + 1):
        h = deaccent(ws.cell(1, c).value)
        if any(kw in h for kw in keywords):
            return c
    raise ValueError(f"Colonne {keywords} introuvable dans la feuille « {ws.title} ».")


def find_desc_col(ws):
    keywords = ("desc", "descreption", "libelle", "motif", "details", "texte")
    for c in range(1, ws.max_column + 1):
        h = deaccent(ws.cell(1, c).value)
        if any(kw in h for kw in keywords):
            return c
    return None


def copy_sheet(src, dst):
    for row in src.iter_rows():
        for cell in row:
            nc = dst.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                nc.number_format = cell.number_format


def add_key(ws, date_col: int) -> str:
    k = ws.max_column + 1
    ws.cell(1, k).value = "Date (clé)"
    for r in range(2, ws.max_row + 1):
        d = norm(ws.cell(r, date_col).value)
        ws.cell(r, k).value = datetime.datetime(d.year, d.month, d.day) if d else None
        ws.cell(r, k).number_format = "dd/mm/yyyy"
    return get_column_letter(k)


def extract_raw(ws, date_col, deb_col, cre_col, desc_col):
    rows = []
    for r in range(2, ws.max_row + 1):
        d = norm(ws.cell(r, date_col).value)
        if d is None:
            continue
        desc = ws.cell(r, desc_col).value if desc_col else None
        rows.append({
            "_date_key": d,
            "Ligne source": r,
            "Description / Libellé": desc if desc not in (None, "") else "—",
            "Montant Débit": round(num(ws.cell(r, deb_col).value), 2),
            "Montant Crédit": round(num(ws.cell(r, cre_col).value), 2),
        })
    return pd.DataFrame(rows)


# ============================================================
# ALGORITHME DE MATCHING AVANCÉ
# ============================================================

def match_many_to_one(erp_items, rlv_items, max_combo=4):
    erp_left = list(erp_items)
    rlv_left = list(rlv_items)

    changed = True
    while changed:
        changed = False
        for r_idx, (r_m, r_l, r_d) in enumerate(rlv_left):
            found_combo = None
            for k in range(2, min(max_combo + 1, len(erp_left) + 1)):
                for combo in itertools.combinations(range(len(erp_left)), k):
                    if round(sum(erp_left[i][0] for i in combo), 2) == round(r_m, 2):
                        found_combo = combo
                        break
                if found_combo:
                    break
            if found_combo:
                for i in sorted(found_combo, reverse=True):
                    erp_left.pop(i)
                rlv_left.pop(r_idx)
                changed = True
                break

    changed = True
    while changed:
        changed = False
        for e_idx, (e_m, e_l, e_d) in enumerate(erp_left):
            found_combo = None
            for k in range(2, min(max_combo + 1, len(rlv_left) + 1)):
                for combo in itertools.combinations(range(len(rlv_left)), k):
                    if round(sum(rlv_left[i][0] for i in combo), 2) == round(e_m, 2):
                        found_combo = combo
                        break
                if found_combo:
                    break
            if found_combo:
                for i in sorted(found_combo, reverse=True):
                    rlv_left.pop(i)
                erp_left.pop(e_idx)
                changed = True
                break

    return erp_left, rlv_left


def master_matching_engine(erp_raw: pd.DataFrame, rlv_raw: pd.DataFrame, dates) -> pd.DataFrame:
    records = []

    for sens, col in (("Débit", "Montant Débit"), ("Crédit", "Montant Crédit")):
        erp_sens = erp_raw[erp_raw[col] > 0].copy() if not erp_raw.empty else pd.DataFrame()
        rlv_sens = rlv_raw[rlv_raw[col] > 0].copy() if not rlv_raw.empty else pd.DataFrame()

        erp_by_date = {}
        if not erp_sens.empty:
            for d, g in erp_sens.groupby("_date_key"):
                erp_by_date[d] = list(g[[col, "Ligne source", "Description / Libellé"]].itertuples(index=False, name=None))

        rlv_by_date = {}
        if not rlv_sens.empty:
            for d, g in rlv_sens.groupby("_date_key"):
                rlv_by_date[d] = list(g[[col, "Ligne source", "Description / Libellé"]].itertuples(index=False, name=None))

        all_dates = sorted(set(erp_by_date.keys()) | set(rlv_by_date.keys()))

        pending_erp = []
        pending_rlv = []

        for d in all_dates:
            e_list = erp_by_date.get(d, [])
            r_list = rlv_by_date.get(d, [])

            e_counts = collections.Counter(round(m, 2) for m, _, _ in e_list)
            r_counts = collections.Counter(round(m, 2) for m, _, _ in r_list)
            common = e_counts & r_counts

            c_erp = collections.Counter(common)
            c_rlv = collections.Counter(common)

            e_after_1to1 = []
            for m, l, desc in e_list:
                key = round(m, 2)
                if c_erp[key] > 0:
                    c_erp[key] -= 1
                else:
                    e_after_1to1.append((m, l, desc))

            r_after_1to1 = []
            for m, l, desc in r_list:
                key = round(m, 2)
                if c_rlv[key] > 0:
                    c_rlv[key] -= 1
                else:
                    r_after_1to1.append((m, l, desc))

            e_orph_day, r_orph_day = match_many_to_one(e_after_1to1, r_after_1to1)

            for m, l, desc in e_orph_day:
                pending_erp.append({"_date_key": d, "montant": round(m, 2), "ligne": l, "desc": desc})
            for m, l, desc in r_orph_day:
                pending_rlv.append({"_date_key": d, "montant": round(m, 2), "ligne": l, "desc": desc})

        erp_final_orphans = []
        rlv_cross_pool = list(pending_rlv)

        for e_item in pending_erp:
            e_amt = e_item["montant"]
            match_idx = None
            for idx, r_item in enumerate(rlv_cross_pool):
                if r_item["montant"] == e_amt:
                    match_idx = idx
                    break
            if match_idx is not None:
                rlv_cross_pool.pop(match_idx)
            else:
                erp_final_orphans.append(e_item)

        rlv_final_orphans = rlv_cross_pool

        for item in erp_final_orphans:
            records.append({
                "_date_key": item["_date_key"],
                "Date": item["_date_key"].strftime("%d/%m/%Y"),
                "Sens": sens,
                "Origine": "ERP",
                "Ligne source": item["ligne"],
                "Description / Libellé": item["desc"],
                "Montant": item["montant"],
            })
        for item in rlv_final_orphans:
            records.append({
                "_date_key": item["_date_key"],
                "Date": item["_date_key"].strftime("%d/%m/%Y"),
                "Sens": sens,
                "Origine": "RLV",
                "Ligne source": item["ligne"],
                "Description / Libellé": item["desc"],
                "Montant": item["montant"],
            })

    cols = ["_date_key", "Date", "Sens", "Origine", "Ligne source", "Description / Libellé", "Montant"]
    df_res = pd.DataFrame(records, columns=cols)
    if not df_res.empty:
        df_res = df_res.sort_values(by=["_date_key", "Sens", "Origine"]).reset_index(drop=True)
    return df_res


def build(erp_file, rlv_file):
    erp_src = pick_sheet(load_workbook(erp_file), "erp")
    rlv_src = pick_sheet(load_workbook(rlv_file), "rlv")

    wb = Workbook()
    wb.remove(wb.active)
    erp = wb.create_sheet("ERP"); copy_sheet(erp_src, erp)
    rlv = wb.create_sheet("RLV"); copy_sheet(rlv_src, rlv)

    erp_date, erp_deb, erp_cre = find_col(erp, "date"), find_col(erp, "deb"), find_col(erp, "cred")
    rlv_date, rlv_deb, rlv_cre = find_col(rlv, "date"), find_col(rlv, "deb"), find_col(rlv, "cred")
    erp_desc = find_desc_col(erp)
    rlv_desc = find_desc_col(rlv)

    erp_last, rlv_last = erp.max_row, rlv.max_row
    ERP_DEB, ERP_CRE = get_column_letter(erp_deb), get_column_letter(erp_cre)
    RLV_DEB, RLV_CRE = get_column_letter(rlv_deb), get_column_letter(rlv_cre)
    ERP_KEY, RLV_KEY = add_key(erp, erp_date), add_key(rlv, rlv_date)

    erp_raw = extract_raw(erp, erp_date, erp_deb, erp_cre, erp_desc)
    rlv_raw = extract_raw(rlv, rlv_date, rlv_deb, rlv_cre, rlv_desc)

    ed = collections.defaultdict(lambda: [0.0, 0.0])
    rd = collections.defaultdict(lambda: [0.0, 0.0])
    for r in range(2, erp_last + 1):
        d = norm(erp.cell(r, erp_date).value)
        if d:
            ed[d][0] += num(erp.cell(r, erp_deb).value)
            ed[d][1] += num(erp.cell(r, erp_cre).value)
    for r in range(2, rlv_last + 1):
        d = norm(rlv.cell(r, rlv_date).value)
        if d:
            rd[d][0] += num(rlv.cell(r, rlv_deb).value)
            rd[d][1] += num(rlv.cell(r, rlv_cre).value)
    dates = sorted(set(ed) | set(rd))

    rows = []
    for d in dates:
        a = round(rd[d][0] - ed[d][0], 2)
        b = round(rd[d][1] - ed[d][1], 2)
        rows.append({
            "Date": d.strftime("%d/%m/%Y"),
            "_date_key": d,
            "Total Débit (ERP)": round(ed[d][0], 2),
            "Total Débit (RLV)": round(rd[d][0], 2),
            "Écart Débit": a,
            "Total Crédit (ERP)": round(ed[d][1], 2),
            "Total Crédit (RLV)": round(rd[d][1], 2),
            "Écart Crédit": b,
            "Écart Global (Créd - Déb)": round(b - a, 2),
        })
    preview = pd.DataFrame(rows)

    anomalies_df = master_matching_engine(erp_raw, rlv_raw, dates)

    ws = wb.create_sheet("Output_Ecarts")
    FONT = "Arial"
    navy, blue, band, white = "1F3864", "1F4E79", "F2F5FA", "FFFFFF"
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = ["Date", "Total Débit (ERP)", "Total Débit (RLV)", "Écart Débit",
               "Total Crédit (ERP)", "Total Crédit (RLV)", "Écart Crédit", "Écart Global (Créd - Déb)"]
    ncol = len(headers)
    HDR_R, FIRST = 3, 4
    EC = (4, 7, 8)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    t = ws.cell(1, 1, "Réconciliation ERP / RLV — Écarts journaliers")
    t.font = Font(name=FONT, size=14, bold=True, color=white)
    t.alignment = Alignment(horizontal="center", vertical="center")
    t.fill = PatternFill("solid", fgColor=navy)
    ws.row_dimensions[1].height = 26
    sub = ws.cell(2, 1, f"Période : {dates[0]:%d/%m/%Y} – {dates[-1]:%d/%m/%Y}  •  Montants en MAD")
    sub.font = Font(name=FONT, size=9, italic=True, color="595959")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)

    for c, h in enumerate(headers, 1):
        cell = ws.cell(HDR_R, c, h)
        cell.font = Font(name=FONT, size=10, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[HDR_R].height = 32

    amt = "#,##0.00"
    ec = "[Green]+#,##0.00;[Red]-#,##0.00;0.00"
    ec_total = "+#,##0.00;-#,##0.00;0.00"

    for i, d in enumerate(dates):
        r = FIRST + i
        a = ws.cell(r, 1, datetime.datetime(d.year, d.month, d.day))
        a.number_format = "dd/mm/yyyy"
        ws.cell(r, 2, f"=SUMIFS(ERP!${ERP_DEB}$2:${ERP_DEB}${erp_last},ERP!${ERP_KEY}$2:${ERP_KEY}${erp_last},$A{r})")
        ws.cell(r, 3, f"=SUMIFS(RLV!${RLV_DEB}$2:${RLV_DEB}${rlv_last},RLV!${RLV_KEY}$2:${RLV_KEY}${rlv_last},$A{r})")
        ws.cell(r, 4, f"=ROUND(C{r}-B{r},2)")
        ws.cell(r, 5, f"=SUMIFS(ERP!${ERP_CRE}$2:${ERP_CRE}${erp_last},ERP!${ERP_KEY}$2:${ERP_KEY}${erp_last},$A{r})")
        ws.cell(r, 6, f"=SUMIFS(RLV!${RLV_CRE}$2:${RLV_CRE}${rlv_last},RLV!${RLV_KEY}$2:${RLV_KEY}${rlv_last},$A{r})")
        ws.cell(r, 7, f"=ROUND(F{r}-E{r},2)")
        ws.cell(r, 8, f"=ROUND(G{r}-D{r},2)")
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.font = Font(name=FONT, size=10)
            if c == 1:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.number_format = ec if c in EC else amt
            if i % 2 and c not in EC:
                cell.fill = PatternFill("solid", fgColor=band)

    LAST = FIRST + len(dates) - 1
    TOT = LAST + 1
    ws.cell(TOT, 1, "Total").font = Font(name=FONT, size=10, bold=True, color=white)
    ws.cell(TOT, 1).alignment = Alignment(horizontal="center")
    for c in range(2, ncol + 1):
        col = get_column_letter(c)
        cell = ws.cell(TOT, c, f"=SUM({col}{FIRST}:{col}{LAST})")
        cell.number_format = ec_total if c in EC else amt
        cell.font = Font(name=FONT, size=10, bold=True, color=white)
    for c in range(1, ncol + 1):
        ws.cell(TOT, c).fill = PatternFill("solid", fgColor=navy)
        ws.cell(TOT, c).border = border
    ws.row_dimensions[TOT].height = 20

    ws.column_dimensions["A"].width = 13
    for c in range(2, ncol + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False

    wsa = wb.create_sheet("Details_Anomalies")
    a_headers = ["Date", "Sens", "Origine", "Ligne source", "Description / Libellé", "Montant"]
    a_ncol = len(a_headers)

    wsa.merge_cells(start_row=1, start_column=1, end_row=1, end_column=a_ncol)
    ta = wsa.cell(1, 1, "Détail des anomalies réelles (écarts non réconciliés)")
    ta.font = Font(name=FONT, size=14, bold=True, color=white)
    ta.alignment = Alignment(horizontal="center", vertical="center")
    ta.fill = PatternFill("solid", fgColor=navy)
    wsa.row_dimensions[1].height = 26
    suba = wsa.cell(2, 1, "Écritures orphelines (excluant le lettrage 1-à-1, Many-to-One et le Cross-Date Multi-Dates)")
    suba.font = Font(name=FONT, size=9, italic=True, color="595959")
    wsa.merge_cells(start_row=2, start_column=1, end_row=2, end_column=a_ncol)

    A_HDR_R, A_FIRST = 3, 4
    for c, h in enumerate(a_headers, 1):
        cell = wsa.cell(A_HDR_R, c, h)
        cell.font = Font(name=FONT, size=10, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    wsa.row_dimensions[A_HDR_R].height = 28

    erp_tint, rlv_tint = "FDEBD0", "D6EAF8"
    if anomalies_df.empty:
        r = A_FIRST
        wsa.merge_cells(start_row=r, start_column=1, end_row=r, end_column=a_ncol)
        msg = wsa.cell(r, 1, "Aucune anomalie réelle : toutes les écritures sont parfaitement réconciliées. ✅")
        msg.font = Font(name=FONT, size=10, italic=True, color="595959")
        msg.alignment = Alignment(horizontal="center")
    else:
        for i, (_, rowd) in enumerate(anomalies_df.iterrows()):
            r = A_FIRST + i
            values = [rowd["Date"], rowd["Sens"], rowd["Origine"], rowd["Ligne source"],
                      rowd["Description / Libellé"], rowd["Montant"]]
            for c, v in enumerate(values, 1):
                cell = wsa.cell(r, c, v)
                cell.border = border
                cell.font = Font(name=FONT, size=10)
                if c == 6:
                    cell.number_format = "#,##0.00"
                else:
                    cell.alignment = Alignment(horizontal="center")
                cell.fill = PatternFill("solid", fgColor=erp_tint if rowd["Origine"] == "ERP" else rlv_tint)

    wsa.column_dimensions["A"].width = 13
    wsa.column_dimensions["B"].width = 12
    wsa.column_dimensions["C"].width = 12
    wsa.column_dimensions["D"].width = 14
    wsa.column_dimensions["E"].width = 45
    wsa.column_dimensions["F"].width = 16
    wsa.freeze_panes = "A4"
    wsa.sheet_view.showGridLines = False

    wb.move_sheet("Output_Ecarts", offset=-(len(wb.sheetnames) - 3))
    wb.move_sheet("Details_Anomalies", offset=-(len(wb.sheetnames) - 4))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), preview, anomalies_df


@st.cache_data(show_spinner=False)
def build_cached(erp_bytes: bytes, rlv_bytes: bytes):
    return build(io.BytesIO(erp_bytes), io.BytesIO(rlv_bytes))


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

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
    /* Style m-khass b l-bouton bash ykon azraq (Bleu Roi #1e3a8a) kima bġiti */
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

st.markdown('<div class="centered-header"> Écarts  ERP / RLV</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="section-header">Fichier ERP</div>', unsafe_allow_html=True)
    erp_up = st.file_uploader("Importer le fichier ERP (Excel/CSV)", type=["xlsx", "xlsm", "csv"], key="ecart_erp")
with col2:
    st.markdown('<div class="section-header">Relevé bancaire (RLV)</div>', unsafe_allow_html=True)
    rlv_up = st.file_uploader("Importer le Relevé bancaire (RLV) (Excel/CSV)", type=["xlsx", "xlsm", "csv"], key="ecart_rlv")

st.markdown("")
run_col = st.columns([1, 2, 1])[1]
with run_col:
    run_btn = st.button("🚀 Lancer le rapprochement des données", use_container_width=True)

if run_btn:
    if erp_up and rlv_up:
        try:
            with st.spinner("Calcul des écarts et exécution du matching intelligent sur tout le mois…"):
                xlsx_bytes, preview, anomalies_df = build_cached(erp_up.getvalue(), rlv_up.getvalue())
            
            tot_deb = preview["Écart Débit"].sum()
            tot_cre = preview["Écart Crédit"].sum()
            n_days = len(preview)
            n_anomalies = len(anomalies_df["Date"].unique()) if not anomalies_df.empty else 0

            m = st.columns(4)
            m[0].metric("Jours analysés", n_days)
            m[1].metric("Jours avec anomalies réelles", n_anomalies)
            m[2].metric("Total Écart Débit", f"{tot_deb:+,.2f}" if round(tot_deb, 2) else "0.00")
            m[3].metric("Total Écart Crédit", f"{tot_cre:+,.2f}" if round(tot_cre, 2) else "0.00")

            st.subheader("Tableau des écarts journaliers")
            st.caption("👉 Cliquez sur une ligne pour afficher les anomalies réelles non réconciliées (excluant les décalages de dates et montants denses).")

            display_cols = ["Date", "Total Débit (ERP)", "Total Débit (RLV)", "Écart Débit",
                            "Total Crédit (ERP)", "Total Crédit (RLV)", "Écart Crédit", "Écart Global (Créd - Déb)"]
            ecart_cols = ["Écart Débit", "Écart Crédit", "Écart Global (Créd - Déb)"]
            plain_cols = [c for c in display_cols if c not in ecart_cols and c != "Date"]

            def fmt_signed(v):
                if not isinstance(v, (int, float)):
                    return v
                v = round(v, 2)
                return "0.00" if v == 0 else f"{v:+,.2f}"

            def colorize(v):
                if not isinstance(v, (int, float)):
                    return ""
                v = round(v, 2)
                if v > 0:
                    return "color:#008000;font-weight:bold"
                if v < 0:
                    return "color:#FF0000;font-weight:bold"
                return ""

            fmt_map = {c: "{:,.2f}" for c in plain_cols}
            fmt_map.update({c: fmt_signed for c in ecart_cols})
            styler = preview[display_cols].style.map(colorize, subset=ecart_cols).format(fmt_map)

            selected_date_str = None
            selection_supported = True
            try:
                event = st.dataframe(
                    styler,
                    use_container_width=True,
                    hide_index=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key="ecart_preview_table",
                )
                sel_rows = list(event.selection.get("rows", [])) if hasattr(event, "selection") else []
                if sel_rows:
                    selected_date_str = preview.iloc[sel_rows[0]]["Date"]
            except TypeError:
                selection_supported = False

            if not selection_supported:
                st.dataframe(styler, use_container_width=True, hide_index=True)
                choix = st.selectbox(
                    "Choisir une date à détailler",
                    options=["—"] + preview["Date"].tolist(),
                    key="ecart_date_selectbox",
                )
                if choix != "—":
                    selected_date_str = choix

            if selected_date_str:
                st.markdown(f"### Anomalies réelles — {selected_date_str}")
                detail_jour = anomalies_df[anomalies_df["Date"] == selected_date_str] if not anomalies_df.empty else pd.DataFrame()

                if detail_jour.empty:
                    st.success("Toutes les écritures de cette journée ont été réconciliées (ou lettrées via Cross-Date / Combinaisons). Aucun montant orphelin réel ! ✅")
                else:
                    d1, d2 = st.columns(2)
                    with d1:
                        st.caption("ERP — Écritures orphelines réelles")
                        sub_erp = detail_jour[detail_jour["Origine"] == "ERP"][["Sens", "Ligne source", "Description / Libellé", "Montant"]]
                        st.dataframe(sub_erp, use_container_width=True, hide_index=True) if not sub_erp.empty \
                            else st.info("Aucune anomalie ERP ce jour.")
                    with d2:
                        st.caption("RLV — Écritures orphelines réelles")
                        sub_rlv = detail_jour[detail_jour["Origine"] == "RLV"][["Sens", "Ligne source", "Description / Libellé", "Montant"]]
                        st.dataframe(sub_rlv, use_container_width=True, hide_index=True) if not sub_rlv.empty \
                            else st.info("Aucune anomalie RLV ce jour.")

            st.subheader("Export")
            st.download_button(
                "Télécharger le classeur Excel complet",
                data=xlsx_bytes,
                file_name="ECART RECONCILIATION.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            st.info("Le fichier téléchargé contient les 4 feuilles : ERP, RLV, Output_Ecarts et Details_Anomalies (filtrée des faux écarts).")
        except Exception as e:
            st.error(f"Impossible de traiter les fichiers : {e}")
    else:
        st.warning("⚠️ Veuillez d'abord importer les deux fichiers ERP et RLV !")