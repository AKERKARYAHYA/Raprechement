"""
Enterprise Financial Processing & ERP Integration System
Clean Architecture implementation with dynamic tax code resolution from CALCUL sheet,
strict category-based account mapping, unified DD/MM/YYYY date formatting, and dynamic bank selection.
"""

import io
import logging
import os
import re
import unicodedata
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import xlsxwriter
from rapidfuzz import fuzz, process

# ── 1. PAGE CONFIGURATION (MUST BE FIRST STREAMLIT COMMAND) ───────────────────
st.set_page_config(
    page_title="COPAG ERP · Traitement Bancaire Multi-Banque & AX",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 2. LOGGING CONFIGURATION ──────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger("FinancialEngine")

# ── 3. CONFIGURATION LAYER ────────────────────────────────────────────────────
SUPPORTED_BANKS: List[str] = [
    "BMCIL",
    "CAML",
    "BMCE",
    "BMCEL",
    "BPL",
    "SGL",
    "AWBAM",
    "AWBEL",
]

FUZZY_SIMILARITY_THRESHOLD: float = 60.0

KEYWORDS_CONFIG: List[Dict[str, Any]] = [
    {
        "keywords": [
            "TVA IMPAYES CHQ",
            "COMM.CHEQUE DE BANQUE",
            "TVA CHEQUE DE BANQUE",
            "FRAIS IMPAYES CHQ",
        ],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION SUR CHQ IMP MOIS",
    },
    {
        "keywords": [
            "COMMISSION / RDOC",
            "RDOC",
            "COMM.CHEQUE DE BANQUE",
            "Remise import",
            "ICR",
        ],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION / RDOC",
    },
    {
        "keywords": ["COMMISSION / CDOC", "CDOC", "Lclmp", "LCC"],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION / CDOC",
    },
    {
        "keywords": [
            "Frais de produits et services",
            "PRODUITS ET SERVICES",
        ],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION / PRODUIT ET SERVICE",
    },
    {
        "keywords": ["VIREMENT", "COMMISSION / VIREMENTS"],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION / VIREMENTS",
    },
    {
        "keywords": ["TIMBRE", "FRAIS DE TIMBRE", "DROITS DE TIMBRE"],
        "category": "DROITS DE TIMBRE",
        "observation_base": "FRAIS DE TIMBRE",
    },
    {
        "keywords": ["TVA COM VR INT", "COM VR INT"],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION SUR VIREMENTS",
    },
    {
        "keywords": ["TRANSFERT RECU"],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION SUR RPT",
    },
    {
        "keywords": [
            "TVA REM VIRT",
            "FRAIS REM VIRT",
            "TVA VIRT",
            "FRAIS VIRT",
        ],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION/VIREMENTS PAIE MOIS",
    },
    {
        "keywords": ["DOMCILIATION", "DOM.IMPORT"],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION/DOM.IMPORT",
    },
    {
        "keywords": ["TVA SUR AGIOS", "AGIOS A PAYER", "AGIOS"],
        "category": "AGIOS  BANCAIRES",
        "observation_base": "AGIOS AU",
    },
    {
        "keywords": ["TVA/FR.COM.ETR", "FR.DOS.TRF.EMI", "COM.SWIFT.TRF"],
        "category": "FRAIS ET COMMISSIONS BANCAIRES",
        "observation_base": "COMMISSION SUR TRF",
    },
]


# ── 4. UTILITY LAYER ──────────────────────────────────────────────────────────
def normalize_string(text: Optional[str]) -> str:
    """Normalize strings by stripping whitespace, uppercase conversion, and accent removal."""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    text_str = str(text)
    nfkd_form = unicodedata.normalize("NFKD", text_str)
    no_accents = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return re.sub(r"\s+", " ", no_accents).strip().upper()


def format_date_uniform(raw_date: Any) -> str:
    """Standardize any date format (Excel serial, YYYY-MM-DD, DD/MM/YYYY) into DD/MM/YYYY."""
    if pd.isna(raw_date):
        return ""
    raw_str = str(raw_date).strip()
    if not raw_str or raw_str.lower() == "nan":
        return ""
    try:
        if re.match(r"^\d{4}-\d{2}-\d{2}", raw_str):
            dt = pd.to_datetime(raw_str[:10], format="%Y-%m-%d")
        else:
            dt = pd.to_datetime(raw_date, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return raw_str.split(" ")[0]


def extract_taux(type_str: str) -> float:
    """Extract VAT rate based on calculation profile type."""
    s = str(type_str).upper()
    if "10" in s:
        return 0.10
    elif "7" in s:
        return 0.07
    elif "20" in s:
        return 0.20
    return 0.20


def calcul_entries(
    ttc: float, prorata_raw: float, type_str: str
) -> Tuple[float, float, float]:
    """Calculate HT, prorata VAT, and charge amounts."""
    taux = extract_taux(type_str)
    prorata_val = (
        prorata_raw / 100.0 if prorata_raw > 1.0 else prorata_raw
    )
    ht = round(ttc / (1 + taux), 2)
    tva_prorata = round((ttc - ht) * prorata_val, 2)
    charge = round(ttc - tva_prorata, 2)
    return ht, tva_prorata, charge


# ── 5. DATA ACCESS LAYER ──────────────────────────────────────────────────────
@st.cache_data(ttl=5)
def load_plan_smart(file_path: str) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """Load and cache accounting plan and CALCUL tables from Excel."""
    if not os.path.exists(file_path):
        logger.error(f"Accounting plan file not found at: {file_path}")
        return None, None
    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names

        sheet_plan, sheet_calc = None, None
        for s in sheet_names:
            cs = s.strip().lower()
            if "plan" in cs or "comptable" in cs:
                sheet_plan = s
            elif "calc" in cs:
                sheet_calc = s

        if not sheet_plan:
            sheet_plan = sheet_names[0]
        if not sheet_calc and len(sheet_names) > 1:
            sheet_calc = sheet_names[1]

        df_plan = pd.read_excel(file_path, sheet_name=sheet_plan)
        df_plan.columns = [str(c).strip() for c in df_plan.columns]

        df_calc = pd.read_excel(file_path, sheet_name=sheet_calc)
        df_calc.columns = [str(c).strip() for c in df_calc.columns]

        logger.info(f"Successfully loaded accounting plan and CALCUL sheet from {file_path}")
        return df_plan, df_calc
    except Exception as e:
        logger.exception(f"Exception occurred while loading accounting plan: {e}")
        return None, None


# ── 6. BUSINESS LOGIC LAYER ──────────────────────────────────────────────────
def find_dept_column(df_calc: pd.DataFrame) -> str:
    """Strictly prioritize 'Département bénéficiaire' column first."""
    for c in df_calc.columns:
        nc = normalize_string(c)
        if "DEPARTEMENT BENEFICIAIRE" in nc or "BENEFICIAIRE" in nc:
            return c
    for c in df_calc.columns:
        nc = normalize_string(c)
        if "DEPARTEMENT" in nc or "DEP" in nc or "ACTIVITE" in nc:
            return c
    return df_calc.columns[0]


def get_department_tax_and_prorata(
    selected_department: str, df_calc: pd.DataFrame
) -> Tuple[str, float]:
    """
    Search the CALCUL worksheet for the selected department in 'Département bénéficiaire',
    retrieve the exact 'TYPE' value for CODETAXE and 'Prorata' value.
    """
    if not selected_department or df_calc is None or df_calc.empty:
        err_msg = "No TYPE found in CALCUL sheet for selected department."
        logger.error(err_msg)
        raise ValueError(err_msg)

    dept_col = find_dept_column(df_calc)
    type_col = next(
        (c for c in df_calc.columns if normalize_string(c) == "TYPE"),
        next((c for c in df_calc.columns if "TYPE" in normalize_string(c)), None),
    )
    prorata_col = next(
        (c for c in df_calc.columns if "PRORATA" in normalize_string(c) and "MNT" not in normalize_string(c)),
        next((c for c in df_calc.columns if "PRORATA" in normalize_string(c)), None),
    )

    if not type_col:
        err_msg = "Column 'TYPE' missing in CALCUL worksheet."
        logger.error(err_msg)
        raise ValueError(err_msg)

    norm_selected = normalize_string(selected_department)
    calc_df = df_calc.copy()
    calc_df["_NORM_DEPT_"] = calc_df[dept_col].apply(normalize_string)

    matched_rows = calc_df[calc_df["_NORM_DEPT_"] == norm_selected]

    if matched_rows.empty:
        for col in df_calc.columns:
            temp_match = calc_df[calc_df[col].apply(normalize_string) == norm_selected]
            if not temp_match.empty:
                matched_rows = temp_match
                break

    if matched_rows.empty:
        err_msg = f"Department '{selected_department}' not found in CALCUL sheet."
        logger.error(err_msg)
        raise ValueError(err_msg)

    valid_types = matched_rows[type_col].dropna().astype(str).str.strip()
    valid_types = valid_types[(valid_types != "") & (valid_types.str.lower() != "nan")]

    if valid_types.empty:
        err_msg = f"No TYPE found in CALCUL sheet for selected department '{selected_department}'."
        logger.error(err_msg)
        raise ValueError(err_msg)

    selected_department_type = valid_types.iloc[0]

    prorata_val = 100.0
    if prorata_col and pd.notna(matched_rows.iloc[0][prorata_col]):
        try:
            raw_p = str(matched_rows.iloc[0][prorata_col]).replace(",", ".").replace("%", "").strip()
            prorata_val = float(raw_p)
        except (ValueError, TypeError):
            prorata_val = 100.0

    return selected_department_type, prorata_val


def find_account_from_plan(
    categorie_text: str, df_plan: pd.DataFrame
) -> str:
    """
    Lookup accounting account strictly using the 'Catégorie' value searched
    against the 'Compte principal2' column of Plan_comptable.xlsx, returning
    the corresponding 'Compte principal'.
    """
    if not categorie_text or df_plan is None or df_plan.empty:
        raise ValueError("Catégorie vide ou plan comptable non disponible.")

    col_compte = next(
        (c for c in df_plan.columns if "compte principal" in c.lower() and "2" not in c.lower()),
        df_plan.columns[0],
    )
    col_label2 = next(
        (
            c
            for c in df_plan.columns
            if "compte principal2" in c.lower()
            or ("principal" in c.lower() and "2" in c.lower())
            or "2" in c.lower()
        ),
        df_plan.columns[1] if len(df_plan.columns) > 1 else df_plan.columns[0],
    )

    norm_cat = normalize_string(categorie_text)

    plan_df = df_plan.copy()
    plan_df["_NORM_LABEL2_"] = plan_df[col_label2].apply(normalize_string)
    plan_df["_STR_COMPTE_"] = plan_df[col_compte].astype(str)

    exact = plan_df[plan_df["_NORM_LABEL2_"] == norm_cat]
    if not exact.empty:
        return str(exact.iloc[0]["_STR_COMPTE_"]).strip()

    partial = plan_df[
        plan_df["_NORM_LABEL2_"].str.contains(norm_cat, na=False, regex=False)
        | plan_df["_NORM_LABEL2_"].apply(lambda x: x in norm_cat if x else False)
    ]
    if not partial.empty:
        return str(partial.iloc[0]["_STR_COMPTE_"]).strip()

    choices = plan_df["_NORM_LABEL2_"].tolist()
    result = process.extractOne(
        norm_cat, choices, scorer=fuzz.token_sort_ratio
    )
    if result:
        best_match_text, score, best_idx = result
        if score >= FUZZY_SIMILARITY_THRESHOLD:
            return str(plan_df.iloc[best_idx]["_STR_COMPTE_"]).strip()

    err_msg = f"No category match found in Compte principal2 for category: '{categorie_text}'"
    logger.error(err_msg)
    raise ValueError(err_msg)


def classify_transaction(description_text: str, bank_name: str) -> Tuple[str, str]:
    """Classify bank transaction description against rules and return category and observation base."""
    text_upper = normalize_string(description_text)
    for item in KEYWORDS_CONFIG:
        if any(normalize_string(kw) in text_upper for kw in item["keywords"]):
            obs_base = f"{bank_name}- {item['observation_base']}"
            return item["category"], obs_base

    default_obs = f"{bank_name}- {description_text}"
    return "COMMISSION", default_obs


def build_bulk_erp_xlsx(all_rows: List[Dict[str, Any]]) -> bytes:
    """Generate professional formatted Dynamics AX Excel workbook."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = wb.add_worksheet("Sheet1")

    hdr = wb.add_format({
        "bold": True,
        "font_name": "Arial",
        "font_size": 10,
        "bg_color": "#F1F5F9",
        "font_color": "#0284C7",
        "border": 1,
        "border_color": "#CBD5E1",
        "align": "center",
    })
    txt = wb.add_format({
        "font_name": "Arial",
        "font_size": 10,
        "border": 1,
        "border_color": "#E2E8F0",
    })
    num = wb.add_format({
        "font_name": "Arial",
        "font_size": 10,
        "num_format": "#,##0.00",
        "border": 1,
        "border_color": "#E2E8F0",
    })

    canevas_cols = [
        "JOURNALNUM",
        "LINENUM",
        "VOUCHERSTD",
        "TRDATEDMY",
        "ACCOUNTTYPE",
        "CUSTACCOUNT",
        "DEBIT",
        "CREDIT",
        "DEVISE",
        "EXCHANGERATE",
        "NUMDOCUMENT",
        "DATEDOCUMENT",
        "PROFILEVALIDATION",
        "ACTIVITE",
        "AGENCEAXE",
        "BU",
        "COSTCENTER",
        "INTERCOS",
        "CODETAXE",
        "GROUPTAXE",
        "INVOICEID",
        "EXCHANGERATE.1",
        "APPROVER",
        "APPROVED",
        "PAYMODE",
        "TYPE_COMPTE",
        "COMPTE",
        "Description",
        "BENEFICIEDID",
        "BENEFICIERBANK",
        "REFpay",
    ]

    for c, col in enumerate(canevas_cols):
        ws.write(0, c, col, hdr)
        ws.set_column(c, c, 20)

    for r, op in enumerate(all_rows, 1):
        ws.write(r, 0, "", txt)
        ws.write(r, 1, op["line_num"], txt)
        ws.write(r, 2, op["voucher_std"], txt)
        ws.write(r, 3, op["op_date"], txt)
        ws.write(r, 4, op["acc_type"], txt)
        ws.write(r, 5, op["cust_account"], txt)
        ws.write(r, 6, op["debit"] if op["debit"] > 0 else "", num)
        ws.write(r, 7, op["credit"] if op["credit"] > 0 else "", num)
        ws.write(r, 8, "MAD", txt)
        ws.write(r, 9, 1, num)
        ws.write(r, 10, "", txt)
        ws.write(r, 11, "", txt)
        ws.write(r, 12, "STD", txt)
        ws.write(r, 13, op["activite"], txt)
        ws.write(r, 14, "", txt)
        ws.write(r, 15, "", txt)
        ws.write(r, 16, "", txt)
        ws.write(r, 17, "", txt)
        ws.write(r, 18, op["code_taxe"], txt)
        ws.write(r, 19, op["group_taxe"], txt)
        ws.write(r, 20, "", txt)
        ws.write(r, 21, "", txt)
        ws.write(r, 22, "", txt)
        ws.write(r, 23, "oui", txt)
        ws.write(r, 24, "", txt)
        ws.write(r, 25, op["type_compte"], txt)
        ws.write(r, 26, "", txt)
        ws.write(r, 27, op["description"], txt)
        ws.write(r, 28, "", txt)
        ws.write(r, 29, "", txt)
        ws.write(r, 30, op["ref_pay"], txt)

    wb.close()
    buf.seek(0)
    return buf.read()


# ── 7. PRESENTATION LAYER & UI APP ────────────────────────────────────────────
if "selected_bank" not in st.session_state:
    st.session_state.selected_bank = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PLAN_FILE = os.path.join(BASE_DIR, "Plan_comptable.xlsx")
if not os.path.exists(EXCEL_PLAN_FILE):
    for file_name in os.listdir(BASE_DIR):
        if file_name.endswith(".xlsx") and (
            "plan" in file_name.lower() or "comptable" in file_name.lower()
        ):
            EXCEL_PLAN_FILE = os.path.join(BASE_DIR, file_name)
            break

df_plan, df_calc = load_plan_smart(EXCEL_PLAN_FILE)

if df_plan is None or df_calc is None:
    st.error(
        "❌ Erreur critique : Impossible de charger le fichier `Plan_comptable.xlsx` ou la feuille `CALCUL`. Veuillez vérifier sa présence dans le répertoire."
    )
    st.stop()

# ── Sidebar Configuration & Order ──────────────────────────────────────────────
st.sidebar.markdown(
    """
    <div style="padding: 5px 0px 15px 0px;">
        <span style="font-size: 1.1em; font-weight: bold; color: #1E3A8A;">⚙️ Configuration ERP</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# 1. Sélectionner la Banque d'abord
selected_bank = st.sidebar.selectbox(
    "Sélectionner la Banque",
    options=["-- Choisir une banque --"] + SUPPORTED_BANKS,
    index=0
    if st.session_state.selected_bank is None
    else SUPPORTED_BANKS.index(st.session_state.selected_bank) + 1
    if st.session_state.selected_bank in SUPPORTED_BANKS
    else 0,
)

if selected_bank != "-- Choisir une banque --":
    st.session_state.selected_bank = selected_bank
    logger.info(f"User selected active bank: {selected_bank}")
else:
    st.session_state.selected_bank = None

# 2. Sélectionner l'Activité / Département تحت اختيار البنية تحتها مباشرة
dept_col_name = find_dept_column(df_calc)
dept_options = sorted(
    [
        d
        for d in df_calc[dept_col_name].dropna().astype(str).str.strip().unique().tolist()
        if d and d.lower() != "nan"
    ]
)

selected_dept = st.sidebar.selectbox(
    "Sélectionner l'Activité / Département", options=dept_options
)

st.title("🏦 COPAG ERP · Traitement Intelligent des Relevés Bancaires")
st.caption(
    " "
)

st.markdown("---")

if not st.session_state.selected_bank:
    # رسالة احترافية بأسلوب أنيق وألوان متناسقة مع هوية COPAG
    st.markdown(
        """
        <div style="background-color: #F0FDF4; border-left: 5px solid #16A34A; padding: 15px; border-radius: 6px; margin-bottom: 20px;">
            <p style="color: #166534; font-weight: 600; margin: 0; font-size: 1.05em;">
                🏛️ Bienvenue sur la plateforme intelligente de gestion des frais et commissions bancaires. Sélectionnez une banque et un département pour automatiser le traitement, le contrôle et l'analyse des opérations bancaires.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

current_bank = st.session_state.selected_bank
st.info(f"🏛️ Banque active sélectionnée : **{current_bank}**")

try:
    selected_department_type, prorata_val = get_department_tax_and_prorata(
        selected_dept, df_calc
    )
    st.caption(
        f"Code Taxe détecté (CALCUL → TYPE) : **{selected_department_type}** | Prorata : **{prorata_val}%**"
    )
except ValueError as tax_err:
    st.error(f"❌ {tax_err}")
    with st.expander("🔍 Diagnostic de la feuille CALCUL"):
        st.write("Colonnes détectées :", list(df_calc.columns))
        st.dataframe(df_calc)
    st.stop()

uploaded_file = st.file_uploader(
    f"📂 Charger le Relevé Bancaire ({current_bank}) [Excel ou CSV]",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)

        logger.info(
            f"File loaded successfully: {uploaded_file.name} for bank {current_bank}"
        )
        st.success(
            f"✓ Fichier chargé avec succès : {uploaded_file.name} ({len(df_raw)} lignes)"
        )

        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        cols = list(df_raw.columns)

        date_col = next((c for c in cols if "date" in c.lower()), cols[0])
        desc_col = next(
            (
                c
                for c in cols
                if any(
                    k in c.lower()
                    for k in ["description", "libelle", "information", "objet"]
                )
            ),
            cols[1],
        )
        debit_col = next((c for c in cols if "debit" in c.lower()), None)

        processed_data = []
        for idx, row in df_raw.iterrows():
            date_val = format_date_uniform(row[date_col])
            desc_val = str(row[desc_col]) if pd.notna(row[desc_col]) else ""

            montant = 0.0
            if debit_col and pd.notna(row[debit_col]):
                try:
                    raw_amt = str(row[debit_col]).replace(",", ".").replace(" ", "").strip()
                    montant = float(raw_amt)
                except (ValueError, TypeError):
                    montant = 0.0

            if montant > 0:
                cat_label, obs_base = classify_transaction(desc_val, current_bank)
                processed_data.append(
                    {
                        "Date": date_val,
                        "Catégorie": cat_label,
                        "Observation": obs_base,
                        "Description RLV": desc_val,
                        "Montant": montant,
                    }
                )

        df_processed = pd.DataFrame(processed_data)

        if df_processed.empty:
            st.warning("⚠️ Aucune ligne de débit/commission trouvée dans le fichier.")
            logger.warning("Processed dataframe is empty.")
        else:
            grouped_summary = (
                df_processed.groupby(
                    ["Date", "Observation", "Catégorie"]
                )
                .agg(
                    Nombre_Operations=("Montant", "count"),
                    Total_Montant=("Montant", "sum"),
                )
                .reset_index()
            )

            grouped_summary["_SORT_DT_"] = pd.to_datetime(
                grouped_summary["Date"], format="%d/%m/%Y", errors="coerce"
            )
            grouped_summary = (
                grouped_summary.sort_values(by=["_SORT_DT_", "Catégorie"])
                .drop(columns=["_SORT_DT_"])
                .reset_index(drop=True)
            )

            editor_df = grouped_summary[
                [
                    "Date",
                    "Catégorie",
                    "Observation",
                    "Nombre_Operations",
                    "Total_Montant",
                ]
            ].copy()
            editor_df.columns = [
                "Date",
                "Catégorie",
                "Observation",
                "Nbr Opérations",
                "Total Montant (MAD)",
            ]

            st.markdown("---")
            st.subheader(
                "2. 📊 Éditeur Interactif · Validation des Catégories & Observations"
            )

            edited_summary = st.data_editor(
                editor_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
            )

            total_global = edited_summary["Total Montant (MAD)"].sum()
            st.metric("💰 Total Général des Frais Validés", f"{total_global:,.2f} MAD")

            dept_tag = f"COPG_{selected_dept.replace(' ', '_')}"
            final_export_rows = []

            for idx_row, row_g in enumerate(edited_summary.itertuples(), 1):
                ttc = float(row_g._5)
                current_obs = str(row_g.Observation)
                current_cat = str(row_g.Catégorie)
                current_date = format_date_uniform(row_g.Date)

                compte_p = find_account_from_plan(current_cat, df_plan)
                activite_val = dept_tag

                b_acc = f"{current_bank}"
                c_acc = f"{compte_p}|{dept_tag}"

                # CONDITION SPECIALE: "DROITS DE TIMBRE" -> 2 lignes seulement (sans TVA)
                if current_cat.strip().upper() == "DROITS DE TIMBRE":
                    op_rows = [
                        {
                            "line_num": 1,
                            "voucher_std": idx_row,
                            "op_date": current_date,
                            "acc_type": "Banque",
                            "cust_account": b_acc,
                            "debit": 0.0,
                            "credit": ttc,
                            "activite": activite_val,
                            "code_taxe": "",
                            "group_taxe": "",
                            "type_compte": "Comptabilité",
                            "description": current_obs,
                            "ref_pay": "",
                        },
                        {
                            "line_num": 1,
                            "voucher_std": idx_row,
                            "op_date": current_date,
                            "acc_type": "Comptabilité",
                            "cust_account": c_acc,
                            "debit": ttc,
                            "credit": 0.0,
                            "activite": activite_val,
                            "code_taxe": "",
                            "group_taxe": "",
                            "type_compte": "Comptabilité",
                            "description": current_obs,
                            "ref_pay": "",
                        },
                    ]
                else:
                    # باقي العمليات العادية (3 أسطر مع حساب TVA و Prorata)
                    ht, tva_prorata, charge = calcul_entries(
                        ttc, prorata_val, selected_department_type
                    )
                    t_acc = f"34552010|{dept_tag}"
                    op_rows = [
                        {
                            "line_num": 1,
                            "voucher_std": idx_row,
                            "op_date": current_date,
                            "acc_type": "Banque",
                            "cust_account": b_acc,
                            "debit": 0.0,
                            "credit": ttc,
                            "activite": activite_val,
                            "code_taxe": "",
                            "group_taxe": "",
                            "type_compte": "Comptabilité",
                            "description": current_obs,
                            "ref_pay": "",
                        },
                        {
                            "line_num": 1,
                            "voucher_std": idx_row,
                            "op_date": current_date,
                            "acc_type": "Comptabilité",
                            "cust_account": c_acc,
                            "debit": charge,
                            "credit": 0.0,
                            "activite": activite_val,
                            "code_taxe": "",
                            "group_taxe": "",
                            "type_compte": "Comptabilité",
                            "description": current_obs,
                            "ref_pay": "",
                        },
                        {
                            "line_num": 1,
                            "voucher_std": idx_row,
                            "op_date": current_date,
                            "acc_type": "Comptabilité",
                            "cust_account": t_acc,
                            "debit": tva_prorata,
                            "credit": 0.0,
                            "activite": activite_val,
                            "code_taxe": selected_department_type,
                            "group_taxe": "",
                            "type_compte": "Comptabilité",
                            "description": current_obs,
                            "ref_pay": "",
                        },
                    ]
                final_export_rows.extend(op_rows)

            st.markdown("---")
            st.subheader(f"3. ⚡ Exportation Directe du Canevas AX ({current_bank})")

            excel_canevas_bytes = build_bulk_erp_xlsx(final_export_rows)
            st.download_button(
                label=f"📥 TÉLÉCHARGER LE FICHIER CANEVAS DYNAMICS AX ({current_bank})",
                data=excel_canevas_bytes,
                file_name=f"CANEVAS_ERP_{current_bank}_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

    except Exception as e:
        logger.exception(f"Unexpected runtime exception during file processing: {e}")
        st.error(f"❌ {e}")