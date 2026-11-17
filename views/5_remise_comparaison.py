# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import io
import difflib
import re

# ─────────────────────────────────────────────
# CONFIG & STYLES (Clean & Professional)
# ─────────────────────────────────────────────
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

st.markdown('<div class="centered-header">📊 Remise des chèques</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="section-header">Fichier ERP</div>', unsafe_allow_html=True)
    file1 = st.file_uploader("Importer le fichier ERP (Excel/CSV)", type=["xlsx", "xls", "csv"], key="remise_f1")

with col2:
    st.markdown('<div class="section-header">Fichier RLV</div>', unsafe_allow_html=True)
    file2 = st.file_uploader("Importer le fichier RLV (Excel/CSV)", type=["xlsx", "xls", "csv"], key="remise_f2")

st.sidebar.header("⚙️ Paramètres")
tolerance = st.sidebar.number_input("Définir la tolérance (Tolerance) :", min_value=0.0, value=0.01, step=0.01, key="remise_tol")

# ─────────────────────────────────────────────
# FONCTIONS INTELLIGENTES (Détection & Extraction)
# ─────────────────────────────────────────────
def clean_string(text):
    if not isinstance(text, str):
        text = str(text)
    return re.sub(r'[^a-z0-9]', '', text.lower())

def find_column_smart(df, possible_keywords):
    columns = list(df.columns)
    for col in columns:
        col_clean = clean_string(col)
        for keyword in possible_keywords:
            kw_clean = clean_string(keyword)
            if kw_clean == col_clean or kw_clean in col_clean:
                return col

    best_match = None
    highest_ratio = 0.0
    for col in columns:
        col_clean = clean_string(col)
        for keyword in possible_keywords:
            kw_clean = clean_string(keyword)
            ratio = difflib.SequenceMatcher(None, col_clean, kw_clean).ratio()
            if ratio > highest_ratio:
                highest_ratio = ratio
                best_match = col
                
    if highest_ratio >= 0.5:
        return best_match
    return None

def normalize_ref(val):
    """Tawhid l-marji3: kay7ayad .0, l-espaces, w l-asfar mn l-bdaya (ex: '09695008' -> '9695008')"""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    s = re.sub(r'\.0$', '', s)
    # Extraction des chiffres uniquement
    digits = re.findall(r'\d+', s)
    if not digits:
        return ""
    # Prendre la plus longue séquence de chiffres et supprimer les zéros initiaux
    longest_num = max(digits, key=len)
    cleaned = longest_num.lstrip('0')
    return cleaned if cleaned else "0"

def extract_ref_from_description(desc_text, erp_refs_set):
    """Kayqallab wst Description dyal RLV 3la ay raqam kaytatabaq m3a chi Reference f ERP"""
    if pd.isna(desc_text):
        return ""
    s = str(desc_text)
    # Jib ga3 l-arqam li kaynin f la description
    numbers = re.findall(r'\d+', s)
    for num in numbers:
        norm = num.lstrip('0')
        if norm in erp_refs_set:
            return norm
    # Ila malqach match direct f set dyal ERP, rja3 atwal raqam kayn f description
    if numbers:
        longest = max(numbers, key=len).lstrip('0')
        return longest if longest else "0"
    return ""

# ─────────────────────────────────────────────
# TRAITEMENT PRINCIPAL
# ─────────────────────────────────────────────
st.markdown("")
run_col = st.columns([1, 2, 1])[1]
with run_col:
    run_btn = st.button("🚀 Lancer le rapprochement des données", use_container_width=True)

if run_btn:
    if file1 and file2:
        try:
            file1_name = os.path.splitext(file1.name)[0]
            file2_name = os.path.splitext(file2.name)[0]
            
            df_erp = pd.read_excel(file1) if file1.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file1)
            df_rlv = pd.read_excel(file2) if file2.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file2)
            
            df_erp.columns = df_erp.columns.astype(str).str.strip()
            df_rlv.columns = df_rlv.columns.astype(str).str.strip()
            
            # Mots-clés pour la détection
            ref_keywords = ['reference', 'ref', 'référence', 'num_chèque', 'cheque', 'chq', 'code', 'num', 'n°']
            desc_keywords = ['description', 'libelle', 'libellé', 'motif', 'detail', 'opération', 'operation']
            amt_keywords = ['amount', 'montant', 'solde', 'prix', 'valeur', 'total', 'mnt', 'credit', 'debit']
            date_keywords = ['date', 'date opération', 'date operation', 'date valeur', 'dt', 'date_op', 'jour']

            # 1. Détection des colonnes ERP
            erp_ref_col = find_column_smart(df_erp, ref_keywords)
            erp_amt_col = find_column_smart(df_erp, amt_keywords)
            erp_date_col = find_column_smart(df_erp, date_keywords)

            # 2. Détection des colonnes RLV (Description, Reference optionnelle, Montant, Date)
            rlv_desc_col = find_column_smart(df_rlv, desc_keywords)
            rlv_ref_col = find_column_smart(df_rlv, ref_keywords)
            rlv_amt_col = find_column_smart(df_rlv, amt_keywords)
            rlv_date_col = find_column_smart(df_rlv, date_keywords)
            
            if not erp_ref_col or not erp_amt_col:
                st.error("❌ Impossible de détecter les colonnes 'Reference' ou 'Montant' dans le fichier ERP.")
                st.stop()
                
            if not rlv_amt_col or (not rlv_desc_col and not rlv_ref_col):
                st.error("❌ Impossible de détecter 'Description'/'Reference' ou 'Montant' dans le fichier RLV.")
                st.stop()
            
            # Préparation ERP
            erp_clean = pd.DataFrame()
            erp_clean['Ref_Unified'] = df_erp[erp_ref_col].apply(normalize_ref)
            erp_clean['Reference_Origine'] = df_erp[erp_ref_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            erp_clean['Amount_ERP'] = pd.to_numeric(df_erp[erp_amt_col].astype(str).str.replace(',', '.').str.replace(' ', ''), errors='coerce').fillna(0.0)
            erp_clean['Date_ERP'] = df_erp[erp_date_col] if erp_date_col else "N/A"
            
            # Filtrer les références vides
            erp_clean = erp_clean[erp_clean['Ref_Unified'] != ""].copy()
            erp_refs_set = set(erp_clean['Ref_Unified'].unique())
            
            # Préparation RLV (Extraction intelligente depuis Description ou colonne Reference)
            rlv_clean = pd.DataFrame()
            if rlv_desc_col:
                rlv_clean['Ref_Unified'] = df_rlv[rlv_desc_col].apply(lambda x: extract_ref_from_description(x, erp_refs_set))
            else:
                rlv_clean['Ref_Unified'] = df_rlv[rlv_ref_col].apply(normalize_ref)
                
            # Fallback si Description n'a rien donné mais une colonne Reference existe dans RLV
            if rlv_desc_col and rlv_ref_col:
                mask_empty = rlv_clean['Ref_Unified'] == ""
                rlv_clean.loc[mask_empty, 'Ref_Unified'] = df_rlv.loc[mask_empty, rlv_ref_col].apply(normalize_ref)

            rlv_clean['Amount_RLV'] = pd.to_numeric(df_rlv[rlv_amt_col].astype(str).str.replace(',', '.').str.replace(' ', ''), errors='coerce').fillna(0.0)
            rlv_clean['Date_RLV'] = df_rlv[rlv_date_col] if rlv_date_col else "N/A"
            rlv_clean['Description_RLV'] = df_rlv[rlv_desc_col] if rlv_desc_col else ""
            
            rlv_clean = rlv_clean[rlv_clean['Ref_Unified'] != ""].copy()

            # 3. Calcul des occurrences par fichier avant fusion
            occ_erp = erp_clean['Ref_Unified'].value_counts().reset_index()
            occ_erp.columns = ['Ref_Unified', f'Occurrences_{file1_name}']
            
            occ_rlv = rlv_clean['Ref_Unified'].value_counts().reset_index()
            occ_rlv.columns = ['Ref_Unified', f'Occurrences_{file2_name}']
            
            # 4. Fusion (Merge) sur la clé unifiée
            merged = pd.merge(erp_clean, rlv_clean, on='Ref_Unified', how='inner')
            
            if merged.empty:
                st.warning("⚠️ Aucune correspondance de référence trouvée entre le fichier ERP et les descriptions du fichier RLV.")
                st.stop()
            
            # 5. Comparaison des montants et calcul de l'Écart
            merged['Ecart'] = (merged['Amount_ERP'].abs() - merged['Amount_RLV'].abs()).abs().round(2)
            
            # Filtrage selon la tolérance définie par l'utilisateur
            matched_df = merged[merged['Ecart'] <= tolerance].copy()
            
            if matched_df.empty:
                st.warning(f"⚠️ Des références ont été trouvées ({len(merged)}), mais aucun montant ne correspond dans la tolérance de {tolerance}.")
                st.stop()
            
            # Mise en forme de la table finale
            matched_df['Date'] = matched_df['Date_RLV'].where(matched_df['Date_RLV'] != "N/A", matched_df['Date_ERP'])
            matched_df['Reference'] = matched_df['Ref_Unified']
            matched_df['Amount'] = matched_df['Amount_ERP'].round(2)
            
            result_cols = ['Date', 'Reference', 'Amount', 'Ecart']
            matched_df = matched_df[result_cols].drop_duplicates()
            
            # 6. Jointure avec les occurrences
            final_df = pd.merge(matched_df, occ_erp, left_on='Reference', right_on='Ref_Unified', how='left').drop(columns=['Ref_Unified'], errors='ignore')
            final_df = pd.merge(final_df, occ_rlv, left_on='Reference', right_on='Ref_Unified', how='left').drop(columns=['Ref_Unified'], errors='ignore')
            
            st.success(f"✅ Rapprochement effectué avec succès ! ({len(final_df)} correspondances validées)")
            st.dataframe(final_df, use_container_width=True)
            
            # 7. Export Excel (.xlsx)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Matching_Résultat')
            buffer.seek(0)
            
            st.download_button(
                label="📥 Télécharger le résultat (Fichier Excel .xlsx)",
                data=buffer,
                file_name="resultat_rapprochement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Une erreur s'est produite lors du traitement : {str(e)}")
    else:
        st.warning("⚠️ Veuillez charger les deux fichiers (ERP et RLV) pour continuer.")
