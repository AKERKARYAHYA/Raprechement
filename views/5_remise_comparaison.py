# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import io
import difflib
import re

# ─────────────────────────────────────────────
# CONFIG & STYLES
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

# 🧠 Fonction intelligente et flexible pour détecter les colonnes
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
            if kw_clean in col_clean or col_clean in kw_clean:
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
                
    if highest_ratio >= 0.4:
        return best_match
        
    return None

def clean_amount(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(' ', '').replace(',', '.')
    val_clean = re.sub(r'[^0-9.\-]', '', val_str)
    try:
        return float(val_clean)
    except:
        return 0.0

st.markdown("")
run_col = st.columns([1, 2, 1])[1]
with run_col:
    run_btn = st.button("🚀 Lancer le rapprochement des données", use_container_width=True)

if run_btn:
    if file1 and file2:
        try:
            file1_name = os.path.splitext(file1.name)[0]
            file2_name = os.path.splitext(file2.name)[0]
            
            # Lecture des fichiers
            df1 = pd.read_excel(file1) if file1.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file1)
            df2 = pd.read_excel(file2) if file2.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file2)
            
            # Nettoyage des noms de colonnes et suppression des doublons d'en-têtes
            df1.columns = df1.columns.astype(str).str.strip()
            df2.columns = df2.columns.astype(str).str.strip()
            df1 = df1.loc[:, ~df1.columns.duplicated()]
            df2 = df2.loc[:, ~df2.columns.duplicated()]
            
            # Détection intelligente des colonnes principales (Mots clés élargis)
            ref_keywords = ['reference', 'ref', 'code', 'référence', 'num_chèque', 'num', 'n°', 'cheque', 'chq']
            amt_keywords = ['amount', 'montant', 'solde', 'prix', 'valeur', 'total', 'mnt', 'somme']
            date_keywords = ['date', 'date opération', 'date operation', 'date valeur', 'dt', 'date_op', 'jour']

            ref_col1 = find_column_smart(df1, ref_keywords)
            ref_col2 = find_column_smart(df2, ref_keywords)
            
            amt_col1 = find_column_smart(df1, amt_keywords)
            amt_col2 = find_column_smart(df2, amt_keywords)
            
            date_col1 = find_column_smart(df1, date_keywords)
            date_col2 = find_column_smart(df2, date_keywords)
            
            if not ref_col1 or not ref_col2:
                st.error(f"❌ Impossible de détecter la colonne 'Reference' automatiquement. (Trouvé ds ERP: {ref_col1}, ds RLV: {ref_col2})")
                st.stop()
                
            if not amt_col1 or not amt_col2:
                st.error(f"❌ Impossible de détecter la colonne 'Amount / Montant' automatiquement. (Trouvé ds ERP: {amt_col1}, ds RLV: {amt_col2})")
                st.stop()
            
            # 🛡️ Copie propre et standardisation des noms
            df1_renamed = pd.DataFrame()
            df2_renamed = pd.DataFrame()
            
            df1_renamed['Reference'] = df1[ref_col1].astype(str).str.strip()
            df2_renamed['Reference'] = df2[ref_col2].astype(str).str.strip()
            
            df1_renamed['Amount'] = df1[amt_col1].apply(clean_amount)
            df2_renamed['Amount'] = df2[amt_col2].apply(clean_amount)
            
            # Gestion sécurisée des dates
            if date_col1:
                df1_renamed['Date'] = df1[date_col1]
            else:
                df1_renamed['Date'] = "N/A"
                
            if date_col2 and not date_col1:
                df2_renamed['Date'] = df2[date_col2]
            
            # 1. Calcul rigoureux des occurrences sans erreur de dimension
            occ1 = df1_renamed['Reference'].value_counts().reset_index()
            occ1.columns = ['Reference', f'Occurrences_{file1_name}']
            
            occ2 = df2_renamed['Reference'].value_counts().reset_index()
            occ2.columns = ['Reference', f'Occurrences_{file2_name}']
            
            # 2. Fusion et calcul de l'Écart
            merged = pd.merge(df1_renamed, df2_renamed, on='Reference', suffixes=('_t1', '_t2'))
            merged['Ecart'] = (merged['Amount_t1'] - merged['Amount_t2']).abs().round(2)
            
            # Filtrage selon la tolérance
            matched_df = merged[merged['Ecart'] <= tolerance].copy()
            
            # Attribuer la date finale
            if 'Date' in matched_df.columns:
                matched_df['Date_Final'] = matched_df['Date']
            elif 'Date_t1' in matched_df.columns:
                matched_df['Date_Final'] = matched_df['Date_t1']
            elif 'Date_t2' in matched_df.columns:
                matched_df['Date_Final'] = matched_df['Date_t2']
            else:
                matched_df['Date_Final'] = "N/A"
                
            matched_df['Date'] = matched_df['Date_Final']
            matched_df['Amount'] = matched_df['Amount_t1'].round(2)
            
            matched_df = matched_df[['Date', 'Reference', 'Amount', 'Ecart']].drop_duplicates()
            
            # 3. Construction de la table finale avec les occurrences (bien structurée)
            final_df = pd.merge(matched_df, occ1, on='Reference', how='left')
            final_df = pd.merge(final_df, occ2, on='Reference', how='left')
            
            st.success("✅ Rapprochement effectué avec succès !")
            st.dataframe(final_df, use_container_width=True)
            
            # Génération du fichier Excel (.xlsx)
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
        st.warning("⚠️ Veuillez charger les two fichiers (ERP et RLV) pour continuer.")
