# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import io

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
    /* Izalat l-khotot l-motaqatti3a (dashed borders) dyal uploader */
    div[data-testid="stFileUploader"] {
        border: none !important;
        padding: 0rem;
    }
    div[data-testid="stFileUploader"] > section {
        border: none !important;
    }
    /* Style dyal l-bouton zrayq (Bleu Roi #1e3a8a) */
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

# Title m-centré
st.markdown('<div class="centered-header">📊 Remise des chèques</div>', unsafe_allow_html=True)

# Section de chargement des fichiers
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="section-header">Fichier ERP</div>', unsafe_allow_html=True)
    file1 = st.file_uploader("Importer le fichier ERP (Excel/CSV)", type=["xlsx", "xls", "csv"], key="remise_f1")

with col2:
    st.markdown('<div class="section-header">Fichier RLV</div>', unsafe_allow_html=True)
    file2 = st.file_uploader("Importer le fichier RLV (Excel/CSV)", type=["xlsx", "xls", "csv"], key="remise_f2")

# Paramètres dans la barre latérale
st.sidebar.header("⚙️ Paramètres")
tolerance = st.sidebar.number_input("Définir la tolérance (Tolerance) :", min_value=0.0, value=0.01, step=0.01, key="remise_tol")

# Fonction intelligente pour détecter les colonnes
def find_column(df, possible_names):
    for col in df.columns:
        if str(col).strip().lower() in [p.lower() for p in possible_names]:
            return col
    return None

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
            
            # Nettoyage des noms de colonnes
            df1.columns = df1.columns.astype(str).str.strip()
            df2.columns = df2.columns.astype(str).str.strip()
            
            # Détection dynamique des colonnes principales
            ref_col1 = find_column(df1, ['reference', 'ref', 'code', 'référence'])
            ref_col2 = find_column(df2, ['reference', 'ref', 'code', 'référence'])
            
            amt_col1 = find_column(df1, ['amount', 'montant', 'solde', 'prix', 'valeur'])
            amt_col2 = find_column(df2, ['amount', 'montant', 'solde', 'prix', 'valeur'])
            
            # Détection avancée de la colonne Date
            date_col1 = find_column(df1, ['date', 'date opération', 'date operation', 'date valeur', 'dt', 'date_op'])
            date_col2 = find_column(df2, ['date', 'date opération', 'date operation', 'date valeur', 'dt', 'date_op'])
            
            if not ref_col1 or not ref_col2:
                st.error("❌ Impossible de trouver la colonne 'Reference' dans l'un des deux fichiers.")
                st.stop()
                
            if not amt_col1 or not amt_col2:
                st.error("❌ Impossible de trouver la colonne 'Amount / Montant' dans l'un des deux fichiers.")
                st.stop()
            
            # Renommage temporaire pour le rapprochement
            df1_renamed = df1.copy()
            df2_renamed = df2.copy()
            
            df1_renamed = df1_renamed.rename(columns={ref_col1: 'Reference', amt_col1: 'Amount'})
            df2_renamed = df2_renamed.rename(columns={ref_col2: 'Reference', amt_col2: 'Amount'})
            
            if date_col1:
                df1_renamed = df1_renamed.rename(columns={date_col1: 'Date'})
            elif date_col2:
                df2_renamed = df2_renamed.rename(columns={date_col2: 'Date'})
            
            # 1. Calcul des occurrences de Reference par table
            occ1 = df1_renamed['Reference'].value_counts().reset_index()
            occ1.columns = ['Reference', f'Occurrences_{file1_name}']
            
            occ2 = df2_renamed['Reference'].value_counts().reset_index()
            occ2.columns = ['Reference', f'Occurrences_{file2_name}']
            
            # 2. Fusion et calcul de l'Écart
            merged = pd.merge(df1_renamed, df2_renamed, on='Reference', suffixes=('_t1', '_t2'))
            merged['Ecart'] = (merged['Amount_t1'] - merged['Amount_t2']).abs().round(2)
            
            # Filtrage selon la tolérance
            matched_df = merged[merged['Ecart'] <= tolerance].copy()
            
            # Gestion de la colonne Date dans le résultat
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
            
            # 3. Construction de la table finale avec les occurrences
            final_df = pd.merge(matched_df, occ1, on='Reference', how='left')
            final_df = pd.merge(final_df, occ2, on='Reference', how='left')
            
            st.success("✅ Rapprochement effectué avec succès !")
            st.dataframe(final_df, use_container_width=True)
            
            # Génération du fichier Excel (.xlsx)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Matching_Résultat')
            buffer.seek(0)
            
            # Bouton de téléchargement
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