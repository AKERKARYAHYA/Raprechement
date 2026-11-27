import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment

# Configuration de la page
st.set_page_config(page_title="TrésoMatch Pro - Rapprochement Plus / Moins", layout="wide")

# --- CSS personnalisé pour le bouton d'export ---
st.markdown("""
    <style>
    div.stDownloadButton > button:first-child {
        background-color: #ff4b4b;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        border: none;
        padding: 0.5rem 1rem;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #ff2222;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 TrésoMatch Pro — Rapprochement & Appariement des Montants")
st.write("Importez votre fichier Excel pour identifier automatiquement chaque montant positif (+) et son écriture d'annulation (-), puis exportez uniquement les paires rapprochées avec une mise en couleur distincte.")


def detect_amount_column(df: pd.DataFrame) -> str | None:
    """Détecte automatiquement la colonne des montants."""
    preferred_names = ["montant", "amount", "solde", "montant_erp", "montant erp"]
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if any(name == col_clean or name in col_clean for name in preferred_names):
            numeric_test = pd.to_numeric(df[col], errors="coerce")
            if numeric_test.notna().sum() > 0:
                return col

    for col in df.columns:
        numeric_test = pd.to_numeric(df[col], errors="coerce")
        if numeric_test.notna().sum() > 0:
            return col

    return None


def normalize_compare_value(val) -> str:
    """Normalise la valeur du critère de comparaison secondaire."""
    if pd.isna(val):
        return ""
    return str(val).strip().upper()


uploaded_file = st.file_uploader("Importer le fichier Excel des écritures", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    with st.expander("📄 Aperçu des données importées"):
        st.dataframe(df.head(), use_container_width=True)

    amount_col = detect_amount_column(df)

    if amount_col is None:
        st.error("❌ Impossible de détecter automatiquement une colonne contenant des montants numériques dans ce fichier.")
        st.stop()

    # Critère de comparaison secondaire (optionnel)
    none_option = "Rapprochement effectué uniquement sur le montant"
    compare_options = [none_option] + [c for c in df.columns.tolist() if c != amount_col]
    selected_compare_col = st.selectbox(
        "🔍 Critère secondaire (optionnel)",
        compare_options
    )
    use_extra_col = selected_compare_col != none_option

    if st.button("🚀 Lancer le rapprochement"):
        data = df.copy()
        data[amount_col] = pd.to_numeric(data[amount_col], errors="coerce").fillna(0)

        rows = data.to_dict("records")
        matched_rows = []
        used_indices = set()

        # Appariement des montants positifs (+) avec leur contrepartie négative (-)
        for i, row in enumerate(rows):
            if i in used_indices:
                continue

            val = row[amount_col]
            if val > 0:
                ref_i = normalize_compare_value(row[selected_compare_col]) if use_extra_col else None

                for j, other_row in enumerate(rows):
                    if j == i or j in used_indices:
                        continue

                    if other_row[amount_col] == -val:
                        if use_extra_col:
                            ref_j = normalize_compare_value(other_row[selected_compare_col])
                            if ref_i != ref_j:
                                continue

                        matched_rows.append(row)
                        matched_rows.append(other_row)
                        used_indices.add(i)
                        used_indices.add(j)
                        break

        df_matched = pd.DataFrame(matched_rows)

        st.success("✅ Rapprochement terminé avec succès.")

        # Indicateurs clés
        m1, m2 = st.columns(2)
        with m1:
            st.metric(label="📊 Total des lignes importées", value=len(df))
        with m2:
            st.metric(label="✅ Lignes rapprochées (Paires +/-)", value=len(df_matched))

        st.divider()

        st.subheader("✅ Écritures rapprochées")
        st.dataframe(df_matched, use_container_width=True)

        # --- Génération du classeur Excel avec coloration par paire ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_matched.to_excel(writer, index=False, sheet_name="Paires_Rapprochees")

            workbook = writer.book
            ws = workbook["Paires_Rapprochees"]

            pair_fills = [
                PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid"),
                PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"),
                PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
                PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
                PatternFill(start_color="EAD1DC", end_color="EAD1DC", fill_type="solid"),
                PatternFill(start_color="D0E0E3", end_color="D0E0E3", fill_type="solid"),
                PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid"),
            ]

            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            total_rows = ws.max_row
            color_index = 0

            for r in range(2, total_rows + 1, 2):
                current_fill = pair_fills[color_index % len(pair_fills)]
                for row_idx in range(r, min(r + 2, total_rows + 1)):
                    for col_num in range(1, ws.max_column + 1):
                        ws.cell(row=row_idx, column=col_num).fill = current_fill
                color_index += 1

        processed_data = output.getvalue()

        st.divider()
        st.subheader("⬇️ Export")

        st.download_button(
            label="📥 Télécharger le classeur Excel complet",
            data=processed_data,
            file_name="TrésoMatch_Paires_Rapprochees.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_button"
        )

        info_text = (
            f"💡 Rapprochement effectué sur le montant et le critère secondaire ({selected_compare_col})."
            if use_extra_col
            else "💡 Rapprochement effectué uniquement sur les montants (+/-). Chaque paire est mise en évidence par une couleur dédiée."
        )
        st.info(info_text)
