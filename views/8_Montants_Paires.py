import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment

# Configuration de la page
st.set_page_config(page_title="تراصف المبالغ - Excel Matching Pro", layout="wide")

# --- Custom CSS ---
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

st.title("📊 نظام تنظيم وتطابق المبالغ الاحترافي (Plus / Moins)")
st.write("قم بتحميل ملف Excel، وسيقوم النظام باختيار المبالغ المتطابقة فقط وتلوين كل زوج (Pair) بلون مميز لسهولة المراجعة.")

uploaded_file = st.file_uploader("اختر ملف Excel للعمليات", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    with st.expander("📄 معاينة البيانات الأصلية"):
        st.dataframe(df.head(), use_container_width=True)

    col_names = df.columns.tolist()
    amount_col = st.selectbox("🎯 اختر الخانة (Colonne) اللي فيها les montants:", col_names)

    if st.button("🚀 بدء المعالجة والتنظيم الذكي"):
        data = df.copy()
        data[amount_col] = pd.to_numeric(data[amount_col], errors='coerce').fillna(0)

        rows = data.to_dict('records')
        matched_rows = []
        used_indices = set()

        # Step 1: Matching (Ghir li 3andhoum mokabil)
        for i, row in enumerate(rows):
            if i in used_indices:
                continue
            
            val = row[amount_col]
            if val > 0:
                for j, other_row in enumerate(rows):
                    if j not in used_indices and other_row[amount_col] == -val and j != i:
                        matched_rows.append(row)
                        matched_rows.append(other_row)
                        used_indices.add(i)
                        used_indices.add(j)
                        break

        df_matched = pd.DataFrame(matched_rows)

        st.success("✅ تم استخراج وتنسيق المبالغ المتطابقة بنجاح!")

        # Métriques
        m1, m2 = st.columns(2)
        with m1:
            st.metric(label="📊 إجمالي الصفوف الأصلية", value=len(df))
        with m2:
            st.metric(label="✅ الصفوف المتطابقة فقط (تم استبعاد الباقي)", value=len(df_matched))

        st.divider()

        st.subheader("✅ جدول المبالغ المتطابقة المنظمة")
        st.dataframe(df_matched, use_container_width=True)

        # --- Génération Excel b Coloration Rotationnelle dyal Les Pairs ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Ghandiro ghir feuille wahda dyal Matchs
            df_matched.to_excel(writer, index=False, sheet_name='Matchs_Complets')
            
            workbook = writer.book
            ws = workbook['Matchs_Complets']
            
            # 9awaim d l-alwan l-pastél bach kolla pair yakhood lon
            pair_fills = [
                PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid"), # Orange pastel
                PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"), # Vert pastel
                PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"), # Jaune pastel
                PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"), # Bleu pastel
            ]
            
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            # Styling Headers
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # T-lovin b Kolla Pair (kolla 2 sfoof m-tadbine b lon)
            total_rows = ws.max_row
            color_index = 0
            
            # Nbdaw mn row 2 htal l-akher, nakhdo kolla 2 sfoof (Pair) b zoj
            for r in range(2, total_rows + 1, 2):
                current_fill = pair_fills[color_index % len(pair_fills)]
                
                # Nloyno saf l-awwal w saf t-tani dyal l-pair
                for row_idx in range(r, min(r + 2, total_rows + 1)):
                    for col_num in range(1, ws.max_column + 1):
                        ws.cell(row=row_idx, column=col_num).fill = current_fill
                
                color_index += 1

        processed_data = output.getvalue()

        st.divider()
        st.subheader("⬇️ Export du Classeur Excel")

        st.download_button(
            label="📥 Télécharger le classeur Excel (Pairs Colorés)",
            data=processed_data,
            file_name="Montants_Paires_Colores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_button"
        )

        st.info("💡 معلومة: الملف المحمل يحتوي فقط على المبالغ التي لها مقابل (Pairs)، وكل زوج من المبالغ ملون بلون مختلف لتسهيل الرؤية تماماً مثل الصورة المطلوبة.")
