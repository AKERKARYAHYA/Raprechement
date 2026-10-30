# -*- coding: utf-8 -*-
"""Suite de Rapprochement Bancaire - Point d'entrée principal.
Lancer : streamlit run app.py
"""
import streamlit as st
import os
from datetime import datetime

# ==========================================
# 🔒 نظام الحماية والفترة التجريبية
# ==========================================
DATE_EXPIRATION = datetime(2026, 12, 30)  # تحديد تاريخ انتهاء الصلاحية
CODE_ACTIVATION_SECRET = "COPAG-2026"  # كود التفعيل الخاص بك

def verifier_securite():
    date_aujourdhui = datetime.now()
    if date_aujourdhui > DATE_EXPIRATION and not st.session_state.get("active", False):
        st.error("🔒 انتهت الفترة التجريبية لهذا التطبيق!")
        st.info("يرجى التواصل مع المسؤول للحصول على كود التفعيل لتمديد الصلاحية.")
        
        code_saisi = st.text_input("أدخل كود التفعيل (Clé d'activation):", type="password")
        
        if st.button("تفعيل التطبيق"):
            if code_saisi == CODE_ACTIVATION_SECRET:
                st.session_state["active"] = True
                st.success("تم تفعيل التطبيق بنجاح!")
                st.rerun()
            else:
                st.error("كود التفعيل غير صحيح!")
        
        st.stop()

# تشغيل فحص الأمان أولاً
verifier_securite()

# ==========================================
# 1. Configuration Globale
# ==========================================
st.set_page_config(
    page_title="Tableau de bord de trésorerie COPAG",
    page_icon="Copag-logo.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Custom CSS سليم بدون إخفاء الهيدر الأساسي باش يبقى زر السايدبار يبان
st.markdown("""
    <style>
    /* Styling Sidebar - Background Azraq */
    [data-testid="stSidebar"] {
        background-color: #1e3a8a;
        color: #ffffff;
    }

    [data-testid="stSidebar"] span, [data-testid="stSidebar"] p, [data-testid="stSidebar"] div, [data-testid="stSidebar"] label {
        color: #f8fafc !important;
    }

    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

# 3. Ta3rif Pages f st.navigation
pages = {
    "": [
        st.Page("views/0_accueil.py", title="Accueil", icon="🏠", default=True),
        st.Page("views/1_rapprochement_3_phases.py", title="Multi-Banques", icon="🏦", url_path="multi-banques"),
        st.Page("views/2_versement_bmci.py", title="BMCIL", icon="🏦", url_path="bmci"),
        st.Page("views/3_versement_caml.py", title="CAML", icon="📊", url_path="caml"),
        st.Page("views/4_ecarts_erp_rlv.py", title="Écarts ERP ↔ RLV", icon="📊", url_path="ecarts"),
        st.Page("views/5_remise_comparaison.py", title="Remise des chèques", icon="🔁", url_path="remise"),
        st.Page("views/6_Canevas.py", title="Canevas", icon="📋", url_path="canevas"),
    ]
}

# 4. BANNER GLOBALE f l-wast (Main Area) l-foq ga3
banner_path = "COPAG.png"
if os.path.exists(banner_path):
    st.image(banner_path, use_container_width=True)

st.markdown("<hr style='margin: 10px 0px 15px 0px;'>", unsafe_allow_html=True)

# 5. Tachghil l-Navigation 
pg = st.navigation(pages)
pg.run()

# 6. 🖼️ Idafa dyal Copag-remov f t-taht w taħtha l-kitaba f s-sidebar
footer_img_path = "Copag-remov.png"
if os.path.exists(footer_img_path):
    st.sidebar.markdown("<hr style='margin: 30px 0px 10px 0px; border-color: #3b82f6;'>", unsafe_allow_html=True)
    st.sidebar.image(footer_img_path, use_container_width=True)
    st.sidebar.markdown("""
        <p style='font-size: 11px; color: #93c5fd; text-align: center; margin-top: 5px; margin-bottom: 10px;'>
            COPAG • Service Trésorerie
        </p>
    """, unsafe_allow_html=True)
