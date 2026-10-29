import streamlit as st
import pandas as pd

# 1. T-tansi9 l-3am dyal Tableau de Bord w l-Banner Al-Wasat
st.markdown("""
    <style>
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .welcome-banner-center {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        padding: 22px;
        border-radius: 12px;
        border-left: 5px solid #1e3a8a;
        margin-bottom: 25px;
        text-align: center; /* Kitaba wllat f wasat s-safha */
    }
    </style>
""", unsafe_allow_html=True)

# 2. Header Al-Wasat (Bonjour w l-wasf m-centré bla date)
st.markdown("""
    <div class="welcome-banner-center">
        <h2 style='margin: 0; color: #1e293b; font-size: 24px;'>👋 Bonjour, Service Trésorerie</h2>
        <p style='margin: 6px 0 0 0; color: #64748b; font-size: 14px;'>Suivi des rapprochements bancaires et des écarts</p>
    </div>
""", unsafe_allow_html=True)

# 3. KPI Cards (4 Cartes dyal l-2ihsa2iyat)
col1, col2, col3, col4 = st.columns(4, gap="medium")

with col1:
    st.markdown("""
        <div class="metric-card">
            <span style='color: #64748b; font-size: 13px; font-weight: 600;'>Taux de rapprochement</span>
            <h2 style='margin: 8px 0; color: #0f172a;'>94,8%</h2>
            <span style='color: #16a34a; font-size: 12px; font-weight: 500;'>▲ +2,3% vs. période précédente</span>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div class="metric-card">
            <span style='color: #64748b; font-size: 13px; font-weight: 600;'>Écarts à traiter</span>
            <h2 style='margin: 8px 0; color: #0f172a;'>28</h2>
            <span style='color: #16a34a; font-size: 12px; font-weight: 500;'>▼ -12 vs. période précédente</span>
        </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
        <div class="metric-card">
            <span style='color: #64748b; font-size: 13px; font-weight: 600;'>Opérations rapprochées</span>
            <h2 style='margin: 8px 0; color: #0f172a;'>1 248</h2>
            <span style='color: #16a34a; font-size: 12px; font-weight: 500;'>▲ +86 vs. période précédente</span>
        </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
        <div class="metric-card">
            <span style='color: #64748b; font-size: 13px; font-weight: 600;'>Solde bancaire</span>
            <h2 style='margin: 8px 0; color: #0f172a; font-size: 22px;'>2 458 620 MAD</h2>
            <span style='color: #16a34a; font-size: 12px; font-weight: 500;'>▲ +5,2% vs. période précédente</span>
        </div>
    """, unsafe_allow_html=True)

st.write("")
st.write("")

# 4. Les Graphiques (Evolution & Repartition)
c_chart1, c_chart2 = st.columns([0.65, 0.35], gap="medium")

with c_chart1:
    st.markdown("#### 📈 Évolution du rapprochement bancaire")
    chart_data = pd.DataFrame({
        'Taux': [70, 75, 78, 80, 81, 82, 85, 88, 90, 92, 94.8],
        'Écarts': [30, 25, 23, 20, 18, 17, 16, 15, 14, 15, 28]
    }, index=['Janv.', 'Fév.', 'Mars', 'Avr.', 'Mai', 'Juin', 'Juil.', 'Août', 'Sept.', 'Oct.', 'Nov.'])
    st.line_chart(chart_data, color=["#1e3a8a", "#0284c7"], height=250)

with c_chart2:
    st.markdown("#### 📊 Répartition des opérations")
    st.info("💡 Rapprochées: **1 248 (97,8%)**  \n💡 Non rapprochées: **28 (2,2%)**")
    st.progress(0.978, text="Taux global de lettrage")

st.write("")

# 5. Tableau: Derniers Rapprochements
st.markdown("#### 📋 Derniers Rapprochements")

data_table = pd.DataFrame({
    "Banque": ["BMCI", "CAML", "Attijariwafa bank", "BMCI", "CAML"],
    "Date de rapprochement": ["21/11/2025", "20/11/2025", "18/11/2025", "15/11/2025", "12/11/2025"],
    "Nombre d'opérations": [324, 287, 198, 176, 142],
    "Montant total (MAD)": ["1 245 680", "986 420", "756 310", "652 870", "487 230"],
    "Taux": ["96,3%", "93,7%", "91,4%", "98,2%", "89,6%"],
    "Statut": ["✅ Terminé", "✅ Terminé", "⏳ En cours", "✅ Terminé", "⚠️ À traiter"]
})

st.dataframe(data_table, use_container_width=True, hide_index=True)