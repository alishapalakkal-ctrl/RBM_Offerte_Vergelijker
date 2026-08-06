#!/usr/bin/env python3
"""
Offerte Vergelijker – landing page.
Run locally:  streamlit run IB_met_OfferteVergelijker/Home.py
"""

import streamlit as st

from common import configure_page, inject_base_style, jumbo_header, VAN_KEULEN_ICON

configure_page("Offerte Vergelijker", icon="📊")
inject_base_style()
jumbo_header("📊", "Offerte Vergelijker", "Kies hieronder een leverancier om te starten")

st.write("")

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.image(str(VAN_KEULEN_ICON), width=120)
        st.markdown("### Van Keulen")
        st.caption("Vergelijk PDF offerte met NETTO prijslijst en IB Budget")
        st.page_link("pages/1_🟡_Van_Keulen.py", label="Openen", icon="➡️")

with col2:
    with st.container(border=True):
        st.markdown("### 🧊 Koeling")
        st.caption("Vergelijk offertes voor de leverancier Koeling")
        st.page_link("pages/2_🧊_Koeling.py", label="Openen", icon="➡️")
