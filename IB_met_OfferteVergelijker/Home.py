#!/usr/bin/env python3
"""
Offerte Vergelijker – landing page.
Run locally:  streamlit run IB_met_OfferteVergelijker/Home.py
"""

import streamlit as st

from common import (
    card_caption, configure_page, inject_base_style, jumbo_header, leverancier_icon,
    AANNEMER_ICON, KOELING_ICON, SLOOPWERK_ICON, VAN_KEULEN_ICON,
)

configure_page("Offerte Vergelijker", icon="📊")
inject_base_style()
jumbo_header("📊", "Offerte Vergelijker", "Kies hieronder een leverancier om te starten")

st.write("")

col1, col2, col3, col4 = st.columns(4)

with col1:
    with st.container(border=True):
        leverancier_icon(VAN_KEULEN_ICON)
        st.markdown("### Van Keulen")
        card_caption("Vergelijk PDF offerte met NETTO prijslijst en IB Budget")
        st.page_link("pages/1_🟡_Van_Keulen.py", label="Openen", icon="➡️")

with col2:
    with st.container(border=True):
        leverancier_icon(KOELING_ICON)
        st.markdown("### Koeling")
        card_caption("Vergelijk offertes voor de leverancier Koeling")
        st.page_link("pages/2_🧊_Koeling.py", label="Openen", icon="➡️")

with col3:
    with st.container(border=True):
        leverancier_icon(AANNEMER_ICON)
        st.markdown("### Aannemer")
        card_caption("Vergelijk offertes voor de leverancier Aannemer")
        st.page_link("pages/3_🧱_Aannemer.py", label="Openen", icon="➡️")

with col4:
    with st.container(border=True):
        leverancier_icon(SLOOPWERK_ICON)
        st.markdown("### Sloopwerk")
        card_caption("Vergelijk offertes voor de leverancier Sloopwerk")
        st.page_link("pages/4_🚧_Sloopwerk.py", label="Openen", icon="➡️")
