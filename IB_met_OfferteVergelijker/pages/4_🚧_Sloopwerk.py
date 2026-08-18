#!/usr/bin/env python3
"""
Offerte Vergelijker – Sloopwerk section.
Run via the app entrypoint:  streamlit run IB_met_OfferteVergelijker/Home.py

Shell only — matching logic (offerte format, IB Budget row range, join
strategy) is not implemented yet; will be added once the leverancier's
offerte/IB layout is known, following the matching_<section>.py + page
pattern already used for Van Keulen/Koeling/Aannemer.
"""

import streamlit as st

from common import SLOOPWERK_ICON, back_to_overview, configure_page, inject_base_style, jumbo_header

configure_page("Offerte Vergelijker – Sloopwerk", icon=SLOOPWERK_ICON)
inject_base_style()

back_to_overview()
jumbo_header(SLOOPWERK_ICON, "Offerte Vergelijker", "Sloopwerk")

with st.sidebar:
    st.header("📂 Bestanden uploaden")
    st.file_uploader("IB macro-werkboek (.xlsm)", type=["xlsm"])
    st.file_uploader("Offerte")
    st.button("▶ Analyseren", type="primary", use_container_width=True, disabled=True)

st.info("Nog niet geïmplementeerd — matching-logica volgt.")
