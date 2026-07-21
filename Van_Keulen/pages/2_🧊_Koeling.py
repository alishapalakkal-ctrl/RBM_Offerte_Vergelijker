#!/usr/bin/env python3
"""
Offerte Vergelijker – Koeling section (placeholder).
Run via the app entrypoint:  streamlit run Van_Keulen/Home.py
"""

import streamlit as st

from common import configure_page, inject_base_style, jumbo_header, back_to_overview

configure_page("Offerte Vergelijker – Koeling", icon="🧊")
inject_base_style()

back_to_overview()
jumbo_header("🧊", "Offerte Vergelijker", "Koeling — vergelijk PDF offerte met NETTO prijslijst en IB Budget")

st.info(
    "Deze sectie is nog in opbouw. Zodra er voorbeeldbestanden (PDF offerte / "
    "NETTO prijslijst / IB Budget) van de leverancier Koeling beschikbaar zijn, "
    "wordt hier dezelfde vergelijkingslogica als bij Van Keulen toegevoegd."
)

with st.sidebar:
    st.header("📂 Bestanden uploaden")
    st.file_uploader("PDF Offerte",      type=["pdf"],  disabled=True)
    st.file_uploader("NETTO Prijslijst", type=["xlsx"], disabled=True)
    st.file_uploader("Budget (IB)",      type=["xlsm", "xlsx"], disabled=True)
    st.caption("Uploaden is uitgeschakeld totdat deze sectie is opgebouwd.")
