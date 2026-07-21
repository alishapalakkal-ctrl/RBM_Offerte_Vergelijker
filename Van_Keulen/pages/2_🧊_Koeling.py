#!/usr/bin/env python3
"""
Offerte Vergelijker – Koeling section (placeholder).
Run via the app entrypoint:  streamlit run Van_Keulen/Home.py
"""

import streamlit as st

from common import configure_page, inject_base_style, jumbo_header, back_to_overview

configure_page("Offerte Vergelijker – Koeling", icon="🧊")
inject_base_style()

if "koeling_leverancier" not in st.session_state:
    st.session_state.koeling_leverancier = None


@st.dialog("Kies leverancier")
def _choose_leverancier():
    st.write("Voor welke leverancier wil je offertes vergelijken?")
    c1, c2 = st.columns(2)
    if c1.button("Carrier", use_container_width=True):
        st.session_state.koeling_leverancier = "Carrier"
        st.rerun()
    if c2.button("Kaplanlaar", use_container_width=True):
        st.session_state.koeling_leverancier = "Kaplanlaar"
        st.rerun()


back_to_overview()

if st.session_state.koeling_leverancier is None:
    jumbo_header("🧊", "Offerte Vergelijker", "Koeling — kies een leverancier om te starten")
    _choose_leverancier()
    st.stop()

leverancier = st.session_state.koeling_leverancier
jumbo_header("🧊", "Offerte Vergelijker", f"Koeling — {leverancier}")

_, wc = st.columns([5, 1])
with wc:
    if st.button("🔁 Wissel leverancier", use_container_width=True):
        st.session_state.koeling_leverancier = None
        st.rerun()

st.info(
    f"Deze sectie ({leverancier}) is nog in opbouw. Zodra er voorbeeldbestanden "
    "(PDF offerte / NETTO prijslijst / IB Budget) beschikbaar zijn, wordt hier "
    "dezelfde vergelijkingslogica als bij Van Keulen toegevoegd."
)

with st.sidebar:
    st.header("📂 Bestanden uploaden")
    st.file_uploader("PDF Offerte",      type=["pdf"],  disabled=True)
    st.file_uploader("NETTO Prijslijst", type=["xlsx"], disabled=True)
    st.file_uploader("Budget (IB)",      type=["xlsm", "xlsx"], disabled=True)
    st.caption("Uploaden is uitgeschakeld totdat deze sectie is opgebouwd.")
