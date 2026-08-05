#!/usr/bin/env python3
"""
Offerte Vergelijker – Koeling section.
Run via the app entrypoint:  streamlit run Van_Keulen/Home.py

Unlike Van Keulen, the matching itself happens in Dynamo (Revit model
position numbers matched to elements via bounding-box containment, see
02_Positienummer Match.dyn). This page just visualizes what that graph
exports:
  - a single combined Excel file (3 tabs: Vergelijking, Posities niet in
    offerte, Accessoires) from the 'Compare & Export Koeling' node
  - the colored layout PDF (green/red per position) from the graph's
    'Export Koeling View to PDF' node
"""

import base64

import pandas as pd
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

with st.sidebar:
    st.header("📂 Bestanden uploaden")
    st.caption("Beide zijn exports van de Dynamo Positienummer Match graph.")
    excel_file = st.file_uploader(
        "Vergelijking (Excel, 3 tabbladen)",
        type=["xlsx"],
        help="Export van de 'Compare & Export Koeling' node — bevat Vergelijking, "
             "Posities niet in offerte en Accessoires in één bestand.",
    )
    layout_pdf_file = st.file_uploader("Layout (PDF, met positienummers gekleurd)", type=["pdf"])

sheets = pd.read_excel(excel_file, sheet_name=None) if excel_file else {}
uploaded = {
    "Vergelijking": (excel_file, sheets.get("Vergelijking")),
    "Posities niet in offerte": (excel_file, sheets.get("Posities niet in offerte")),
    "Accessoires": (excel_file, sheets.get("Accessoires")),
}

if excel_file is None and layout_pdf_file is None:
    st.info(
        "Upload de vergelijkings-Excel en/of de layout-PDF via de zijbalk — beide zijn "
        "exports van de Dynamo Positienummer Match graph (zie 00_RBM_Offerte Vergelijker/Koeling)."
    )
    st.stop()

tab_layout, tab_compare, tab_extra, tab_accessories = st.tabs(
    ["🗺️ Layout", "📋 Vergelijking", "➕ Posities niet in offerte", "🪟 Accessoires (Glass/Mirror)"]
)

with tab_layout:
    if layout_pdf_file is None:
        st.info("Upload de layout-PDF om deze tab te zien (export vanuit de 'Export Koeling View to PDF' Dynamo-node).")
    else:
        pdf_bytes = layout_pdf_file.getvalue()
        b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
        st.caption("Groen = code en lengte kloppen met de offerte, rood = afwijking.")
        st.markdown(
            f"""
            <iframe src="data:application/pdf;base64,{b64_pdf}"
                    width="100%" height="900" style="border:none;"></iframe>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            "📥 Download layout (PDF)",
            data=pdf_bytes,
            file_name=layout_pdf_file.name,
            mime="application/pdf",
        )

with tab_compare:
    df = uploaded["Vergelijking"][1]
    if df is None:
        st.info("Upload de vergelijkings-Excel om deze tab te zien.")
    else:
        total = len(df)
        code_mismatch = int((df["code_match"] == False).sum())
        length_out = int((df["length_ok"] == False).sum())
        no_length = int(df["revit_length_mm"].isna().sum())

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Posities", total)
        c2.metric("Code afwijkend", code_mismatch)
        c3.metric("Lengte buiten tolerantie", length_out, help="Tolerantie: 80 mm")
        c4.metric("Geen lengte data", no_length)

        def _style_row(row):
            if row.get("code_match") is False or row.get("length_ok") is False:
                return ["background-color:#FFC7CE"] * len(row)
            if pd.isna(row.get("revit_length_mm")):
                return ["background-color:#FFCC00"] * len(row)
            return [""] * len(row)

        st.caption(f"{total} posities vergeleken — rood = code/lengte afwijkend, geel = geen lengte data")
        st.dataframe(df.style.apply(_style_row, axis=1), use_container_width=True, height=600)

        st.download_button(
            "📥 Download vergelijking (CSV)",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name="koeling_vergelijking.csv",
            mime="text/csv",
        )

with tab_extra:
    df_extra = uploaded["Posities niet in offerte"][1]
    if df_extra is None:
        st.info("Upload de vergelijkings-Excel om deze tab te zien.")
    elif df_extra.empty:
        st.success("Geen extra posities gevonden — alle posities in het model staan ook in de offerte.")
    else:
        st.caption(
            f"{len(df_extra)} positie(s) gevonden in het Revit-model die niet in de offerte voorkomen "
            "(mogelijk door een andere leverancier gedekt, of ontbrekend in de offerte)."
        )
        st.dataframe(
            df_extra.rename(columns={
                "pos": "Positienummer",
                "manual_code": "Manual code",
                "description": "Omschrijving",
                "revit_length_mm": "Lengte (mm)",
            }),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "📥 Download posities niet in offerte (CSV)",
            data=df_extra.to_csv(index=False).encode("utf-8-sig"),
            file_name="koeling_extra_posities.csv",
            mime="text/csv",
        )

with tab_accessories:
    df_acc = uploaded["Accessoires"][1]
    if df_acc is None:
        st.info("Upload de vergelijkings-Excel om deze tab te zien.")
    else:
        total = len(df_acc)
        mismatches = int(((df_acc["glass_match"] == False) | (df_acc["mirror_match"] == False)).sum())

        c1, c2 = st.columns(2)
        c1.metric("Groepen (fysieke runs)", total)
        c2.metric("Afwijkend", mismatches)

        def _style_acc(row):
            if row.get("glass_match") is False or row.get("mirror_match") is False:
                return ["background-color:#FFC7CE"] * len(row)
            if isinstance(row.get("source"), str) and "default" in row["source"]:
                return ["background-color:#FFF2CC"] * len(row)
            return [""] * len(row)

        st.caption(
            f"{total} groepen — posities zijn gegroepeerd per fysieke run (opeenvolgende posities met "
            "dezelfde code, gesplitst op offerte-rijgrenzen). Rood = afwijkend, geel = default-aanname "
            "gebruikt (geen REF_Custom data op het element)."
        )
        st.dataframe(df_acc.style.apply(_style_acc, axis=1), use_container_width=True, height=600)

        st.download_button(
            "📥 Download accessoires (CSV)",
            data=df_acc.to_csv(index=False).encode("utf-8-sig"),
            file_name="koeling_accessoires.csv",
            mime="text/csv",
        )
