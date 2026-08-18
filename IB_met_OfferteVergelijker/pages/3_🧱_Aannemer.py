#!/usr/bin/env python3
"""
Offerte Vergelijker – Aannemer section.
Run via the app entrypoint:  streamlit run IB_met_OfferteVergelijker/Home.py

Compares a contractor's PDF offerte (e.g. Van Wijnen 'Afbouw', structurally a
printed Excel bidsheet — same layout as Koeling's 'Offertetemplate KT') against
the IB macro-werkboek's Budget sheet. Unlike Koeling, the Budget rows for this
section have no numeric-prefix bridge to the offerte's own Art.Nr.Jumbo, so
matching is description-based (fuzzy, with a persistent CSV override) — see
matching_aannemer.py for the details.
"""

import pandas as pd
import streamlit as st

from common import AANNEMER_ICON, back_to_overview, configure_page, inject_base_style, jumbo_header
from matching_aannemer import (
    HAS_RAPIDFUZZ,
    build_aannemer_df,
    load_art_nr_mapping,
    load_budget_df,
    match_aannemer_budget,
    parse_pdf,
)

configure_page("Offerte Vergelijker – Aannemer", icon=AANNEMER_ICON)
inject_base_style()


def _style_aannemer(df: pd.DataFrame):
    status = df["_status"].reset_index(drop=True)
    hidden = [c for c in df.columns if c.startswith("_")] + ["Match", "Prijs verschil (€)"]
    display = df.drop(columns=hidden).rename(columns={"Prijs verschil (%)": "Prijs verschil"})

    def _status_cell(target_status: str, css: str):
        def styler(col):
            return [css if status.iloc[i] == target_status else "" for i in range(len(col))]
        return styler

    money = lambda x: f"€ {x:,.2f}" if isinstance(x, (int, float)) else ""
    num = lambda x: f"{x:g}" if isinstance(x, (int, float)) else ""

    def prijs_arrow(x):
        if not isinstance(x, (int, float)):
            return ""
        if x > 0.05:
            return f"▲ {x:.1f}%"
        if x < -0.05:
            return f"▼ {abs(x):.1f}%"
        return ""

    return (
        display.style
        .apply(_status_cell("unmatched", "background-color:#FFC7CE; font-weight:600"), subset=["IB omschrijving"])
        .apply(_status_cell("aantal_afwijking", "background-color:#FFCC00"),
               subset=["Offerte Aantal", "IB Aantal", "Aantal verschil"])
        .format({
            "Offerte Prijs p.e.": money, "IB Prijs p.e.": money,
            "Prijs verschil": prijs_arrow,
            "Offerte Aantal": num, "IB Aantal": num, "Aantal verschil": num,
        }, na_rep="")
    )


def _analyze_aannemer(ib_file, pdf_file):
    ib_bytes = ib_file.read()
    pdf_bytes = pdf_file.read()

    with st.spinner("Analyseren…"):
        offerte = parse_pdf(pdf_bytes)
        mapping = load_art_nr_mapping("vanwijnen")
        bdf = load_budget_df(ib_bytes)
        matches = match_aannemer_budget(offerte, bdf, mapping)
        df = build_aannemer_df(offerte, matches)

    st.session_state.aannemer_df = df
    st.rerun()


def _show_aannemer_results(df: pd.DataFrame):
    df = df[df["Offerte Aantal"].notna()].reset_index(drop=True)
    nrfp_df = df[df["Pos"] == "NRFP"]
    df = df[df["Pos"] != "NRFP"].reset_index(drop=True)

    total = len(df)
    handmatig = int((df["Match"] == "handmatig").sum())
    fuzzy = int((df["Match"] == "fuzzy").sum())
    unmatched = int((df["_status"] == "unmatched").sum())
    prijs_afw = int((df["_status"] == "prijs_afwijking").sum())
    aantal_afw = int((df["_status"] == "aantal_afwijking").sum())

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Offerte regels", total)
    c2.metric("Handmatig", handmatig, help="Gematcht via Art.Nr.Jumbo → Budget-omschrijving mapping")
    c3.metric("Fuzzy", fuzzy, help="Gematcht via tekst-vergelijking")
    c4.metric("Niet gematcht", unmatched)
    c5.metric("Prijs ≠", prijs_afw)
    c6.metric("Aantal ≠", aantal_afw)

    # Samenvatting first — the tab a user should land on right after
    # Analyseren, per project convention (see CLAUDE.md).
    tab_sam, tab_all, tab_fuzzy, tab_exp = st.tabs(
        ["📊 Samenvatting", "📋 Alle resultaten", "🔎 Fuzzy matches (controleren)", "💾 Export"]
    )

    with tab_sam:
        offerte_totaal = float(df["_offerte_totaal"].sum(skipna=True))
        ib_totaal = float(df["_ib_totaal"].sum(skipna=True))
        diff = offerte_totaal - ib_totaal
        diff_pct = diff / ib_totaal * 100 if ib_totaal else 0

        k1, k2, k3 = st.columns(3)
        k1.metric("Totaal Offerte", f"€ {offerte_totaal:,.2f}",
                   help="Prijs per eenheid (arbeid & materiaal) × Offerte Aantal, alleen regels met een ingevuld aantal")
        k2.metric("Totaal IB", f"€ {ib_totaal:,.2f}", help="IB Prijs p.e. × Offerte Aantal, voor dezelfde regels")
        diff_label = f"{'▲' if diff > 0 else '▼'} € {abs(diff):,.2f}  ({diff_pct:+.1f}%)"
        k3.metric("Verschil", diff_label, delta_color="inverse" if diff > 0 else "normal")

        st.divider()

        nrfp_totaal = float(nrfp_df["_offerte_totaal"].sum(skipna=True))
        st.subheader("NRFP — posities niet in de RFP")
        st.caption(
            "Extra/wijzigingsposities buiten de standaard tender (bijv. huur rijplaten, precariokosten) — "
            "deze worden niet tegen het IB-budget gematcht (er is geen 1-op-1 Budget-regel voor), maar tellen "
            "wel mee in de offerte's eigen totaalprijs."
        )
        n1, n2 = st.columns(2)
        n1.metric("NRFP totaal", f"€ {nrfp_totaal:,.2f}")
        n2.metric("Totaal Offerte incl. NRFP", f"€ {offerte_totaal + nrfp_totaal:,.2f}")

        st.divider()

        st.info(
            "📊 Totaal per categorie (Pos 1-10) volgt zodra de koppeling tussen de offerte's Pos-categorieën "
            "en de bijbehorende Budget-subtotalen (A.02.xx) is bevestigd."
        )

        st.divider()

        st.subheader("Match-methode overzicht")
        meth_df = df["Match"].value_counts().reset_index()
        meth_df.columns = ["Methode", "Aantal regels"]
        st.dataframe(meth_df, use_container_width=True, hide_index=True)

    with tab_all:
        st.caption(
            f"{total} offerte-regels — rood = geen IB-match, geel = aantalverschil > 10%, "
            "▲/▼ in Prijs verschil = prijsafwijking"
        )
        st.dataframe(_style_aannemer(df), use_container_width=True, height=600)

    with tab_fuzzy:
        fdf = df[df["Match"] == "fuzzy"].reset_index(drop=True)
        if fdf.empty:
            st.success("Alle matches zijn handmatig bevestigd via de Art.Nr.Jumbo → Budget mapping.")
        else:
            st.caption(
                f"{len(fdf)} regels zijn via fuzzy tekst-matching gekoppeld — controleer deze en voeg zo nodig "
                "een bevestigd paar toe aan data/aannemer_art_nr_mapping_vanwijnen.csv voor een zekere match."
            )
            st.dataframe(_style_aannemer(fdf), use_container_width=True, height=500)

    with tab_exp:
        csv = df.drop(columns=[c for c in df.columns if c.startswith("_")]).to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button(
            "📥 Download vergelijking (CSV)", data=csv,
            file_name="aannemer_vergelijking.csv", mime="text/csv",
        )


back_to_overview()
jumbo_header(AANNEMER_ICON, "Offerte Vergelijker", "Aannemer")

if not HAS_RAPIDFUZZ:
    st.warning("rapidfuzz is niet geïnstalleerd — fuzzy tekst-matching is uitgeschakeld.")

if "aannemer_df" not in st.session_state:
    st.session_state.aannemer_df = None

with st.sidebar:
    st.header("📂 Bestanden uploaden")
    ib_file = st.file_uploader(
        "IB macro-werkboek (.xlsm)", type=["xlsm"],
        help="Het macro-werkboek met 'IB' in de bestandsnaam, bevat de sheet 'Budget'.",
    )
    pdf_file = st.file_uploader(
        "Offerte (.pdf)", type=["pdf"],
        help="De PDF-offerte van de aannemer (bijv. Van Wijnen Afbouw bidsheet).",
    )
    st.caption(
        "Matching gebeurt via een handmatig bevestigde Art.Nr.Jumbo → Budget-omschrijving mapping, met "
        "fuzzy tekst-matching als terugval voor de rest."
    )
    if st.button("▶ Analyseren", type="primary", use_container_width=True):
        if not ib_file or not pdf_file:
            st.error("Upload zowel het **IB macro-werkboek** als de **PDF-offerte**.")
        else:
            _analyze_aannemer(ib_file, pdf_file)

if st.session_state.aannemer_df is None:
    st.info("Upload beide bestanden via de zijbalk en klik op **Analyseren** om te beginnen.")
else:
    _show_aannemer_results(st.session_state.aannemer_df)
