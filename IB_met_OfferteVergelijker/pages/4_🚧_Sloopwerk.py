#!/usr/bin/env python3
"""
Offerte Vergelijker – Sloopwerk section.
Run via the app entrypoint:  streamlit run IB_met_OfferteVergelijker/Home.py

Compares a demolition contractor's offerte (e.g. Van de Laar 'vertaling
bidsheet' — the Excel workbook, or a PDF printed from it) against the IB
macro-werkboek's Budget sheet,
section A.01 (Sloopwerk). Matching is description-based (fuzzy), restricted
per Groep/categorie — see matching_sloopwerk.py for the details.
"""

import pandas as pd
import streamlit as st

from common import (
    SLOOPWERK_ICON,
    back_to_overview,
    configure_page,
    inject_base_style,
    jumbo_header,
    show_alle_resultaten,
)
from matching_sloopwerk import (
    BUDGET_SLOOPWERK_VANDELAAR_ROW_END,
    BUDGET_SLOOPWERK_VANDELAAR_ROW_START,
    HAS_RAPIDFUZZ,
    build_categorie_totalen,
    build_sloopwerk_df,
    build_sloopwerk_vergelijking,
    get_sloopwerk_totaal_ib,
    load_budget_df,
    load_categorie_budget_codes,
    match_sloopwerk_budget,
    read_offerte_file,
)

configure_page("Offerte Vergelijker – Sloopwerk", icon=SLOOPWERK_ICON)
inject_base_style()


def _style_sloopwerk(df: pd.DataFrame):
    status = df["_status"].reset_index(drop=True)
    hidden = [c for c in df.columns if c.startswith("_")] + ["Match", "Totaal verschil (€)"]
    display = df.drop(columns=hidden).rename(columns={"Totaal verschil (%)": "Totaal verschil"})

    def _status_cell(target_status: str, css: str):
        def styler(col):
            return [css if status.iloc[i] == target_status else "" for i in range(len(col))]
        return styler

    money = lambda x: f"€ {x:,.2f}" if isinstance(x, (int, float)) else ""
    num = lambda x: f"{x:g}" if isinstance(x, (int, float)) else ""

    def verschil_arrow(x):
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
            "Offerte Totaal": money, "IB Totaal": money,
            "Totaal verschil": verschil_arrow,
            "Offerte Aantal": num, "IB Aantal": num, "Aantal verschil": num,
        }, na_rep="")
    )


def _analyze_sloopwerk(ib_file, offerte_file):
    ib_bytes = ib_file.read()
    offerte_bytes = offerte_file.read()

    with st.spinner("Analyseren…"):
        offerte, leverancier = read_offerte_file(offerte_bytes, offerte_file.name)
        bdf = load_budget_df(ib_bytes)
        categorie_mapping = load_categorie_budget_codes()
        matches = match_sloopwerk_budget(offerte, bdf, categorie_mapping)
        df = build_sloopwerk_df(offerte, matches)
        categorie_totalen = build_categorie_totalen(
            bdf, categorie_mapping,
            row_start=BUDGET_SLOOPWERK_VANDELAAR_ROW_START, row_end=BUDGET_SLOOPWERK_VANDELAAR_ROW_END,
        )
        totaal_ib_row = get_sloopwerk_totaal_ib(bdf)

    st.session_state.sloopwerk_df = df
    st.session_state.sloopwerk_vergelijking = build_sloopwerk_vergelijking(
        offerte, matches, categorie_totalen, naam=leverancier, bestand=offerte_file.name,
        bdf=bdf, categorie_mapping=categorie_mapping,
    )
    st.session_state.sloopwerk_categorie_totalen = categorie_totalen
    st.session_state.sloopwerk_totaal_ib_row = totaal_ib_row
    st.rerun()


def _show_sloopwerk_results(df: pd.DataFrame, categorie_totalen=None, totaal_ib_row=None, vergelijking=None):
    categorie_totalen = categorie_totalen or {}
    df = df[df["Offerte Aantal"].notna()].reset_index(drop=True)
    nrfp_df = df[df["Groep"] == "NRFP"]
    df = df[df["Groep"] != "NRFP"].reset_index(drop=True)

    total = len(df)
    unmatched = int((df["_status"] == "unmatched").sum())
    prijs_afw = int((df["_status"] == "prijs_afwijking").sum())
    aantal_afw = int((df["_status"] == "aantal_afwijking").sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Offerte regels", total)
    c2.metric("Niet gematcht", unmatched)
    c3.metric("Prijs ≠", prijs_afw)
    c4.metric("Aantal ≠", aantal_afw)

    # Samenvatting first — the tab a user should land on right after
    # Analyseren, per project convention (see CLAUDE.md).
    tab_sam, tab_all, tab_fuzzy, tab_exp = st.tabs(
        ["📊 Samenvatting", "📋 Alle resultaten", "🔎 Fuzzy matches (controleren)", "💾 Export"]
    )

    with tab_sam:
        offerte_totaal = float(df["_offerte_totaal"].sum(skipna=True))

        cat_df = (
            df.groupby("Groep", dropna=False)[["_offerte_totaal", "_ib_totaal"]]
            .sum(min_count=1)
            .reset_index()
            .rename(columns={"Groep": "Categorie", "_offerte_totaal": "Totaal Offerte", "_ib_totaal": "Totaal IB"})
        )
        if categorie_totalen:
            cat_df["Totaal IB"] = cat_df.apply(
                lambda r: categorie_totalen.get(r["Categorie"], r["Totaal IB"]), axis=1
            )
        cat_df["Verschil"] = cat_df["Totaal Offerte"] - cat_df["Totaal IB"]
        cat_df = cat_df.sort_values("Totaal Offerte", ascending=False, na_position="last")

        # Sum of every A.01.0X subheading's own Totaal on the Budget sheet
        # (all 16 — including any without a matching offerte Groep), falling
        # back to the per-categorie rollup only if the Budget sheet itself
        # couldn't be read.
        ib_totaal = totaal_ib_row if totaal_ib_row is not None else float(cat_df["Totaal IB"].sum(skipna=True))
        diff = offerte_totaal - ib_totaal
        diff_pct = diff / ib_totaal * 100 if ib_totaal else 0

        k1, k2, k3 = st.columns(3)
        k1.metric("Totaal Offerte", f"€ {offerte_totaal:,.2f}",
                   help="Som van 'Waarde offerte Totaal', alleen regels met een ingevuld aantal")
        k2.metric("Totaal IB", f"€ {ib_totaal:,.2f}",
                   help="Som van alle A.01.0X subtotalen op het Budget-tabblad")
        diff_label = f"{'▲' if diff > 0 else '▼'} € {abs(diff):,.2f}  ({diff_pct:+.1f}%)"
        k3.metric("Verschil", diff_label, delta_color="inverse" if diff > 0 else "normal")

        st.divider()

        nrfp_totaal = float(nrfp_df["_offerte_totaal"].sum(skipna=True))
        st.subheader("NRFP — posities niet in de RFP")
        st.caption(
            "Extra/wijzigingsposities buiten de standaard tender — deze worden niet tegen het IB-budget "
            "gematcht (er is geen 1-op-1 Budget-regel voor), maar tellen wel mee in de offerte's eigen totaalprijs."
        )
        n1, n2 = st.columns(2)
        n1.metric("NRFP totaal", f"€ {nrfp_totaal:,.2f}")
        n2.metric("Totaal Offerte incl. NRFP", f"€ {offerte_totaal + nrfp_totaal:,.2f}")

        st.divider()

        st.subheader("Totaal per categorie")
        st.caption(
            "Totaal IB komt, waar bekend, rechtstreeks uit de bijbehorende subtotaal-regel op het "
            "Budget-tabblad — zie data/sloopwerk_categorie_budget_codes.csv."
        )
        st.dataframe(
            cat_df.style.format({
                "Totaal Offerte": lambda x: f"€ {x:,.2f}" if pd.notna(x) else "",
                "Totaal IB": lambda x: f"€ {x:,.2f}" if pd.notna(x) else "",
                "Verschil": lambda x: f"€ {x:,.2f}" if pd.notna(x) else "",
            }),
            use_container_width=True, hide_index=True,
        )

    with tab_all:
        if vergelijking is None:  # results from before this layout existed
            st.info("Klik opnieuw op **Analyseren** om deze weergave te zien.")
        else:
            show_alle_resultaten(vergelijking, key="sloopwerk")

    with tab_fuzzy:
        fdf = df[df["Match"] == "fuzzy"].reset_index(drop=True)
        if fdf.empty:
            st.info("Nog geen fuzzy matches.")
        else:
            st.caption(
                f"{len(fdf)} regels zijn via fuzzy tekst-matching gekoppeld binnen dezelfde categorie — "
                "controleer deze."
            )
            st.dataframe(_style_sloopwerk(fdf), use_container_width=True, height=500)

    with tab_exp:
        csv = df.drop(columns=[c for c in df.columns if c.startswith("_")]).to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button(
            "📥 Download vergelijking (CSV)", data=csv,
            file_name="sloopwerk_vergelijking.csv", mime="text/csv",
        )


back_to_overview()
jumbo_header(SLOOPWERK_ICON, "Offerte Vergelijker", "Sloopwerk")

if not HAS_RAPIDFUZZ:
    st.warning("rapidfuzz is niet geïnstalleerd — fuzzy tekst-matching is uitgeschakeld.")

if "sloopwerk_df" not in st.session_state:
    st.session_state.sloopwerk_df = None
    st.session_state.sloopwerk_categorie_totalen = {}
    st.session_state.sloopwerk_totaal_ib_row = None

with st.sidebar:
    st.header("📂 Bestanden uploaden")
    ib_file = st.file_uploader(
        "IB macro-werkboek (.xlsm)", type=["xlsm"],
        help="Het macro-werkboek met 'IB' in de bestandsnaam, bevat de sheet 'Budget'.",
    )
    offerte_file = st.file_uploader(
        "Offerte (Excel of .pdf)", type=["xlsm", "xlsx", "pdf"],
        help="De 'vertaling bidsheet' offerte van de sloopaannemer (bijv. Van de Laar) — als Excel, of als PDF "
             "afgedrukt uit hetzelfde template.",
    )
    st.caption("Matching gebeurt via fuzzy tekst-matching, per categorie (Groep) beperkt tot dezelfde IB-rubriek.")
    if st.button("▶ Analyseren", type="primary", use_container_width=True):
        if not ib_file or not offerte_file:
            st.error("Upload zowel het **IB macro-werkboek** als de **offerte**.")
        else:
            _analyze_sloopwerk(ib_file, offerte_file)

if st.session_state.sloopwerk_df is None:
    st.info("Upload beide bestanden via de zijbalk en klik op **Analyseren** om te beginnen.")
else:
    _show_sloopwerk_results(
        st.session_state.sloopwerk_df,
        st.session_state.sloopwerk_categorie_totalen,
        st.session_state.sloopwerk_totaal_ib_row,
        st.session_state.get("sloopwerk_vergelijking"),
    )
