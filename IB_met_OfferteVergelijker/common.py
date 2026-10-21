"""Shared page-config, styling and header helpers for the Offerte Vergelijker app."""

import base64
from dataclasses import replace
from pathlib import Path

import pandas as pd
import streamlit as st

from vergelijking import OfferteVergelijking, vergelijking_regels_df

ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "jumbo_logo.png"
VAN_KEULEN_ICON = ASSETS_DIR / "vankeulen_icon.png"
AANNEMER_ICON = ASSETS_DIR / "Van Wijnen.png"
KOELING_ICON = ASSETS_DIR / "Frimex.png"
SLOOPWERK_ICON = ASSETS_DIR / "fried-van-de-laar.png"


def _data_uri(path: Path) -> str:
    data = base64.b64encode(path.read_bytes()).decode()
    return f"data:image/png;base64,{data}"


def _is_image(icon) -> bool:
    return isinstance(icon, Path) or (isinstance(icon, str) and icon.lower().endswith((".png", ".jpg", ".jpeg")))


def configure_page(title: str, icon="📊"):
    st.set_page_config(page_title=title, page_icon=str(icon) if _is_image(icon) else icon, layout="wide")


def inject_base_style():
    st.markdown("""
    <style>
      /* Pull the entire main content area up */
      [data-testid="stAppViewBlockContainer"] {
          padding-top: 0.5rem !important;
      }
      .jumbo-hdr {
          background: linear-gradient(135deg, #FDC400 0%, #e8ac00 100%);
          padding: 12px 24px; border-radius: 10px; margin-bottom: 12px;
          display: flex; align-items: center; gap: 18px;
          box-shadow: 0 3px 10px rgba(0,0,0,.15);
      }
      .jumbo-hdr h1 { margin: 0; font-size: 26px; font-weight: 800; color: #1a1a1a; }
      .jumbo-hdr p  { margin: 3px 0 0; font-size: 12px; color: #444; }
      [data-testid="stSidebar"] { background: #f5f5f5; }
      [data-testid="stMetricValue"] { font-size: 28px !important; }
      .legend-row { display:flex; gap:18px; margin: 6px 0 14px; font-size:13px; }
      .legend-chip { padding: 3px 12px; border-radius: 5px; font-weight:600; }

      /* Push modal dialogs (st.dialog) down toward the middle of the screen */
      div[data-testid="stDialog"] {
          align-items: flex-start !important;
          padding-top: 22vh !important;
      }
    </style>
    """, unsafe_allow_html=True)


def jumbo_header(icon, title: str, subtitle: str):
    if _is_image(icon):
        icon_html = f'<img src="{_data_uri(Path(icon))}" style="height:30px;vertical-align:middle;border-radius:3px" />'
    else:
        icon_html = icon
    st.markdown(f"""
    <div class="jumbo-hdr">
        <img src="{_data_uri(LOGO_PATH)}" style="height:48px;border-radius:4px" />
        <div>
            <h1>{icon_html} {title}</h1>
            <p>{subtitle}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def back_to_overview():
    st.page_link("Home.py", label="Terug naar overzicht", icon="⬅️")


def leverancier_icon(path: Path, height: int = 100):
    """Render a leverancier logo at a fixed height regardless of its source
    aspect ratio — plain st.image(..., width=N) renders non-square source
    images (e.g. vankeulen_icon.png at 148x119) shorter than square ones
    (Frimex.png/Van Wijnen.png at 148x148), making Home.py's cards uneven
    heights side by side."""
    st.markdown(
        f'<div style="height:{height}px; display:flex; align-items:center; justify-content:center;">'
        f'<img src="{_data_uri(path)}" style="max-height:{height}px; max-width:100%; object-fit:contain;" />'
        f'</div>',
        unsafe_allow_html=True,
    )


def card_caption(text: str, height: int = 40):
    """Render Home.py's card caption at a fixed min-height — captions of
    different lengths (e.g. Koeling's one-liner vs Aannemer's/Sloopwerk's
    '(bijv. ...)' suffix) wrap to a different number of lines, which makes
    st.container(border=True) cards uneven heights side by side even with
    leverancier_icon() already normalizing the logo above them."""
    st.markdown(
        f'<div style="min-height:{height}px; font-size:0.875rem; color:rgb(120,120,120); '
        f'line-height:1.3;">{text}</div>',
        unsafe_allow_html=True,
    )


#: Light column-group tints for 'Alle resultaten', as in the Excel export's
#: Vergelijking sheet (IB grey, then one colour per offerte).
_GROEP_TINTS = ["#DDEBF7", "#FCE4D6", "#E2EFDA", "#FFF2CC", "#EDE1F5"]


def _verschil_label(diff, offerte_totaal, ib_totaal) -> str:
    """'▲ € 1,234.56 (+5.3%)' / '▼ € 500.00 (-2.1%)' for an offerte − IB
    difference, like the Samenvatting's Verschil KPI. No arrow when the
    offerte didn't quote the line — the difference is then just −IB, not a
    saving."""
    if diff is None or pd.isna(diff):
        return ""
    if abs(diff) < 0.005:
        return "-"
    if not offerte_totaal:
        return f"€ {diff:,.2f}"
    pct = f"  ({diff / ib_totaal * 100:+.1f}%)" if ib_totaal else ""
    return f"{'▲' if diff > 0 else '▼'} € {abs(diff):,.2f}{pct}"


def show_alle_resultaten(v: OfferteVergelijking, key: str):
    """Every page's '📋 Alle resultaten' tab: the offerte(s) next to the IB
    per line, laid out like the Excel export's Vergelijking sheet — column
    tints per party, ▲/▼ in the '− IB' columns, the lowest offerte in green
    (2+ offertes), a Totaal row and a Categorie filter. `key` keeps the
    filter's widget state separate per page."""
    ranked = len(v.namen) > 1
    cat = st.selectbox("Categorie", ["Alle categorieën", *v.categorieen], key=f"{key}_alle_cat")
    if cat != "Alle categorieën":
        v = replace(v, regels=[r for r in v.regels if r.categorie == cat])
    rdf = vergelijking_regels_df(v)
    st.caption(
        f"{len(v.regels)} regels — het IB naast {'elke offerte' if ranked else 'de offerte'}, zelfde opbouw als "
        "het tabblad Vergelijking in de Excel-export. ▲/▼ = offerte boven/onder het IB."
        + (" Groen = laagste offerte op die regel." if ranked else "")
    )
    if rdf.empty:
        return

    groups = [["IB Aantal", "IB Prijs p.e.", "IB Totaal"]] + [
        [f"{n} Aantal", f"{n} Prijs p.e.", f"{n} Totaal", f"{n} − IB"] for n in v.namen]
    tint = {c: "#F2F2F2" for c in groups[0]}
    for i, cols in enumerate(groups[1:]):
        tint.update({c: _GROEP_TINTS[i % len(_GROEP_TINTS)] for c in cols})
    tot_cols = [f"{n} Totaal" for n in v.namen]

    def _row_style(row):
        quoted = {c: row[c] for c in tot_cols if row[c] > 0}
        best = min(quoted, key=quoted.get) if ranked and quoted else None
        is_total = row["Categorie"] == "Totaal"
        styles = []
        for c in row.index:
            css = ""
            if c == best:
                css = "background-color:#C6EFCE; color:#1F1F1F"
            elif c in tint:
                css = f"background-color:{tint[c]}; color:#1F1F1F"
            if is_total:
                css += "; font-weight:700"
            styles.append(css)
        return styles

    # ▲/▼ per line needs the row's own totals, which a Styler.format callable
    # can't see — so the '− IB' columns become display strings up front.
    for n in v.namen:
        rdf[f"{n} − IB"] = [
            _verschil_label(d, o, ib)
            for d, o, ib in zip(rdf[f"{n} − IB"], rdf[f"{n} Totaal"], rdf["IB Totaal"], strict=True)
        ]

    money = lambda x: "" if x is None or pd.isna(x) else ("-" if x == 0 else f"€ {x:,.2f}")
    num = lambda x: "" if x is None or pd.isna(x) else f"{x:g}"
    money_cols = ["IB Prijs p.e.", "IB Totaal", *[f"{n} {k}" for n in v.namen for k in ("Prijs p.e.", "Totaal")]]
    num_cols = ["IB Aantal", *[f"{n} Aantal" for n in v.namen]]
    st.dataframe(
        rdf.style.apply(_row_style, axis=1).format({**{c: money for c in money_cols}, **{c: num for c in num_cols}}),
        use_container_width=True, hide_index=True, height=600,
    )
