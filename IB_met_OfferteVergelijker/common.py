"""Shared page-config, styling and header helpers for the Offerte Vergelijker app."""

import base64
from pathlib import Path

import streamlit as st

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
