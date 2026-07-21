"""Shared page-config, styling and header helpers for the Offerte Vergelijker app."""

import base64
from pathlib import Path

import streamlit as st

_LOGO_PATH = Path(__file__).parent / "assets" / "jumbo_logo.png"


def _logo_data_uri() -> str:
    data = base64.b64encode(_LOGO_PATH.read_bytes()).decode()
    return f"data:image/png;base64,{data}"


def configure_page(title: str, icon: str = "📊"):
    st.set_page_config(page_title=title, page_icon=icon, layout="wide")


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
    </style>
    """, unsafe_allow_html=True)


def jumbo_header(icon: str, title: str, subtitle: str):
    st.markdown(f"""
    <div class="jumbo-hdr">
        <img src="{_logo_data_uri()}" style="height:48px;border-radius:4px" />
        <div>
            <h1>{icon} {title}</h1>
            <p>{subtitle}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def back_to_overview():
    st.page_link("Home.py", label="Terug naar overzicht", icon="⬅️")
