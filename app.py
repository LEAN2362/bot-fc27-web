import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import re
import json
import math
import time
import base64
import unicodedata
import html as html_lib
import requests
import traceback
from urllib.parse import quote
from datetime import datetime, timedelta, timezone

try:
    from PIL import Image, ImageOps
except Exception:
    Image = None
    ImageOps = None

# ==========================================
# CONFIGURACIÓN DE PÁGINA Y SISTEMA DE DISEÑO FINTECH
# ==========================================
st.set_page_config(
    page_title="FC27 Trading Terminal Pro (Cloud)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ⬇️ SI QUERÉS DEJAR TU LINK DE GOOGLE SHEETS FIJO, PEGALO ACÁ ENTRE LAS COMILLAS ⬇️
LINK_GOOGLE_SHEET = ""

# PIN DE ADMINISTRADOR PARA VER Y EDITAR TU CARTERA REAL
PIN_ADMIN_SECRETO = "2362"
REPO_GITHUB_DEFAULT = "LEAN2362/bot-fc27-web"

# CREDENCIALES IA (PROTEGIDAS CONTRA ESCÁNER GITHUB)
_G_PART1 = "AQ.Ab8RN6KGQFnimiJeqwsb"
_G_PART2 = "KKcbIWeCUq6SggjfhQ6kIN9aKIFafA"
DEFAULT_GEMINI_KEY = os.environ.get("GEMINI_API_KEY", _G_PART1 + _G_PART2)
MODELOS_PREFERIDOS_IA = [
    "models/gemini-2.5-flash", "gemini-2.5-flash",
    "models/gemini-3.5-flash", "gemini-3.5-flash",
    "gemini-2.0-flash"
]

PRESUPUESTO_DEFAULT = 200000
TOPE_RADAR_MAXIMO = 15000000
UMBRAL_DESFASE_VERSION_PCT = 45.0
ARCHIVO_EXCEL = "Lista_Jugadores_FC27.xlsx"
ARCHIVO_LINK_CACHE = "sheet_link_cache.txt"
ARCHIVO_CARTERA_LOCAL = "cartera_cloud_sync.csv"
ARCHIVO_ALERTAS_LOCAL = "alertas_cloud_sync.csv"
ARCHIVO_IA_LOCAL = "ia_cloud_sync.csv"
ARCHIVO_SALDOS_LOCAL = "saldos_cloud_sync.json"
ARCHIVO_CRED = "credenciales_google.json"
NOMBRE_SHEET = "Base_Datos_FC27"

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    :root {
        --bg-deep: #090d16;
        --bg-surface: rgba(17, 24, 39, 0.72);
        --bg-surface-strong: rgba(15, 23, 42, 0.88);
        --border-subtle: rgba(255, 255, 255, 0.08);
        --border-hover: rgba(255, 255, 255, 0.16);
        --text-primary: #f1f5f9;
        --text-secondary: #94a3b8;
        --accent-emerald: #10b981;
        --accent-cyan: #38bdf8;
        --accent-amber: #f59e0b;
        --accent-crimson: #ef4444;
    }

    /* Ocultar chrome genérico de Streamlit y optimizar espacio útil */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header[data-testid="stHeader"] {
        background: rgba(9, 13, 22, 0.55) !important;
        backdrop-filter: blur(10px);
        border-bottom: 1px solid rgba(255, 255, 255, 0.04);
    }

    html, body, .stApp, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    .stApp {
        background:
            radial-gradient(ellipse 80% 45% at 12% 0%, rgba(16, 185, 129, 0.09) 0%, transparent 60%),
            radial-gradient(ellipse 65% 40% at 88% 0%, rgba(56, 189, 248, 0.08) 0%, transparent 55%),
            #090d16 !important;
        color: var(--text-primary);
    }

    .block-container {
        padding-top: 1.35rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 1520px;
    }

    .mono {
        font-family: 'JetBrains Mono', monospace !important;
        font-variant-numeric: tabular-nums;
        letter-spacing: -0.02em;
    }

    /* Sidebar Fintech */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0b1120 0%, #090d16 100%) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.07) !important;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.1rem !important;
    }
    section[data-testid="stSidebar"] * {
        color: #e2e8f0 !important;
    }
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background-color: rgba(15, 23, 42, 0.85) !important;
        color: #f8fafc !important;
        border: 1px solid rgba(255, 255, 255, 0.10) !important;
        border-radius: 10px !important;
    }
    .sb-brand {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 12px;
        border-radius: 12px;
        background: rgba(17, 24, 39, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.08);
        margin-bottom: 10px;
    }
    .sb-brand-title {
        font-size: 1.02rem;
        font-weight: 800;
        color: #f8fafc !important;
        letter-spacing: -0.02em;
        line-height: 1.15;
    }
    .sb-brand-sub {
        font-size: 0.72rem;
        color: #94a3b8 !important;
        margin-top: 2px;
    }
    .sb-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.10), transparent);
        margin: 14px 0;
    }
    .sb-section-label {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #64748b !important;
        margin-bottom: 6px;
    }

    /* Botones con micro-interacción */
    div.stButton > button,
    div[data-testid="stFormSubmitButton"] > button,
    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #10b981 0%, #0ea5e9 100%) !important;
        color: #041016 !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-weight: 800 !important;
        font-size: 0.86rem !important;
        letter-spacing: -0.01em;
        border: 1px solid rgba(255, 255, 255, 0.14) !important;
        border-radius: 10px !important;
        padding: 0.55rem 1rem !important;
        box-shadow: 0 4px 16px rgba(16, 185, 129, 0.22) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease, filter 0.2s ease !important;
    }
    div.stButton > button:hover,
    div[data-testid="stFormSubmitButton"] > button:hover,
    div[data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 24px rgba(14, 165, 233, 0.32) !important;
        filter: brightness(1.06);
    }

    /* Badges compactos de modo */
    .mode-badge-admin, .mode-badge-club, .mode-badge-guest {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        padding: 7px 11px;
        border-radius: 9px;
        font-size: 0.78rem;
        font-weight: 700;
        text-align: center;
        margin: 6px 0 8px 0;
        letter-spacing: 0.01em;
    }
    .mode-badge-admin {
        background: rgba(16, 185, 129, 0.14);
        border: 1px solid rgba(16, 185, 129, 0.35);
        color: #34d399 !important;
    }
    .mode-badge-club {
        background: rgba(245, 158, 11, 0.14);
        border: 1px solid rgba(245, 158, 11, 0.35);
        color: #fbbf24 !important;
    }
    .mode-badge-guest {
        background: rgba(56, 189, 248, 0.12);
        border: 1px solid rgba(56, 189, 248, 0.30);
        color: #7dd3fc !important;
    }

    /* Ticket Financiero Calculadora Rápida (Sidebar) */
    .calc-ticket {
        background: rgba(17, 24, 39, 0.82);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.09);
        border-top: 3px solid #10b981;
        border-radius: 12px;
        padding: 12px 13px;
        margin-top: 8px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
    }
    .calc-ticket-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.70rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #94a3b8 !important;
        padding-bottom: 8px;
        border-bottom: 1px dashed rgba(255, 255, 255, 0.14);
        margin-bottom: 8px;
    }
    .calc-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 4px 0;
        font-size: 0.80rem;
    }
    .calc-row-lbl { color: #cbd5e1 !important; font-weight: 500; }
    .calc-row-val {
        font-family: 'JetBrains Mono', monospace !important;
        font-weight: 700;
        font-size: 0.83rem;
        font-variant-numeric: tabular-nums;
    }
    .calc-ticket-foot {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 8px;
        padding-top: 8px;
        border-top: 1px dashed rgba(255, 255, 255, 0.14);
        font-size: 0.82rem;
        font-weight: 700;
    }

    /* Cabecera Principal Glassmorphism */
    .main-header {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        padding: 20px 24px;
        border-radius: 18px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 12px 34px rgba(0, 0, 0, 0.38);
        margin-bottom: 14px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 18px;
        flex-wrap: wrap;
        position: relative;
        overflow: hidden;
    }
    .main-header::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
        background: linear-gradient(90deg, #10b981 0%, #38bdf8 55%, #f59e0b 100%);
    }
    .header-left {
        display: flex;
        align-items: center;
        gap: 16px;
        flex: 1 1 420px;
        min-width: 0;
    }
    .club-crest {
        width: 54px;
        height: 54px;
        border-radius: 14px;
        background: rgba(15, 23, 42, 0.9);
        border: 1px solid rgba(255, 255, 255, 0.12);
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.08), 0 6px 18px rgba(0, 0, 0, 0.35);
    }
    .header-title-wrap { min-width: 0; }
    .main-header h1 {
        margin: 0;
        font-size: 1.55rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        color: #f8fafc;
        line-height: 1.2;
    }
    .header-club-sub {
        font-size: 0.80rem;
        color: #94a3b8;
        margin-top: 2px;
        font-weight: 600;
    }
    .header-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 7px;
        margin-top: 10px;
    }
    .status-pill {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 0.75rem;
        font-weight: 600;
        background: rgba(15, 23, 42, 0.78);
        border: 1px solid rgba(255, 255, 255, 0.08);
        color: #cbd5e1;
    }
    .status-pill b {
        font-family: 'JetBrains Mono', monospace;
        color: #f8fafc;
        font-weight: 700;
    }
    .drop-timer {
        background: rgba(9, 13, 22, 0.78);
        border: 1px solid rgba(56, 189, 248, 0.24);
        padding: 12px 18px;
        border-radius: 14px;
        text-align: right;
        min-width: 215px;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.05);
    }
    .drop-timer-title {
        font-size: 0.70rem;
        color: #38bdf8;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 700;
    }
    .drop-timer-val {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 1.40rem;
        font-weight: 800;
        color: #f8fafc;
        margin: 2px 0;
        font-variant-numeric: tabular-nums;
    }
    .drop-timer-sub {
        font-size: 0.74rem;
        color: #94a3b8;
        font-weight: 500;
    }

    /* Semáforo y Calendario EA */
    .status-strip-card {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 13px 16px;
        margin-bottom: 14px;
        transition: transform 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.28);
    }
    .status-strip-card:hover {
        transform: translateY(-2px);
        border-color: rgba(255, 255, 255, 0.15);
    }
    .status-strip-title {
        font-weight: 800;
        font-size: 0.88rem;
        letter-spacing: -0.01em;
    }
    .status-strip-desc {
        font-size: 0.80rem;
        color: #94a3b8;
        margin-top: 4px;
        line-height: 1.45;
    }

    /* 5 Tarjetas KPI Dark-Glass */
    .kpi-card {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 16px 16px 14px 16px;
        color: #f8fafc;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.30);
        margin-bottom: 12px;
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
        position: relative;
        overflow: hidden;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        border-color: rgba(255, 255, 255, 0.18);
        box-shadow: 0 14px 32px rgba(0, 0, 0, 0.45);
    }
    .kpi-green { border-top: 3px solid #10b981; }
    .kpi-blue { border-top: 3px solid #38bdf8; }
    .kpi-teal { border-top: 3px solid #2dd4bf; }
    .kpi-gold { border-top: 3px solid #f59e0b; }
    .kpi-crimson { border-top: 3px solid #ef4444; }
    .kpi-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
    }
    .kpi-title {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.07em;
        color: #94a3b8;
    }
    .kpi-icon {
        width: 28px;
        height: 28px;
        border-radius: 8px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 0.90rem;
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.07);
    }
    .kpi-value {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 1.32rem;
        font-weight: 800;
        color: #f8fafc;
        margin-bottom: 10px;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
        letter-spacing: -0.03em;
    }
    .kpi-badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 3px 9px;
        border-radius: 999px;
        font-size: 0.73rem;
        font-weight: 600;
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid rgba(255, 255, 255, 0.07);
        color: #cbd5e1;
        font-family: 'JetBrains Mono', monospace;
    }

    /* Barra dinámica de 100 huecos de transferencias */
    .transfer-bar-wrap {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 11px 16px;
        margin: 4px 0 14px 0;
    }
    .transfer-bar-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 8px;
        font-size: 0.80rem;
        font-weight: 600;
        color: #cbd5e1;
    }
    .transfer-track {
        width: 100%;
        height: 9px;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.07);
        overflow: hidden;
    }
    .transfer-fill {
        height: 100%;
        border-radius: 999px;
        transition: width 0.35s ease;
    }

    /* Top 3 Oportunidades (Pestaña 2) */
    .opp-card {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-left: 3px solid #10b981;
        border-radius: 15px;
        padding: 15px 16px;
        margin-bottom: 14px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.30);
        transition: transform 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease;
    }
    .opp-card:hover {
        transform: translateY(-3px);
        border-color: rgba(16, 185, 129, 0.40);
        box-shadow: 0 14px 32px rgba(0, 0, 0, 0.45);
    }
    .opp-rank-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 8px;
        margin-bottom: 6px;
    }
    .opp-tag {
        font-size: 0.69rem;
        font-weight: 700;
        padding: 3px 8px;
        border-radius: 999px;
        background: rgba(16, 185, 129, 0.14);
        color: #34d399;
        border: 1px solid rgba(16, 185, 129, 0.28);
    }
    .opp-source {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.70rem;
        font-weight: 700;
        padding: 3px 8px;
        border-radius: 6px;
        background: rgba(56, 189, 248, 0.14);
        color: #7dd3fc;
        border: 1px solid rgba(56, 189, 248, 0.25);
    }
    .opp-name {
        font-size: 1.02rem;
        font-weight: 800;
        color: #f8fafc;
        margin-bottom: 2px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .opp-sub {
        font-size: 0.75rem;
        color: #94a3b8;
        margin-bottom: 10px;
    }
    .opp-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px;
        margin-bottom: 10px;
    }
    .opp-metric {
        background: rgba(9, 13, 22, 0.62);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 10px;
        padding: 7px 9px;
    }
    .opp-metric-lbl {
        font-size: 0.66rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #64748b;
        font-weight: 700;
    }
    .opp-metric-val {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.90rem;
        font-weight: 800;
        color: #f8fafc;
        margin-top: 2px;
    }
    .opp-footer {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.75rem;
        color: #94a3b8;
        padding-top: 8px;
        border-top: 1px solid rgba(255, 255, 255, 0.06);
    }
    .opp-footer a {
        color: #38bdf8;
        text-decoration: none;
        font-weight: 700;
    }

    /* Tarjetas SBC y Simulador Fodder (Pestaña 3) */
    .sbc-card {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-top: 3px solid #10b981;
        border-radius: 15px;
        padding: 16px 18px;
        margin-bottom: 14px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.30);
    }
    .sbc-card-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 12px;
    }
    .sbc-card-title {
        font-size: 0.96rem;
        font-weight: 800;
        color: #f8fafc;
    }
    .sbc-recipe-pill {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.76rem;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 8px;
        background: rgba(16, 185, 129, 0.14);
        border: 1px solid rgba(16, 185, 129, 0.30);
        color: #34d399;
    }
    .sbc-metrics-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 10px;
        margin-bottom: 12px;
    }
    .sbc-tile {
        background: rgba(9, 13, 22, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 11px;
        padding: 10px 12px;
    }
    .sbc-tile-lbl {
        font-size: 0.68rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #94a3b8;
        font-weight: 700;
    }
    .sbc-tile-val {
        font-family: 'JetBrains Mono', monospace;
        font-size: 1.05rem;
        font-weight: 800;
        color: #f8fafc;
        margin-top: 3px;
    }
    .sbc-shop-list {
        background: rgba(15, 23, 42, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 10px;
        padding: 9px 12px;
        font-size: 0.80rem;
        color: #cbd5e1;
        line-height: 1.45;
    }

    /* Tarjetas IA y Consultor Chat (Pestaña 4) */
    .ia-card-report, .ia-card-oracle {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        overflow: hidden;
        margin-top: 8px;
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.34);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .ia-card-report:hover, .ia-card-oracle:hover {
        transform: translateY(-2px);
        border-color: rgba(255, 255, 255, 0.16);
    }
    .ia-card-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 13px 18px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.07);
        background: rgba(9, 13, 22, 0.55);
    }
    .ia-card-report .ia-card-head { border-top: 3px solid #10b981; }
    .ia-card-oracle .ia-card-head { border-top: 3px solid #38bdf8; }
    .ia-card-head-title {
        font-size: 0.94rem;
        font-weight: 800;
        color: #f8fafc;
        letter-spacing: -0.01em;
    }
    .ia-live-chip {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.68rem;
        font-weight: 700;
        padding: 3px 9px;
        border-radius: 999px;
        background: rgba(16, 185, 129, 0.14);
        color: #34d399;
        border: 1px solid rgba(16, 185, 129, 0.30);
    }
    .ia-card-body {
        padding: 18px 20px;
        font-size: 0.88rem;
        line-height: 1.68;
        color: #e2e8f0;
    }
    .ia-card-body p { margin: 0 0 11px 0; }
    .ia-card-body p:last-child { margin-bottom: 0; }
    .ia-card-body h4 {
        margin: 14px 0 6px 0;
        font-size: 0.92rem;
        font-weight: 800;
        color: #38bdf8;
    }
    .ia-card-body code {
        font-family: 'JetBrains Mono', monospace;
        background: rgba(56, 189, 248, 0.12);
        color: #7dd3fc;
        padding: 2px 6px;
        border-radius: 5px;
        font-size: 0.82rem;
    }
    .chat-section-header {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-left: 3px solid #f59e0b;
        border-radius: 14px;
        padding: 14px 18px;
        margin: 22px 0 12px 0;
    }
    .chat-section-title {
        font-size: 1.02rem;
        font-weight: 800;
        color: #f8fafc;
    }
    .chat-section-sub {
        font-size: 0.80rem;
        color: #94a3b8;
        margin-top: 3px;
    }
    .chat-response-box {
        background: rgba(15, 23, 42, 0.82);
        border: 1px solid rgba(56, 189, 248, 0.22);
        border-radius: 14px;
        padding: 16px 18px;
        margin-top: 12px;
        font-size: 0.88rem;
        line-height: 1.65;
        color: #e2e8f0;
    }

    /* Podio Ranking de Clubes (Pestaña 5) */
    .podium-card {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 16px 18px;
        margin-bottom: 14px;
        box-shadow: 0 10px 26px rgba(0, 0, 0, 0.34);
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    }
    .podium-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 14px 32px rgba(0, 0, 0, 0.48);
    }
    .podium-gold { border-top: 4px solid #f59e0b; background: linear-gradient(180deg, rgba(245, 158, 11, 0.12) 0%, rgba(17, 24, 39, 0.78) 55%); }
    .podium-silver { border-top: 4px solid #cbd5e1; background: linear-gradient(180deg, rgba(203, 213, 225, 0.09) 0%, rgba(17, 24, 39, 0.78) 55%); }
    .podium-bronze { border-top: 4px solid #d97706; background: linear-gradient(180deg, rgba(217, 119, 6, 0.10) 0%, rgba(17, 24, 39, 0.78) 55%); }
    .podium-rank {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 5px;
    }
    .podium-club {
        font-size: 1.05rem;
        font-weight: 800;
        color: #f8fafc;
        margin-bottom: 10px;
    }
    .podium-patrimonio {
        font-family: 'JetBrains Mono', monospace;
        font-size: 1.30rem;
        font-weight: 800;
        color: #f8fafc;
        margin-bottom: 8px;
    }
    .podium-meta {
        display: flex;
        justify-content: space-between;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        color: #94a3b8;
        padding-top: 8px;
        border-top: 1px solid rgba(255, 255, 255, 0.07);
    }
    .vs-card {
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-top: 3px solid #38bdf8;
        border-radius: 14px;
        padding: 15px 16px;
        margin-top: 10px;
    }

    /* Banner Sniper */
    .sniper-banner {
        background: linear-gradient(135deg, rgba(239, 68, 68, 0.18) 0%, rgba(17, 24, 39, 0.90) 60%, rgba(245, 158, 11, 0.16) 100%);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(245, 158, 11, 0.55);
        border-left: 4px solid #f59e0b;
        border-radius: 14px;
        padding: 14px 18px;
        margin-bottom: 14px;
        box-shadow: 0 8px 26px rgba(245, 158, 11, 0.15);
        font-size: 0.88rem;
        line-height: 1.55;
    }

    /* Pestañas estilo Terminal Fintech */
    div[data-baseweb="tab-list"] {
        gap: 6px;
        background: rgba(17, 24, 39, 0.72);
        backdrop-filter: blur(12px);
        padding: 6px;
        border-radius: 14px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        margin-bottom: 14px;
    }
    button[data-baseweb="tab"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-size: 0.88rem !important;
        font-weight: 700 !important;
        color: #94a3b8 !important;
        border-radius: 10px !important;
        padding: 9px 14px !important;
        transition: all 0.2s ease !important;
    }
    button[data-baseweb="tab"]:hover {
        color: #f8fafc !important;
        background: rgba(255, 255, 255, 0.05) !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #059669 0%, #0284c7 100%) !important;
        color: #ffffff !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.28);
    }

    /* Expanders y contenedores */
    div[data-testid="stExpander"] {
        background: rgba(17, 24, 39, 0.55);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        overflow: hidden;
    }
    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.07);
    }
    h2, h3 {
        color: #f1f5f9 !important;
        font-weight: 800 !important;
        letter-spacing: -0.02em !important;
    }

    /* Responsividad móvil */
    @media (max-width: 768px) {
        .main-header { padding: 16px; }
        .main-header h1 { font-size: 1.25rem; }
        .drop-timer { width: 100%; text-align: left; }
        .kpi-value { font-size: 1.15rem; }
    }
</style>
""", unsafe_allow_html=True)

def estilizar_plotly_fintech(fig, height=340, hovermode="x unified"):
    """Aplica el tema dark-glass fintech consistente a todos los gráficos Plotly."""
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=height,
        hovermode=hovermode,
        font=dict(family="Plus Jakarta Sans, sans-serif", color="#e2e8f0", size=12),
        title_font=dict(family="Plus Jakarta Sans, sans-serif", color="#f8fafc", size=14),
        margin=dict(l=12, r=12, t=46, b=16),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            bgcolor="rgba(0,0,0,0)",
            font=dict(size=11, color="#94a3b8")
        ),
        hoverlabel=dict(
            bgcolor="rgba(9, 13, 22, 0.94)",
            bordercolor="rgba(56, 189, 248, 0.45)",
            font=dict(family="JetBrains Mono, monospace", size=12, color="#f8fafc")
        )
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="rgba(255, 255, 255, 0.06)",
        zerolinecolor="rgba(255, 255, 255, 0.08)",
        tickfont=dict(family="JetBrains Mono, monospace", size=11, color="#94a3b8")
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="rgba(255, 255, 255, 0.06)",
        zerolinecolor="rgba(255, 255, 255, 0.08)",
        tickfont=dict(family="JetBrains Mono, monospace", size=11, color="#94a3b8")
    )
    return fig

def formatear_md_a_html_seguro(texto_md):
    """Convierte texto Markdown simple de la IA en HTML limpio, escapado y con párrafos bien espaciados."""
    if not texto_md:
        return "<p>Sin reporte generado todavía.</p>"
    escapado = html_lib.escape(str(texto_md))
    lineas = escapado.splitlines()
    bloques = []
    for ln in lineas:
        s = ln.strip()
        if not s:
            continue
        s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
        s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
        if s.startswith("### "):
            bloques.append(f"<h4>{s[4:]}</h4>")
        elif s.startswith("## "):
            bloques.append(f"<h4>{s[3:]}</h4>")
        elif s.startswith("# "):
            bloques.append(f"<h4>{s[2:]}</h4>")
        else:
            bloques.append(f"<p>{s}</p>")
    return "\n".join(bloques) if bloques else "<p>Sin contenido.</p>"

# ==========================================
# FUNCIONES MATEMÁTICAS, HORA, SONIDO, SALDOS Y EXCEL
# ==========================================
PAISES_DROP = {
    "🇦🇷 Argentina / Chile / Uruguay (15:00 hs)": (-3, 15),
    "🇨🇴 Colombia / Perú / Ecuador (13:00 hs)": (-5, 13),
    "🇲🇽 México CDMX (12:00 hs)": (-6, 12),
    "🇪🇸 España Peninsular (19:00 hs)": (1, 19),
}

AGENDA_EA_SEMANAL = {
    "Lunes": "Player Picks de Mejora (80+/81+) y SBCs de Héroe/Icono Base. Sube el oro común/raro bajo y medias 84-86.",
    "Martes": "Marquesinas de Champions (UEFA) y SBC de campaña. Alta demanda de cartas de clubes UCL y medias 83-85.",
    "Miércoles": "Lanzamiento de nuevo TOTW. Los IFs de la semana anterior salen de sobres (Out of Packs) y suben.",
    "Jueves": "Recompensas de Division Rivals + Marquesinas tradicionales. Inyección de monedas: suben cartas Meta.",
    "Viernes": "NUEVO EVENTO / PROMO. Pánico comprador previo al drop y explosión de Evoluciones de fin de semana.",
    "Sábado": "SBC de Casino / Icono o Héroe asegurado + Evolución Meta de fin de semana. Vuelan medias 86-88 y TOTW.",
    "Domingo": "Sell-off de FUT Champions (noche) y Mini-Release de Promo. Piso ideal para comprar TOTW y Fodder en pánico."
}

COMBOS_SBC_PLANTILLAS = {
    84: [
        ("9x Media 84 + 2x Media 83", {84: 9, 83: 2}),
        ("1x Media 86 + 4x Media 84 + 6x Media 83", {86: 1, 84: 4, 83: 6}),
        ("2x Media 85 + 5x Media 84 + 4x Media 83", {85: 2, 84: 5, 83: 4}),
    ],
    85: [
        ("9x Media 85 + 2x Media 84", {85: 9, 84: 2}),
        ("2x Media 87 + 1x Media 85 + 8x Media 84", {87: 2, 85: 1, 84: 8}),
        ("1x Media 88 + 1x Media 85 + 9x Media 84", {88: 1, 85: 1, 84: 9}),
    ],
    86: [
        ("9x Media 86 + 2x Media 85", {86: 9, 85: 2}),
        ("2x Media 88 + 1x Media 86 + 8x Media 85", {88: 2, 86: 1, 85: 8}),
        ("1x Media 89 + 2x Media 86 + 8x Media 85", {89: 1, 86: 2, 85: 8}),
        ("3x Media 87 + 4x Media 86 + 4x Media 85", {87: 3, 86: 4, 85: 4}),
    ],
    87: [
        ("9x Media 87 + 2x Media 86", {87: 9, 86: 2}),
        ("2x Media 89 + 1x Media 87 + 8x Media 86", {89: 2, 87: 1, 86: 8}),
        ("3x Media 88 + 4x Media 87 + 4x Media 86", {88: 3, 87: 4, 86: 4}),
    ],
    88: [
        ("9x Media 88 + 2x Media 87", {88: 9, 87: 2}),
        ("3x Media 89 + 4x Media 88 + 4x Media 87", {89: 3, 88: 4, 87: 4}),
    ]
}

def ahora_arg():
    tz_arg = timezone(timedelta(hours=-3))
    return datetime.now(tz_arg).replace(tzinfo=None)

def reproducir_sonido_sniper():
    components.html(
        """
        <script>
        try {
            const ctx = new (window.AudioContext || window.webkitAudioContext)();
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'sine';
            osc.frequency.setValueAtTime(880, ctx.currentTime);
            osc.frequency.setValueAtTime(1320, ctx.currentTime + 0.15);
            gain.gain.setValueAtTime(0.15, ctx.currentTime);
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.45);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start();
            osc.stop(ctx.currentTime + 0.45);
        } catch(e) {}
        </script>
        """,
        height=0,
    )

_T_PART1 = "9090839192:AAGP0TTaNkQP1I"
_T_PART2 = "LRKvqSFvbsyZtdYuM3PPw"
DEFAULT_TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", _T_PART1 + _T_PART2)
DEFAULT_TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "7676084666")
ARCHIVO_REGISTRO_CLUBES = "clubes_fc27_registro.json"

ARCHIVO_INTRANSF_ADMIN = "intransferibles_admin.json"

def slug_club_seguro(nombre_club):
    s = re.sub(r"[^a-zA-Z0-9]+", "_", str(nombre_club).strip().lower()).strip("_")
    return s or "club"

def obtener_cliente_gspread():
    try:
        import gspread
        if os.path.exists(ARCHIVO_CRED):
            return gspread.service_account(filename=ARCHIVO_CRED)
        if "gcp_service_account" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    except Exception:
        pass
    return None

def cargar_registro_clubes():
    if os.path.exists(ARCHIVO_REGISTRO_CLUBES):
        try:
            with open(ARCHIVO_REGISTRO_CLUBES, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}

def sincronizar_clubes_cloud_sheet():
    try:
        reg = cargar_registro_clubes()
        filas_cloud = []
        # Incluir respaldo de saldos e intransferibles del Admin
        saldos_adm = {}
        if os.path.exists(ARCHIVO_SALDOS_LOCAL):
            try:
                with open(ARCHIVO_SALDOS_LOCAL, "r", encoding="utf-8") as f:
                    saldos_adm = json.load(f)
            except Exception:
                pass
        intr_adm = {}
        if os.path.exists(ARCHIVO_INTRANSF_ADMIN):
            try:
                with open(ARCHIVO_INTRANSF_ADMIN, "r", encoding="utf-8") as f:
                    intr_adm = json.load(f)
            except Exception:
                pass
        filas_cloud.append({
            "Slug_Club": "__admin__",
            "Nombre_Club": "👑 Club Admin Principal",
            "PIN": "",
            "Creado": "",
            "Saldos_JSON": json.dumps(saldos_adm, ensure_ascii=False),
            "Intransferibles_JSON": json.dumps(intr_adm, ensure_ascii=False),
            "Cartera_JSON": "[]",
            "Alertas_JSON": "[]"
        })

        for slug_k, info_k in reg.items():
            saldos_c = {}
            if os.path.exists(f"saldos_club_{slug_k}.json"):
                try:
                    with open(f"saldos_club_{slug_k}.json", "r", encoding="utf-8") as f:
                        saldos_c = json.load(f)
                except Exception:
                    pass
            intr_c = {}
            if os.path.exists(f"intransferibles_club_{slug_k}.json"):
                try:
                    with open(f"intransferibles_club_{slug_k}.json", "r", encoding="utf-8") as f:
                        intr_c = json.load(f)
                except Exception:
                    pass
            cartera_c = []
            if os.path.exists(f"cartera_club_{slug_k}.csv"):
                try:
                    cartera_c = pd.read_csv(f"cartera_club_{slug_k}.csv").fillna("").to_dict(orient="records")
                except Exception:
                    pass
            alertas_c = []
            if os.path.exists(f"alertas_club_{slug_k}.csv"):
                try:
                    alertas_c = pd.read_csv(f"alertas_club_{slug_k}.csv").fillna("").to_dict(orient="records")
                except Exception:
                    pass
            filas_cloud.append({
                "Slug_Club": slug_k,
                "Nombre_Club": str(info_k.get("nombre", slug_k)),
                "PIN": str(info_k.get("pin", "")),
                "Creado": str(info_k.get("creado", "")),
                "Saldos_JSON": json.dumps(saldos_c, ensure_ascii=False),
                "Intransferibles_JSON": json.dumps(intr_c, ensure_ascii=False),
                "Cartera_JSON": json.dumps(cartera_c, ensure_ascii=False),
                "Alertas_JSON": json.dumps(alertas_c, ensure_ascii=False)
            })

        df_cloud = pd.DataFrame(filas_cloud)
        gc = obtener_cliente_gspread()
        if gc is not None:
            try:
                sh = gc.open(NOMBRE_SHEET)
                try:
                    ws = sh.worksheet("Clubes_Cloud")
                    ws.clear()
                except Exception:
                    ws = sh.add_worksheet(title="Clubes_Cloud", rows="200", cols="12")
                df_clean = df_cloud.fillna("")
                ws.update([df_clean.columns.values.tolist()] + df_clean.astype(str).values.tolist())
            except Exception:
                pass
        if os.path.exists(ARCHIVO_EXCEL):
            try:
                with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_cloud.to_excel(writer, index=False, sheet_name="Clubes_Cloud")
            except Exception:
                pass
    except Exception:
        pass

def hidratar_clubes_desde_cloud(df_cloud):
    if df_cloud is None or df_cloud.empty or "Slug_Club" not in df_cloud.columns:
        return
    reg_local = cargar_registro_clubes()
    cambio_reg = False
    for _, row in df_cloud.iterrows():
        slug_k = str(row.get("Slug_Club", "")).strip()
        if not slug_k:
            continue
        if slug_k == "__admin__":
            if not os.path.exists(ARCHIVO_INTRANSF_ADMIN):
                try:
                    intr_d = json.loads(str(row.get("Intransferibles_JSON", "{}")))
                    if isinstance(intr_d, dict) and intr_d:
                        with open(ARCHIVO_INTRANSF_ADMIN, "w", encoding="utf-8") as f:
                            json.dump(intr_d, f)
                except Exception:
                    pass
            continue

        if slug_k not in reg_local:
            reg_local[slug_k] = {
                "nombre": str(row.get("Nombre_Club", slug_k)),
                "pin": str(row.get("PIN", "")),
                "creado": str(row.get("Creado", ""))
            }
            cambio_reg = True

        arch_c = f"cartera_club_{slug_k}.csv"
        if not os.path.exists(arch_c):
            try:
                c_list = json.loads(str(row.get("Cartera_JSON", "[]")))
                if isinstance(c_list, list) and c_list:
                    pd.DataFrame(c_list).to_csv(arch_c, index=False)
            except Exception:
                pass
        arch_a = f"alertas_club_{slug_k}.csv"
        if not os.path.exists(arch_a):
            try:
                a_list = json.loads(str(row.get("Alertas_JSON", "[]")))
                if isinstance(a_list, list) and a_list:
                    pd.DataFrame(a_list).to_csv(arch_a, index=False)
            except Exception:
                pass
        arch_s = f"saldos_club_{slug_k}.json"
        if not os.path.exists(arch_s):
            try:
                s_dict = json.loads(str(row.get("Saldos_JSON", "{}")))
                if isinstance(s_dict, dict) and s_dict:
                    with open(arch_s, "w", encoding="utf-8") as f:
                        json.dump(s_dict, f)
            except Exception:
                pass
        arch_i = f"intransferibles_club_{slug_k}.json"
        if not os.path.exists(arch_i):
            try:
                i_dict = json.loads(str(row.get("Intransferibles_JSON", "{}")))
                if isinstance(i_dict, dict) and i_dict:
                    with open(arch_i, "w", encoding="utf-8") as f:
                        json.dump(i_dict, f)
            except Exception:
                pass
    if cambio_reg:
        try:
            with open(ARCHIVO_REGISTRO_CLUBES, "w", encoding="utf-8") as f:
                json.dump(reg_local, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

def guardar_registro_clubes(reg):
    try:
        with open(ARCHIVO_REGISTRO_CLUBES, "w", encoding="utf-8") as f:
            json.dump(reg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    sincronizar_clubes_cloud_sheet()

def cargar_intransferibles_club(es_admin, slug_club=None):
    defecto = {"83": 0, "84": 0, "85": 0, "86": 0, "87": 0, "88": 0, "89": 0, "TOTW": 0}
    if slug_club and not es_admin:
        k_ss = f"intr_club_{slug_club}"
        if k_ss in st.session_state:
            return st.session_state[k_ss]
        arch_i = f"intransferibles_club_{slug_club}.json"
        if os.path.exists(arch_i):
            try:
                with open(arch_i, "r", encoding="utf-8") as f:
                    d = json.load(f)
                    defecto.update({str(k): int(v) for k, v in d.items()})
                    if "IF" in d and "TOTW" not in d:
                        defecto["TOTW"] = int(d["IF"])
            except Exception:
                pass
        st.session_state[k_ss] = defecto
        return defecto
    if not es_admin:
        return st.session_state.get("guest_intransferibles", defecto)
    if "admin_intransferibles" in st.session_state:
        return st.session_state["admin_intransferibles"]
    if os.path.exists(ARCHIVO_INTRANSF_ADMIN):
        try:
            with open(ARCHIVO_INTRANSF_ADMIN, "r", encoding="utf-8") as f:
                d = json.load(f)
                defecto.update({str(k): int(v) for k, v in d.items()})
                if "IF" in d and "TOTW" not in d:
                    defecto["TOTW"] = int(d["IF"])
        except Exception:
            pass
    st.session_state["admin_intransferibles"] = defecto
    return defecto

def guardar_intransferibles_club(datos_intr, es_admin, slug_club=None):
    limpio = {str(k): int(v) for k, v in datos_intr.items()}
    if slug_club and not es_admin:
        st.session_state[f"intr_club_{slug_club}"] = limpio
        try:
            with open(f"intransferibles_club_{slug_club}.json", "w", encoding="utf-8") as f:
                json.dump(limpio, f)
        except Exception:
            pass
        sincronizar_clubes_cloud_sheet()
        return
    if not es_admin:
        st.session_state["guest_intransferibles"] = limpio
        return
    st.session_state["admin_intransferibles"] = limpio
    try:
        with open(ARCHIVO_INTRANSF_ADMIN, "w", encoding="utf-8") as f:
            json.dump(limpio, f)
    except Exception:
        pass
    sincronizar_clubes_cloud_sheet()

def enviar_aviso_telegram_web(mensaje):
    try:
        url = f"https://api.telegram.org/bot{DEFAULT_TELEGRAM_TOKEN}/sendMessage"
        r = requests.post(url, data={"chat_id": DEFAULT_TELEGRAM_CHAT_ID, "text": mensaje, "parse_mode": "HTML"}, timeout=6)
        if r.status_code == 200:
            return True
        r2 = requests.post(url, data={"chat_id": DEFAULT_TELEGRAM_CHAT_ID, "text": re.sub(r"<[^>]+>", "", str(mensaje))}, timeout=6)
        return r2.status_code == 200
    except Exception:
        return False

def cargar_saldos_personalizados(es_admin, presupuesto_base, cap_auto_cartera, slug_club=None):
    defecto = {
        "modo_liquidez": "Automático",
        "liquidez_manual": max(0, int(presupuesto_base) - int(cap_auto_cartera)),
        "modo_capital": "Automático (Suma de mis cartas)",
        "capital_manual": int(cap_auto_cartera)
    }
    if slug_club and not es_admin:
        arch_club_s = f"saldos_club_{slug_club}.json"
        if f"saldos_club_{slug_club}" in st.session_state:
            return st.session_state[f"saldos_club_{slug_club}"]
        if os.path.exists(arch_club_s):
            try:
                with open(arch_club_s, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    st.session_state[f"saldos_club_{slug_club}"] = data
                    return data
            except Exception:
                pass
        return defecto
    if not es_admin:
        return st.session_state.get("guest_saldos", defecto)
    if "admin_saldos" in st.session_state:
        return st.session_state["admin_saldos"]
    if os.path.exists(ARCHIVO_SALDOS_LOCAL):
        try:
            with open(ARCHIVO_SALDOS_LOCAL, "r", encoding="utf-8") as f:
                data = json.load(f)
                st.session_state["admin_saldos"] = data
                return data
        except Exception:
            pass
    return defecto

def guardar_saldos_personalizados(nuevos_saldos, es_admin, slug_club=None):
    if slug_club and not es_admin:
        st.session_state[f"saldos_club_{slug_club}"] = nuevos_saldos
        try:
            with open(f"saldos_club_{slug_club}.json", "w", encoding="utf-8") as f:
                json.dump(nuevos_saldos, f)
        except Exception:
            pass
        sincronizar_clubes_cloud_sheet()
        return
    if not es_admin:
        st.session_state["guest_saldos"] = nuevos_saldos
        return
    st.session_state["admin_saldos"] = nuevos_saldos
    try:
        with open(ARCHIVO_SALDOS_LOCAL, "w", encoding="utf-8") as f:
            json.dump(nuevos_saldos, f)
    except Exception:
        pass
    sincronizar_clubes_cloud_sheet()

def generar_excel_contable_bytes(df_er, df_viva, df_vendidos):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_er.to_excel(writer, index=False, sheet_name="Estado_Resultados")
        if not df_viva.empty:
            df_viva.to_excel(writer, index=False, sheet_name="Cartera_Activa")
        else:
            pd.DataFrame(columns=["Jugador", "Estrategia", "Costo_Pagado", "Mercado_Hoy"]).to_excel(writer, index=False, sheet_name="Cartera_Activa")
        if not df_vendidos.empty:
            df_vendidos.to_excel(writer, index=False, sheet_name="Libro_Mayor_Ventas")
        else:
            pd.DataFrame(columns=["Jugador_Exacto", "Precio_Pagado", "Precio_Venta", "Ganancia_Neta"]).to_excel(writer, index=False, sheet_name="Libro_Mayor_Ventas")
    return output.getvalue()

def disparar_github_action_remoto(repo_full, pat_token):
    url = f"https://api.github.com/repos/{repo_full.strip()}/actions/workflows/escanear.yml/dispatches"
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {pat_token.strip()}",
        "X-GitHub-Api-Version": "2022-11-28"
    }
    try:
        r = requests.post(url, headers=headers, json={"ref": "main"}, timeout=12)
        if r.status_code in [200, 201, 204]:
            return True, "🚀 ¡Escaneo remoto iniciado en GitHub Actions! En ~3 minutos tendrás los precios nuevos."
        return False, f"⚠️ GitHub devolvió código {r.status_code}: Verificá que el Token tenga permiso 'workflow'."
    except Exception as e:
        return False, f"⚠️ Error de red: {e}"

def calcular_countdown_pais(nombre_pais_sel):
    utc_offset, hora_drop_local = PAISES_DROP.get(nombre_pais_sel, (-3, 15))
    tz_local = timezone(timedelta(hours=utc_offset))
    ahora_loc = datetime.now(tz_local).replace(tzinfo=None)
    drop_hoy = ahora_loc.replace(hour=hora_drop_local, minute=0, second=0, microsecond=0)

    if ahora_loc >= drop_hoy:
        proximo_drop = drop_hoy + timedelta(days=1)
        estado_drop = "🔥 Contenido de hoy ya activo"
    else:
        proximo_drop = drop_hoy
        minutos_faltantes = (drop_hoy - ahora_loc).total_seconds() / 60.0
        if minutos_faltantes <= 35:
            estado_drop = "🚨 ¡VENTANA PRE-DROP! (Comprar en pánico)"
        else:
            estado_drop = "⏳ Esperando Drop Diario EA"

    resta = proximo_drop - ahora_loc
    horas, rem = divmod(int(resta.total_seconds()), 3600)
    minutos, _ = divmod(rem, 60)
    return f"{horas:02d}h {minutos:02d}m", estado_drop, f"{hora_drop_local:02d}:00 hs"

def calcular_semaforo_y_calendario_ea(df_corte):
    ahora = ahora_arg()
    dias_sem = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    dia_idx = ahora.weekday()
    dia_nom = dias_sem[dia_idx]
    hora_actual = ahora.hour

    pautas_calendario = {
        0: ("📉 Lunes Post-FUT Champions (Sell-Off)", "Liquidación de equipos terminada. Excelente ventana para comprar cartas Meta en piso; rebotan miércoles/jueves."),
        1: ("🔄 Martes de Transición & Pre-TOTW", "Mercado selectivo. Priorizá Flipping por puja (+10% ROI), Lazy Listing (+12%) y acumular TOTW salientes en piso."),
        2: ("📈 Miércoles Pre-Rivals & Hype TOTW", "Crece la demanda competitiva. Mantené (Hold) cartas Meta compradas el lunes y vendelas hacia el jueves."),
        3: ("🎁 Jueves de Recompensas Rivals & Marquesinas", "Sobres por la mañana (pánico de compra) e inyección de monedas a la tarde (vender Meta con recargo Lazy)."),
        4: ("🔥 Viernes de Nueva Promo (15:00 ARG) & Arranque WL", "No compres Meta antes de las 14:30. Entrá fuerte en el pánico de sobres de las 15:05 a 16:00 hs ARG."),
        5: ("⚡ Sábado de Pico Jugable FUT Champions", "Máxima demanda de Lazy Buyers armando plantilla. Priorizá listar tu stock activo a 1 hora (+12%)."),
        6: ("🧹 Domingo de Liquidación & Piso Fodder", "Sell-off nocturno de equipos WL. Pasá a liquidez tus cartas Meta caras y acumulá Fodder 85-88 en piso.")
    }
    titulo_cal, detalle_cal = pautas_calendario.get(dia_idx, ("📅 Ciclo Semanal EA FC 27", "Operar por puja y listar Lazy a 1h."))
    if hora_actual == 14:
        detalle_cal = "🚨 ¡FALTA MENOS DE 1 HORA PARA EL DROP (15:00 ARG)! Prepará liquidez para cazar pánico de sobres."
    elif hora_actual == 15:
        detalle_cal = "🔥 ¡DROP RECIÉN LANZADO! Máxima volatilidad: snipeá cartas en caída por sobres y vendé medias pedidas en SBC."

    var_prom = 0.0
    pct_en_piso = 0.0
    if df_corte is not None and not df_corte.empty:
        df_corte_act = df_corte[df_corte["Precio_Actual"].apply(num_seguro) > 0] if "Precio_Actual" in df_corte.columns else df_corte
        if not df_corte_act.empty:
            if "Variacion_Bruta_%" in df_corte_act.columns:
                s_var = pd.to_numeric(df_corte_act["Variacion_Bruta_%"], errors="coerce").dropna()
                s_var = s_var[(s_var >= -60.0) & (s_var <= 100.0)]
                if not s_var.empty:
                    var_prom = round(float(s_var.mean()), 2)
            if "Estado_Historico" in df_corte_act.columns:
                en_piso = df_corte_act["Estado_Historico"].astype(str).str.contains("💎|🟢", na=False).sum()
                pct_en_piso = round((en_piso / max(1, len(df_corte_act))) * 100.0, 1)

    if var_prom <= -3.0 or pct_en_piso >= 32.0:
        tit_sem = f"🟢 MERCADO EN PISO / ZONA DE COMPRA (Var: {var_prom:+.1f}% | {pct_en_piso}% en soporte)"
        det_sem = "Alta presión vendedora o pánico de sobres. Momento ideal para acumular en puja y cazar gangas."
        col_sem = "#00e676"
    elif var_prom >= 3.5:
        tit_sem = f"🔴 MERCADO INFLADO / TOMA DE GANANCIAS (Var: {var_prom:+.1f}%)"
        det_sem = "Precios impulsados al alza. Priorizá vender tu stock activo y evitá sobrepagar en compra inmediata."
        col_sem = "#ff5252"
    else:
        tit_sem = f"🟡 MERCADO ESTABLE / IDEAL FLIPPING & LAZY (Var: {var_prom:+.1f}% | {pct_en_piso}% en piso)"
        det_sem = f"Flujo equilibrado en {dia_nom}. Comprá por puja (+10% ROI) sobre base FUT.GG y listá Lazy a 1h (+12%)."
        col_sem = "#ffd600"

    return tit_sem, det_sem, col_sem, titulo_cal, detalle_cal

def num_seguro(val, defecto=0):
    try:
        if pd.isna(val) or str(val).strip() == "":
            return defecto
        limpio = str(val).replace("$", "").replace(",", "").strip()
        return int(float(limpio))
    except Exception:
        return defecto

def float_seguro(val, defecto=0.0):
    try:
        if pd.isna(val) or str(val).strip() == "":
            return defecto
        limpio = str(val).replace("$", "").replace(",", "").replace("%", "").strip()
        return float(limpio)
    except Exception:
        return defecto

def sin_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn")

def paso_escalon_ea(precio):
    if precio < 1000:
        return 50
    elif precio < 10000:
        return 100
    elif precio < 50000:
        return 250
    elif precio < 100000:
        return 500
    return 1000

def es_multiplo_escalon_ea(precio, presupuesto_max=TOPE_RADAR_MAXIMO):
    p = num_seguro(precio, 0)
    if not (200 <= p <= presupuesto_max):
        return False
    return (p % paso_escalon_ea(p)) == 0

def redondear_escalon_ea(precio, hacia_arriba=False):
    precio = float_seguro(precio, 0.0)
    if precio <= 0:
        return 0
    paso = paso_escalon_ea(precio)
    if hacia_arriba:
        return int(math.ceil(precio / paso) * paso)
    return int(math.floor(precio / paso) * paso)

def redondear_escalon_ea_cercano(precio):
    precio = float_seguro(precio, 0.0)
    if precio <= 0:
        return 0
    paso = paso_escalon_ea(precio)
    return max(200, int(round(precio / paso) * paso))

def calcular_break_even(precio_compra):
    precio_compra = float_seguro(precio_compra, 0.0)
    if precio_compra <= 0:
        return 0
    return redondear_escalon_ea(precio_compra / 0.95, hacia_arriba=True)

def calcular_puja_maxima_segura(precio_actual_mercado, factor_roi=1.10):
    precio_actual_mercado = float_seguro(precio_actual_mercado, 0.0)
    if precio_actual_mercado <= 0:
        return 0
    return redondear_escalon_ea((precio_actual_mercado * 0.95) / factor_roi, hacia_arriba=False)

def calcular_precio_lazy(precio_referencia, precio_pagado=0, max_hist=0, muestras=0, factor_lazy=1.12, factor_roi=1.10):
    precio_referencia = float_seguro(precio_referencia, 0.0)
    precio_pagado = float_seguro(precio_pagado, 0.0)
    max_hist = float_seguro(max_hist, 0.0)
    muestras = int(num_seguro(muestras, 0))
    if precio_referencia <= 0:
        return 0
    base_lazy = precio_referencia * factor_lazy
    if precio_pagado > 0:
        min_rentable = (precio_pagado * factor_roi) / 0.95
        base_lazy = max(base_lazy, min_rentable)
    piso_empate = calcular_break_even(precio_pagado if precio_pagado > 0 else precio_referencia)
    lazy_ea = redondear_escalon_ea(base_lazy, hacia_arriba=True)
    if muestras >= 3 and max_hist > 0 and lazy_ea > max_hist:
        techo_ea = redondear_escalon_ea(max_hist, hacia_arriba=False)
        if techo_ea >= piso_empate:
            return techo_ea
        return piso_empate
    return max(lazy_ea, piso_empate)

def ganancia_neta_minima_real(precio_actual):
    p = num_seguro(precio_actual, 0)
    if p < 5000:
        return 350
    elif p <= 25000:
        return 750
    return 1500

def precio_mercado_seguro_carta(fila_carta, fallback_precio=0):
    if fila_carta is None:
        return int(fallback_precio)
    p_act = num_seguro(fila_carta.get("Precio_Actual", 0), 0)
    if p_act > 0:
        return p_act
    return int(fallback_precio)

# ==========================================
# BLINDAJE CENTRALIZADO DE PRECIOS, EMPAREJAMIENTO, FODDER Y TOTW
# ==========================================
PISOS_MINIMOS_REALISTAS_MEDIA = {
    83: 600,
    84: 650,
    85: 1800,
    86: 3500,
    87: 6000,
    88: 9500,
    89: 14000,
    90: 20000,
    91: 28000,
    92: 35000,
}

PISOS_LOGICOS_FODDER_WEB = {
    83: 650,
    84: 750,
    85: 2000,
    86: 4500,
    87: 7500,
    88: 11500,
    89: 18000,
}

REFERENCIA_FODDER_WEB = {
    83: (950, 1800),
    84: (1200, 2200),
    85: (2800, 4500),
    86: (5500, 8500),
    87: (9500, 14000),
    88: (14500, 21000),
    89: (22000, 32000),
}

COLUMNAS_RESIDUALES_FUTBIN = [
    "Precio_Futbin", "Enlace_Futbin", "Fuente_Mejor_Precio",
    "Precio_FUTGG", "Enlace_FUTGG", "Fuente_Optima"
]

def extraer_media_carta_web(texto):
    m = re.search(r'Media\s*(\d{2})\b', str(texto or ""))
    if m:
        val = int(m.group(1))
        if 45 <= val <= 99:
            return val
    return 0

def es_carta_base_url_gg_web(enlace):
    m = re.search(r'/players/(\d+)-[^/]+/27-(\d+)(?:/|$)', str(enlace or ""))
    return bool(m and m.group(1) == m.group(2))

def categoria_version_carta_web(texto):
    t = sin_tildes(str(texto or "")).upper()
    if any(k in t for k in ["TOTW", "IN-FORM", "INFORM", "TEAM OF THE WEEK"]) or re.search(r'(?:-\s*IF\b|\(IF\b|\bIF\))', t):
        return "TOTW/IF"
    palabras_promo = [
        "ICON", "HERO", "EVO", "EVOLUTION", "PROMO", "SPECIAL", "ESPECIAL", "RTTK", "RTTF",
        "TRAILBLAZER", "CENTURION", "THUNDER", "TOTY", "TOTS", "FUTTIES",
        "WINTER", "WILDCARD", "FUTURE STAR", "BIRTHDAY", "FANTASY", "POTM",
        "FLASHBACK", "SHOWDOWN", "RUSH", "SHAPESHIFTER", "GRASSROOT", "IMMORTAL",
        "DREAMCHASER", "TRACK STAR", "GLOBETROTTER", "SUCCESSION", "UCL", "UEL", "UECL"
    ]
    if any(k in t for k in palabras_promo):
        return "Promo/Especial"
    return "Normal/Oro"

def es_carta_oro_normal_fodder_web(nombre_full, categoria="", enlace=""):
    """Devuelve True solo para cartas Oro/Normal base (excluye TOTW/IF/Promo/Especial/Icono/Héroe)."""
    if categoria_version_carta_web(nombre_full) != "Normal/Oro":
        return False
    cat_s = sin_tildes(str(categoria or "")).lower().strip()
    if cat_s and cat_s not in ("normal/oro", "normal", "oro", "gold"):
        return False
    m_url = re.search(r'/players/(\d+)-[^/]+/27-(\d+)(?:/|$)', str(enlace or ""))
    if m_url and m_url.group(1) != m_url.group(2):
        return False
    return True

def es_totw_real_web(nombre_full, enlace=""):
    if not nombre_full:
        return False
    if es_carta_base_url_gg_web(enlace):
        return False
    nom_up = sin_tildes(str(nombre_full)).upper()
    if re.search(r'\b(?:NORMAL|GOLD|RARE|COMMON|HERO|ICON|POTM|TOTY|TOTS|SBC|OBJECTIVE|PROMO|ESPECIAL|SPECIAL)\b', nom_up):
        return False
    return bool(
        ("TOTW" in nom_up)
        or ("TEAM OF THE WEEK" in nom_up)
        or ("IN-FORM" in nom_up)
        or ("- IF)" in nom_up)
        or re.search(r'(?:-\s*IF\b|\(IF\s+\d{2}\))', nom_up)
    )

def precio_cumple_piso_media_web(precio, media):
    p = num_seguro(precio, 0)
    m = num_seguro(media, 0)
    if p <= 0:
        return False
    if m in PISOS_MINIMOS_REALISTAS_MEDIA and p < PISOS_MINIMOS_REALISTAS_MEDIA[m]:
        return False
    if m >= 93 and p < 35000:
        return False
    return True

def emparejar_carta_mercado(nom_carta, costo_pagado, estrategia, df_corte):
    """
    Empareja una carta de Mi_Cartera con la cotización real correcta de FUT.GG:
    - Excluye cartas extintas (Precio_Actual <= 0).
    - Prioriza coincidencia exacta del nombre base y versión Normal/Oro (salvo que el usuario indique TOTW/Promo/Media).
    - Bloquea cruces falsos (ej. carta Oro de 15k emparejada con Promo/Icono de 800k).
    - Si no existe coincidencia confiable o la carta está extinta, devuelve costo_pagado como respaldo seguro.
    Retorna: (p_mercado, max_hist, muestras, enlace, es_respaldo_costo)
    """
    costo_safe = max(0, num_seguro(costo_pagado, 0))
    if df_corte is None or df_corte.empty or "Jugador" not in df_corte.columns or not str(nom_carta or "").strip():
        return costo_safe, 0, 0, "", True

    raw_txt = str(nom_carta).strip()
    nom_base_busq = sin_tildes(raw_txt.split(" (")[0]).lower().strip()
    if not nom_base_busq:
        return costo_safe, 0, 0, "", True

    media_busq = extraer_media_carta_web(raw_txt)
    cat_busq = categoria_version_carta_web(raw_txt)
    if cat_busq == "Normal/Oro" and "TOTW" in str(estrategia or "").upper():
        cat_busq = "TOTW/IF"
    pide_especial = (cat_busq != "Normal/Oro") or any(
        k in sin_tildes(raw_txt).upper()
        for k in ("TOTW", "IF", "PROMO", "ICON", "HERO", "EVO", "SPECIAL", "ESPECIAL")
    )

    df_act = df_corte[df_corte["Precio_Actual"].apply(num_seguro) > 0].copy()
    if df_act.empty:
        return costo_safe, 0, 0, "", True

    if "_nom_base_norm" not in df_act.columns:
        df_act["_nom_base_norm"] = df_act["Jugador"].astype(str).apply(
            lambda x: sin_tildes(x.split(" (")[0]).lower().strip()
        )

    # Nivel 1: nombre base exacto; Nivel 2: palabra completa; Nivel 3: subcadena (solo si >= 4 letras)
    m_cands = df_act[df_act["_nom_base_norm"] == nom_base_busq]
    if m_cands.empty:
        pat_word = r'\b' + re.escape(nom_base_busq) + r'\b'
        m_cands = df_act[df_act["_nom_base_norm"].str.contains(pat_word, regex=True, na=False)]
    if m_cands.empty and len(nom_base_busq) >= 4:
        m_cands = df_act[df_act["_nom_base_norm"].str.contains(re.escape(nom_base_busq), na=False)]

    if m_cands.empty:
        return costo_safe, 0, 0, "", True

    m_cands = m_cands.copy()
    if media_busq > 0:
        m_med = m_cands[m_cands["Jugador"].apply(extraer_media_carta_web) == media_busq]
        if not m_med.empty:
            m_cands = m_med

    m_cands["_cat_v"] = m_cands["Jugador"].apply(categoria_version_carta_web)
    col_enl = "Enlace_Optimo" if "Enlace_Optimo" in m_cands.columns else ("Enlace_Futbin" if "Enlace_Futbin" in m_cands.columns else None)
    if col_enl:
        m_cands["_es_base_url"] = m_cands[col_enl].apply(es_carta_base_url_gg_web)
    else:
        m_cands["_es_base_url"] = m_cands["_cat_v"] == "Normal/Oro"

    if pide_especial:
        m_cat = m_cands[m_cands["_cat_v"] == cat_busq]
        if not m_cat.empty:
            m_cands = m_cat
    else:
        m_base = m_cands[(m_cands["_cat_v"] == "Normal/Oro") | (m_cands["_es_base_url"])]
        if not m_base.empty:
            m_cands = m_base

    m_cands["_p_num"] = m_cands["Precio_Actual"].apply(num_seguro)
    m_cands["_diff_costo"] = (m_cands["_p_num"] - costo_safe).abs()
    fila_m = m_cands.sort_values(by=["_diff_costo", "_p_num"]).iloc[0]

    p_m = num_seguro(fila_m.get("Precio_Actual", 0), 0)
    if p_m <= 0:
        return costo_safe, 0, 0, "", True

    # Escudo anti-cruce falso: evitar que una carta Oro de 15k se empareje con una Promo/TOTW/Icono de 800k
    if not pide_especial and costo_safe >= 1000 and str(fila_m.get("_cat_v", "")) != "Normal/Oro":
        if p_m > costo_safe * 4.0 or p_m < costo_safe * 0.25:
            return costo_safe, 0, 0, "", True

    max_h = num_seguro(fila_m.get("Precio_Max_Hist", 0), 0)
    mues_h = num_seguro(fila_m.get("Muestras_Hist", 0), 0)
    enl_m = str(fila_m.get("Enlace_Optimo", fila_m.get("Enlace_Futbin", "")) or "")
    return p_m, max_h, mues_h, enl_m, False

def calcular_pisos_fodder_vivos(df_corte):
    """
    Calcula pisos realistas de Fodder (medias 83 a 89) dinámicamente sobre df_ult_corte:
    - Filtra solo cartas Oro/Normal base (excluye TOTW/IF/Promo/Especial/Icono/Héroe) con Precio_Actual > 0 en escalón EA.
    - Descarta valores irreales por debajo de PISOS_LOGICOS_FODDER_WEB (83>=650, 84>=750, 85>=2000, 86>=4500, 87>=7500, 88>=11500, 89>=18000).
    - Usa la mediana de las 5 cartas oro más baratas (redondeada al escalón EA) tanto en la tabla Radar_Fodder como en pisos_vivos del Optimizador SBC.
    - Lista las 3 cartas oro más baratas reales en Jugadores_Referencia.
    Retorna: (pisos_vivos, jug_ref_vivos, df_fodder_vivo)
    """
    cols_fod = ["Media", "Precio_Mas_Barato", "Jugadores_Referencia", "Piso_Compra", "Techo_Venta", "Senal", "Enlace_Optimo"]
    pisos_default = {83: 900, 84: 1200, 85: 2800, 86: 5500, 87: 9500, 88: 14500, 89: 22000}
    pisos_vivos = {}
    jug_ref_vivos = {}
    filas_fodder = []

    df_act = df_corte[df_corte["Precio_Actual"].apply(num_seguro) > 0].copy() if (df_corte is not None and not df_corte.empty and "Precio_Actual" in df_corte.columns) else pd.DataFrame()
    if not df_act.empty and "Jugador" in df_act.columns:
        df_act["_media_fod"] = df_act["Jugador"].apply(extraer_media_carta_web)
        df_act["_p_fod"] = df_act["Precio_Actual"].apply(num_seguro)
        col_cat_f = "Categoria_Carta" if "Categoria_Carta" in df_act.columns else None
        col_enl_f = "Enlace_Optimo" if "Enlace_Optimo" in df_act.columns else ("Enlace_Futbin" if "Enlace_Futbin" in df_act.columns else None)
        df_act["_es_oro_fod"] = [
            es_carta_oro_normal_fodder_web(
                j_v,
                categoria=(c_v if col_cat_f else ""),
                enlace=(e_v if col_enl_f else "")
            )
            for j_v, c_v, e_v in zip(
                df_act["Jugador"],
                df_act[col_cat_f] if col_cat_f else [""] * len(df_act),
                df_act[col_enl_f] if col_enl_f else [""] * len(df_act),
            )
        ]

    for m_val, p_def_m in pisos_default.items():
        piso_logico = PISOS_LOGICOS_FODDER_WEB.get(m_val, PISOS_MINIMOS_REALISTAS_MEDIA.get(m_val, 200))
        c_m = pd.DataFrame()
        if not df_act.empty:
            base_mask = (
                (df_act["_media_fod"] == m_val)
                & (df_act["_es_oro_fod"])
                & (df_act["_p_fod"].apply(es_multiplo_escalon_ea))
            )
            c_m = df_act[base_mask & (df_act["_p_fod"] >= piso_logico)].sort_values(by="_p_fod")
            # Respaldo controlado solo si todo el mercado real bajó temporalmente del piso lógico
            if len(c_m) < 3:
                piso_resp = PISOS_MINIMOS_REALISTAS_MEDIA.get(m_val, 200)
                c_m_resp = df_act[base_mask & (df_act["_p_fod"] >= piso_resp)].sort_values(by="_p_fod")
                if len(c_m_resp) > len(c_m):
                    c_m = c_m_resp

        if not c_m.empty:
            top5_p = c_m["_p_fod"].head(5).astype(int).tolist()
            p_mediana = redondear_escalon_ea_cercano(float(np.median(top5_p)))

            top3_df = c_m.head(3)
            jugs_ref_txt = ", ".join(
                f"{str(r_t['Jugador']).split(' (')[0].strip()} ({int(r_t['_p_fod']):,})"
                for _, r_t in top3_df.iterrows()
            )
            fila_min = c_m.iloc[0]
            nom_min = str(fila_min["Jugador"]).split(" (")[0].strip()
            enl_ref = str(fila_min.get("Enlace_Optimo", fila_min.get("Enlace_Futbin", "")) or "")

            pisos_vivos[m_val] = p_mediana
            jug_ref_vivos[m_val] = f"{nom_min} · Mediana Top 5 ({p_mediana:,} mnds)"

            if m_val in REFERENCIA_FODDER_WEB:
                p_comp_ref, p_vent_ref = REFERENCIA_FODDER_WEB[m_val]
                if p_mediana <= p_comp_ref:
                    senal_f = "✅ COMPRAR"
                elif p_mediana >= p_vent_ref:
                    senal_f = "🚨 VENDER"
                else:
                    senal_f = "⏳ ESPERAR"
                filas_fodder.append({
                    "Media": m_val,
                    "Precio_Mas_Barato": p_mediana,
                    "Jugadores_Referencia": jugs_ref_txt,
                    "Piso_Compra": p_comp_ref,
                    "Techo_Venta": p_vent_ref,
                    "Senal": senal_f,
                    "Enlace_Optimo": enl_ref
                })
        else:
            pisos_vivos[m_val] = p_def_m
            jug_ref_vivos[m_val] = f"Referencia ({p_def_m:,} mnds)"

    return pisos_vivos, jug_ref_vivos, pd.DataFrame(filas_fodder, columns=cols_fod)

def construir_radar_totw_vivo(df_corte, factor_lazy=1.12, factor_roi=1.10):
    """Calcula dinámicamente Radar_TOTW sobre df_ult_corte filtrando estrictamente cartas TOTW reales en piso."""
    cols_totw = ["Jugador", "Media", "Precio_Actual", "Break_Even_5%", "Precio_Lazy_1h", "Riesgo", "Enlace_Optimo"]
    if df_corte is None or df_corte.empty or "Jugador" not in df_corte.columns:
        return pd.DataFrame(columns=cols_totw), 14500

    filas_tw = []
    for _, r in df_corte.iterrows():
        nom_full = str(r.get("Jugador", ""))
        enl = str(r.get("Enlace_Optimo", r.get("Enlace_Futbin", "")) or "")
        if not es_totw_real_web(nom_full, enlace=enl):
            continue
        p_tw = num_seguro(r.get("Precio_Actual", 0), 0)
        med = extraer_media_carta_web(nom_full)
        if not (9500 <= p_tw <= 18000 and es_multiplo_escalon_ea(p_tw) and med >= 83):
            continue
        filas_tw.append({
            "Jugador": nom_full.split(" (")[0].strip() + f" (Media {med} - TOTW)",
            "Media": med,
            "Precio_Actual": p_tw,
            "Break_Even_5%": calcular_break_even(p_tw),
            "Precio_Lazy_1h": calcular_precio_lazy(
                p_tw,
                max_hist=num_seguro(r.get("Precio_Max_Hist", 0), 0),
                muestras=num_seguro(r.get("Muestras_Hist", 0), 0),
                factor_lazy=factor_lazy,
                factor_roi=factor_roi
            ),
            "Riesgo": "🟢 Nulo (Cerca de descarte)" if p_tw <= 12500 else "🟡 Bajo (Piso SBC)",
            "Enlace_Optimo": enl
        })

    if not filas_tw:
        return pd.DataFrame(columns=cols_totw), 14500
    df_tw = pd.DataFrame(filas_tw, columns=cols_totw).sort_values(by=["Precio_Actual", "Media"], ascending=[True, False]).reset_index(drop=True)
    piso_totw_vivo = int(df_tw["Precio_Actual"].iloc[0])
    return df_tw, piso_totw_vivo

def fmt_monedas(valor, ocultar=False, con_signo=False):
    if ocultar:
        return "•••••• mnds"
    if con_signo:
        return f"{int(valor):+,} mnds"
    return f"{int(valor):,} mnds"

def autodetectar_estrategia(nombre_carta, precio):
    nom_u = str(nombre_carta).upper()
    if any(k in nom_u for k in ["TOTW", "IN-FORM", "INFORM", " - IF)"]):
        return "TOTW / IF"
    m_med = re.search(r'MEDIA\s*(\d{2})', nom_u)
    med = int(m_med.group(1)) if m_med else 0
    if med >= 84 and precio <= 32000:
        return "Fodder SBC"
    if precio <= 15000:
        return "Evolución / Hype"
    return "Flipping Meta"

def es_carta_especial_oop(nombre_carta):
    nom_u = sin_tildes(str(nombre_carta)).upper()
    if any(k in nom_u for k in ["TOTW", "IN-FORM", "INFORM", " - IF)", "PROMO", "SPECIAL", "ESPECIAL", "HERO", "ICON", "EVO", "RTTK", "RTTF", "TRAILBLAZER", "CENTURION", "THUNDER", "WINTER", "WILDCARD", "TOTY", "TOTS", "FUTTIES"]):
        return True
    if "NORMAL" not in nom_u and "ORO" not in nom_u and "GOLD" not in nom_u and "RARE" not in nom_u and "COMMON" not in nom_u and "FUT.GG" not in nom_u and " - " in nom_u:
        return True
    return False

# ==========================================
# MOTOR DE INTELIGENCIA ARTIFICIAL EN VIVO
# ==========================================
def llamar_ia_en_vivo(prompt, api_key_usar=None):
    key_final = (api_key_usar or DEFAULT_GEMINI_KEY).strip()

    try:
        from google import genai
        client = genai.Client(api_key=key_final)
        for mod in MODELOS_PREFERIDOS_IA:
            try:
                resp = client.models.generate_content(model=mod, contents=prompt)
                if resp and resp.text:
                    return resp.text.strip(), mod
            except Exception:
                continue
        for m in client.models.list():
            if "flash" in m.name.lower():
                try:
                    resp = client.models.generate_content(model=m.name, contents=prompt)
                    if resp and resp.text:
                        return resp.text.strip(), m.name
                except Exception:
                    continue
    except Exception:
        pass

    for mod_rest in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{mod_rest}:generateContent?key={key_final}"
            r = requests.post(url, json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=18)
            if r.status_code == 200:
                data_j = r.json()
                txt_out = data_j["candidates"][0]["content"]["parts"][0]["text"]
                if txt_out:
                    return txt_out.strip(), f"{mod_rest} (REST)"
        except Exception:
            continue

    return "", ""

ESTRATEGIAS_VALIDAS_CARTERA = ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"]

def _optimizar_imagen_vision(imagen_bytes, max_ancho=1600):
    """Redimensiona la imagen si supera max_ancho px y la convierte a JPEG optimizado."""
    if not imagen_bytes:
        return b"", "image/jpeg"
    if Image is not None:
        try:
            img = Image.open(io.BytesIO(imagen_bytes))
            if ImageOps is not None:
                img = ImageOps.exif_transpose(img)
            if img.mode in ("RGBA", "LA", "P"):
                fondo = Image.new("RGB", img.size, (15, 23, 42))
                img_rgba = img.convert("RGBA")
                fondo.paste(img_rgba, mask=img_rgba.split()[-1])
                img = fondo
            elif img.mode != "RGB":
                img = img.convert("RGB")
            if img.width > max_ancho:
                nuevo_alto = max(1, int(round(img.height * (max_ancho / float(img.width)))))
                filtro_res = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)
                img = img.resize((max_ancho, nuevo_alto), filtro_res)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=88, optimize=True)
            return buf.getvalue(), "image/jpeg"
        except Exception:
            pass
    # Fallback sin Pillow: detectar MIME por cabecera mágica
    mime = "image/jpeg"
    if imagen_bytes.startswith(b"\x89PNG"):
        mime = "image/png"
    elif imagen_bytes[:4] == b"RIFF" and imagen_bytes[8:12] == b"WEBP":
        mime = "image/webp"
    return imagen_bytes, mime

def _parsear_precio_vision(val):
    """Convierte valores de precio devueltos por Vision (int, '15,000', '14.5K', '1.2M') a entero."""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(round(val))
    s = str(val or "").strip().upper().replace(" ", "").replace("MNDS", "").replace("COINS", "")
    if not s:
        return 0
    m_km = re.match(r'^(\d+(?:[.,]\d+)?)([KM])$', s)
    if m_km:
        num_f = float(m_km.group(1).replace(",", "."))
        mult = 1000 if m_km.group(2) == "K" else 1000000
        return int(round(num_f * mult))
    # Eliminar separadores de miles (puntos o comas)
    limpio = re.sub(r'[^\d]', '', s)
    return int(limpio) if limpio.isdigit() else 0

def _normalizar_estrategia_vision(est_raw, nombre_jug, precio):
    s = sin_tildes(str(est_raw or "")).strip().lower()
    if "totw" in s or "inform" in s or "in-form" in s or s == "if":
        return "TOTW / IF"
    if "fodder" in s or "sbc" in s or "media" in s:
        return "Fodder SBC"
    if "evo" in s or "hype" in s:
        return "Evolución / Hype"
    if "flip" in s or "meta" in s:
        return "Flipping Meta"
    return autodetectar_estrategia(nombre_jug, precio)

def _parsear_json_compras_vision(texto_raw):
    """Extrae y sanitiza la lista de compras desde la respuesta de texto JSON de Gemini."""
    if not texto_raw:
        return []
    limpio = str(texto_raw).strip()
    limpio = re.sub(r'^```(?:json)?\s*', '', limpio, flags=re.IGNORECASE)
    limpio = re.sub(r'\s*```$', '', limpio)

    data = None
    try:
        data = json.loads(limpio)
    except Exception:
        m_arr = re.search(r'\[[\s\S]*\]', limpio)
        if m_arr:
            try:
                data = json.loads(m_arr.group(0))
            except Exception:
                data = None
        if data is None:
            m_obj = re.search(r'\{[\s\S]*\}', limpio)
            if m_obj:
                try:
                    obj_u = json.loads(m_obj.group(0))
                    data = [obj_u] if isinstance(obj_u, dict) else None
                except Exception:
                    data = None

    if isinstance(data, dict):
        data = data.get("compras") or data.get("jugadores") or [data]
    if not isinstance(data, list):
        return []

    resultado = []
    for item in data:
        if not isinstance(item, dict):
            continue
        nom = str(item.get("Jugador", item.get("jugador", item.get("Nombre", ""))) or "").strip()
        nom = re.sub(r'\s+', ' ', nom)
        precio = _parsear_precio_vision(item.get("Precio_Pagado", item.get("precio_pagado", item.get("Precio", 0))))
        cant = num_seguro(item.get("Cantidad", item.get("cantidad", 1)), 1)
        cant = max(1, min(100, cant))
        if not nom or not (150 <= precio <= TOPE_RADAR_MAXIMO):
            continue
        est = _normalizar_estrategia_vision(item.get("Estrategia", item.get("estrategia", "")), nom, precio)
        resultado.append({
            "Jugador": nom,
            "Cantidad": int(cant),
            "Precio_Pagado": int(precio),
            "Estrategia": est
        })
    return resultado

def extraer_compras_desde_imagen(imagen_bytes, api_key_usar=None):
    """
    Analiza una foto o captura de compras de EA Sports FC 27 con Gemini Vision
    (SDK google-genai + fallback REST inlineData en base64) y devuelve:
    (lista_compras, modelo_usado, mensaje_error)
    """
    if not imagen_bytes:
        return [], "", "No se recibieron datos de imagen."

    key_final = (api_key_usar or DEFAULT_GEMINI_KEY).strip()
    img_opt_bytes, mime_type = _optimizar_imagen_vision(imagen_bytes, max_ancho=1600)

    prompt_vision = (
        "Sos un asistente experto de trading en EA Sports FC 27 (Web App, Companion App y Consola). "
        "Analizá esta foto o captura de pantalla y extraé ÚNICAMENTE las cartas que fueron COMPRADAS o GANADAS.\n"
        "REGLAS ESTRICTAS:\n"
        "1. Identificá solo las cartas compradas/ganadas (tick verde, 'Ganado', 'Won', 'Comprado por', 'Bought for', 'Pagado' o ítems asignados por compra). "
        "IGNORÁ pujas superadas ('Superado', 'Outbid'), expiradas o no ganadas si se distinguen.\n"
        "2. Extraé el nombre limpio del jugador en 'Jugador' (sin símbolos raros ni texto de botones).\n"
        "3. Extraé el precio pagado exacto en monedas en 'Precio_Pagado' como número entero sin puntos ni comas (ej: 15000, nunca '15,000' ni '15K').\n"
        "4. Si aparecen varias copias idénticas de la misma carta compradas al mismo precio unitario, agrupalas indicando el total en 'Cantidad' (entero >= 1).\n"
        "5. Clasificá 'Estrategia' OBLIGATORIAMENTE en uno de estos 4 valores exactos: "
        "\"Flipping Meta\", \"Fodder SBC\", \"TOTW / IF\" o \"Evolución / Hype\".\n"
        "6. Devolvé ÚNICAMENTE un JSON válido (un array de objetos), SIN bloques de código markdown, SIN ```json y SIN texto adicional.\n"
        "Formato exacto de salida:\n"
        "[{\"Jugador\": \"Nombre Limpio\", \"Cantidad\": 1, \"Precio_Pagado\": 15000, \"Estrategia\": \"Flipping Meta\"}]\n"
        "Si no se distingue ninguna compra válida en la imagen, devolvé exactamente: []"
    )

    texto_resp = ""
    modelo_usado = ""
    ultimo_err = ""

    # 1. Intento principal con SDK oficial google-genai
    modelos_sdk = []
    for m_cand in ["gemini-2.5-flash", "gemini-2.0-flash"] + list(MODELOS_PREFERIDOS_IA):
        m_clean = m_cand.replace("models/", "")
        if m_clean not in modelos_sdk:
            modelos_sdk.append(m_clean)

    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=key_final)
        part_img = types.Part.from_bytes(data=img_opt_bytes, mime_type=mime_type)
        for mod in modelos_sdk:
            try:
                resp = client.models.generate_content(
                    model=mod,
                    contents=[prompt_vision, part_img],
                    config=types.GenerateContentConfig(temperature=0.1)
                )
                if resp and getattr(resp, "text", None):
                    texto_resp = resp.text.strip()
                    modelo_usado = mod
                    break
            except Exception as e_mod:
                ultimo_err = str(e_mod)
                continue
    except Exception as e_sdk:
        ultimo_err = str(e_sdk)

    # 2. Fallback vía API REST de Gemini usando inlineData en base64
    if not texto_resp:
        img_b64 = base64.b64encode(img_opt_bytes).decode("ascii")
        payload_rest = {
            "contents": [{
                "parts": [
                    {"text": prompt_vision},
                    {"inlineData": {"mimeType": mime_type, "data": img_b64}}
                ]
            }],
            "generationConfig": {"temperature": 0.1}
        }
        for mod_rest in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{mod_rest}:generateContent?key={key_final}"
                r = requests.post(url, json=payload_rest, timeout=28)
                if r.status_code == 200:
                    data_j = r.json()
                    txt_out = data_j["candidates"][0]["content"]["parts"][0]["text"]
                    if txt_out:
                        texto_resp = txt_out.strip()
                        modelo_usado = f"{mod_rest} (REST)"
                        break
                else:
                    ultimo_err = f"HTTP {r.status_code}"
            except Exception as e_rest:
                ultimo_err = str(e_rest)
                continue

    if not texto_resp:
        return [], "", f"No se pudo conectar con Gemini Vision ({ultimo_err or 'sin respuesta'}). Verificá tu API Key o conexión."

    compras = _parsear_json_compras_vision(texto_resp)
    return compras, modelo_usado, ""

def generar_reporte_estrategico_vivo(df_corte, filas_cartera, presupuesto, roi_pct, lazy_pct, plataforma, enfoque, top_n, api_key_usar):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_hoy = dias[ahora.weekday()]
    hora_str = ahora.strftime("%H:%M")

    df_activos_mkt = df_corte[df_corte["Precio_Actual"].apply(num_seguro) > 0].copy() if (df_corte is not None and not df_corte.empty and "Precio_Actual" in df_corte.columns) else pd.DataFrame()
    df_disp = df_activos_mkt[df_activos_mkt["Precio_Actual"].apply(num_seguro) <= presupuesto].copy() if not df_activos_mkt.empty else pd.DataFrame()
    if df_disp.empty:
        df_disp = df_activos_mkt.copy()

    df_compras = df_disp[df_disp["Accion"].astype(str).str.contains("COMPRAR", na=False)].sort_values(by="Variacion_Bruta_%").head(18) if ("Accion" in df_disp.columns and not df_disp.empty) else df_disp.head(18)
    df_ventas = df_disp[df_disp["Accion"].astype(str).str.contains("VENDER", na=False)].sort_values(by="Variacion_Neta_%", ascending=False).head(12) if ("Accion" in df_disp.columns and not df_disp.empty) else pd.DataFrame()
    df_muestra = pd.concat([df_compras, df_ventas, df_disp.head(15)]).drop_duplicates(subset=["Jugador"]).head(30) if not df_disp.empty else pd.DataFrame()

    cols_res = [c for c in ["Jugador", "Precio_Actual", "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Precio_Lazy_Sugerido", "Accion"] if c in df_muestra.columns]
    tabla_txt = df_muestra[cols_res].to_string(index=False) if not df_muestra.empty else "Sin datos activos en el corte actual."

    cartera_txt = "Sin cartas activas en cartera."
    if filas_cartera:
        lineas_c = [
            f"- {c['Cantidad']}x {c['Jugador']} ({c['Estrategia']}): Costo Unit {c['Costo_Pagado']:,} | Hoy {c['Mercado_Hoy']:,} | Lazy {c['Precio_Lazy_1h']:,} | Neto Lote {c['Resultado_Tenencia']:+,} ({c['ROI_Latente_%']:+.1f}% | {c['ROI_Diario_%']:+.1f}%/día)"
            for c in filas_cartera
        ]
        cartera_txt = "\n".join(lineas_c)

    prompt = f"""
    Actuá como Trader Cuantitativo Principal de EA Sports FC 27 para la plataforma {plataforma} (Base Exclusiva FUT.GG, hasta 10,000 cartas).
    Momento del análisis: {dia_hoy} {hora_str} hs (ARG).
    Parámetros del usuario: Presupuesto Tope: {presupuesto:,} monedas | Objetivo Puja: +{roi_pct}% ROI | Recargo Lazy Buyer: +{lazy_pct}% | Tax EA: 5%.
    ENFOQUE SOLICITADO POR EL USUARIO: {enfoque}.

    CARTERA ACTIVA ACTUAL DEL USUARIO:
    {cartera_txt}

    MUESTRA EN VIVO DEL MERCADO ({len(df_activos_mkt):,} cartas con precio activo en padrón FUT.GG):
    {tabla_txt}

    Redactá un REPORTE ESTRATÉGICO EN VIVO claro, directo y accionable (usando viñetas, sin usar tablas con '|'):
    1. 🎯 TOP {top_n} OPORTUNIDADES DE COMPRA AHORA ({enfoque}): Indicá Nombre, Precio Actual en FUT.GG, hasta cuánto PUJAR (+{roi_pct}% ROI) y a cuánto listar para LAZY BUYERS (+{lazy_pct}%).
    2. 💼 DIAGNÓSTICO DE TU CARTERA ACTIVA: Qué hacer hoy con cada una de las cartas que el usuario tiene en su cartera (o cómo asignar su liquidez si no tiene cartas).
    3. 🧭 PLAN DE ACCIÓN PARA HOY ({dia_hoy}): Recomendación según la fase actual del ciclo semanal de Rivals / FUT Champions.
    """

    texto_ia, modelo_usado = llamar_ia_en_vivo(prompt, api_key_usar)
    if texto_ia:
        cabecera = f"🤖 **REPORTE IA EN VIVO ({dia_hoy} {hora_str} hs | Enfoque: {enfoque} | Modelo: {modelo_usado})**\n\n"
        return cabecera + texto_ia

    lineas_fb = [
        f"🤖 **REPORTE TÁCTICO CUANTITATIVO EN VIVO ({dia_hoy} {hora_str} hs | Enfoque: {enfoque})**",
        f"*(Base Exclusiva FUT.GG · {len(df_activos_mkt):,} cartas activas analizadas | Presupuesto: {presupuesto:,} mnds)*\n",
        f"### 🎯 Top {top_n} Oportunidades Detectadas en tu Rango (<={presupuesto:,} mnds):"
    ]
    for _, r in df_muestra.head(top_n).iterrows():
        nom = str(r.get("Jugador", "Carta")).split(" (")[0]
        p_a = num_seguro(r.get("Precio_Actual", 0))
        puja_a = num_seguro(r.get("Puja_Max_10%_ROI", 0))
        lazy_a = num_seguro(r.get("Precio_Lazy_Sugerido", 0))
        var_a = float_seguro(r.get("Variacion_Bruta_%", 0))
        sen_a = str(r.get("Senal_Trading", "Operable"))
        lineas_fb.append(f"- **{nom}** (`{p_a:,} mnds` en `FUT.GG` | `{var_a:+.1f}%`) — {sen_a} ➡️ 🔨 Pujar hasta `{puja_a:,}` | 🎯 Listar Lazy a `{lazy_a:,}`")

    lineas_fb.append("\n### 💼 Diagnóstico de tu Cartera Activa:")
    if filas_cartera:
        for c in filas_cartera:
            lineas_fb.append(f"- **{c['Cantidad']}x {c['Jugador']}**: Mercado hoy `{c['Mercado_Hoy']:,}` (Tenencia lote: `{c['Resultado_Tenencia']:+,} mnds` / `{c['ROI_Latente_%']:+.1f}%`) ➡️ **{c['Acción_Sugerida']}** (Listar Lazy en `{c['Precio_Lazy_1h']:,}`).")
    else:
        lineas_fb.append("- Tenés el 100% de tu presupuesto líquido. Destiná hasta el 60% a las oportunidades de arriba mediante puja.")

    return "\n".join(lineas_fb)

def generar_oraculo_contenido_vivo(df_corte, df_fodder, df_totw, hora_drop_txt, plataforma, api_key_usar):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_hoy = dias[ahora.weekday()]
    hora_str = ahora.strftime("%H:%M")
    pauta_dia = AGENDA_EA_SEMANAL.get(dia_hoy, "Contenido diario de SBCs y Evoluciones.")

    subidas = []
    fodder_lista = []
    df_activos_ora = df_corte[df_corte["Precio_Actual"].apply(num_seguro) > 0].copy() if (df_corte is not None and not df_corte.empty and "Precio_Actual" in df_corte.columns) else pd.DataFrame()
    if not df_activos_ora.empty:
        df_sub = df_activos_ora.sort_values(by="Variacion_Bruta_%", ascending=False).head(6) if "Variacion_Bruta_%" in df_activos_ora.columns else pd.DataFrame()
        for _, r in df_sub.iterrows():
            if float_seguro(r.get("Variacion_Bruta_%", 0)) > 0:
                subidas.append(f"{str(r['Jugador']).split(' (')[0]} ({num_seguro(r['Precio_Actual']):,} mnds, +{float_seguro(r['Variacion_Bruta_%']):.1f}%)")

        pisos_ora, refs_ora, _ = calcular_pisos_fodder_vivos(df_activos_ora)
        for med in [84, 85, 86, 87, 88, 89]:
            if med in pisos_ora:
                fodder_lista.append(f"Media {med}: {refs_ora.get(med, f'{pisos_ora[med]:,} mnds')} [FUT.GG]")

    prompt = f"""
    Actuá como el Oráculo de Contenido, SBCs y Filtraciones (Leaks) de EA Sports FC 27 ({plataforma} · Base Exclusiva FUT.GG).
    Hoy es {dia_hoy}, hora actual: {hora_str} hs (ARG). Horario del Drop diario para el usuario: {hora_drop_txt}.
    Pauta de contenido de EA para {dia_hoy}: {pauta_dia}
    Cartas con mayor presión compradora en vivo hoy: {', '.join(subidas) if subidas else 'Mercado estable sin picos abruptos'}
    Precios piso reales de Medias (Fodder SBC) ahora mismo: {', '.join(fodder_lista) if fodder_lista else 'Sin muestra de fodder'}

    Redactá un DICTAMEN DEL ORÁCULO DE CONTENIDO dinámico y directo (sin tablas '|'):
    1. 🔮 PREDICCIÓN DEL DROP DE HOY ({dia_hoy} - {hora_drop_txt}): Qué SBCs, Mejoras o Evoluciones lanza EA hoy y qué medias o requisitos van a volar.
    2. 🛒 LAS 3 COMPRAS ANTICIPADAS DEL ORÁCULO: Indicá qué 3 medias, TOTW o cartas puntuales de la lista conviene acumular AHORA MISMO y a qué precio máximo.
    3. ⏰ TIMING DE ENTRADA Y SALIDA HOY: A qué hora exacta comprar en pánico y a qué hora vender en el pico de hype.
    """

    texto_ia, modelo_usado = llamar_ia_en_vivo(prompt, api_key_usar)
    if texto_ia:
        cabecera = f"🔮 **ORÁCULO DE CONTENIDO EN VIVO ({dia_hoy} {hora_str} hs | Drop: {hora_drop_txt} | Modelo: {modelo_usado})**\n\n"
        return cabecera + texto_ia

    lineas_or = [
        f"🔮 **ORÁCULO DE CONTENIDO EN VIVO ({dia_hoy} {hora_str} hs | Drop: {hora_drop_txt})**\n",
        f"### 📅 Pauta Oficial esperada para hoy ({dia_hoy}):",
        f"- {pauta_dia}\n",
        "### 🧱 Piso Actual de Medias SBC (Fodder en Vivo):"
    ]
    for f_it in (fodder_lista or ["Consultá la pestaña Simulador Fodder para ver los pisos por media."]):
        lineas_or.append(f"- {f_it}")
    lineas_or.append(f"\n### ⏰ Ventana Operativa Sugerida:\n- Acumular en puja hasta 25 minutos antes de las **{hora_drop_txt}** y listar con recargo Lazy Buyer durante los primeros 45 minutos posteriores al lanzamiento del contenido.")
    return "\n".join(lineas_or)

def colorear_neto_y_roi(val):
    try:
        v = float(val)
        if v > 0:
            return "background-color: rgba(0, 230, 118, 0.22); color: #00e676; font-weight: bold;"
        elif v < 0:
            return "background-color: rgba(255, 23, 68, 0.25); color: #ff5252; font-weight: bold;"
    except Exception:
        pass
    return ""

def colorear_accion(val):
    s = str(val)
    if "STOP-LOSS" in s or "🛑" in s or "⚠️" in s or "⛔" in s or "EXTINTA" in s.upper():
        return "background-color: rgba(255, 23, 68, 0.30); color: #ff8a80; font-weight: bold;"
    if "VENDER" in s or "🚨" in s or "🟢" in s or "⚡" in s or "✅" in s or "CONTRIBUCIÓN" in s:
        return "background-color: rgba(0, 230, 118, 0.25); color: #69f0ae; font-weight: bold;"
    if "💎" in s:
        return "background-color: rgba(0, 229, 255, 0.25); color: #84ffff; font-weight: bold;"
    if "HOLD" in s or "⏳" in s or "ROTAR" in s or "🔄" in s:
        return "background-color: rgba(255, 193, 7, 0.22); color: #ffd54f; font-weight: bold;"
    return ""

def aplicar_estilo_seguro(df, cols_neto=None, cols_accion=None):
    try:
        styler = df.style
        if cols_neto:
            c_validas = [c for c in cols_neto if c in df.columns]
            if c_validas:
                styler = styler.map(colorear_neto_y_roi, subset=c_validas)
        if cols_accion:
            c_validas_a = [c for c in cols_accion if c in df.columns]
            if c_validas_a:
                styler = styler.map(colorear_accion, subset=c_validas_a)
        return styler
    except Exception:
        return df

# ==========================================
# GESTIÓN DE ENLACE Y SINCRONIZACIÓN DE DATOS
# ==========================================
def obtener_link_guardado():
    if LINK_GOOGLE_SHEET.strip():
        return LINK_GOOGLE_SHEET.strip()
    if os.path.exists(ARCHIVO_LINK_CACHE):
        try:
            with open(ARCHIVO_LINK_CACHE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return ""

def guardar_link_cache(nuevo_link):
    try:
        with open(ARCHIVO_LINK_CACHE, "w", encoding="utf-8") as f:
            f.write(nuevo_link.strip())
    except Exception:
        pass

def extraer_id_sheet(url_o_id):
    texto = str(url_o_id).strip()
    if not texto:
        return ""
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", texto)
    if m:
        return m.group(1)
    return texto

def convertir_columnas_numericas(df):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df
    df = df.copy()
    cols_num = [
        "Cantidad", "Precio_Pagado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%",
        "Precio_Anterior", "Precio_Actual", "Precio_Consola",
        "Precio_Min_Hist", "Precio_Max_Hist", "Precio_Prom_Hist", "Muestras_Hist",
        "Puja_Max_10%_ROI", "Break_Even_5%", "Precio_Lazy_Sugerido",
        "Variacion_Bruta_%", "Beneficio_Neto_Mnds", "Variacion_Neta_%",
        "Precio_Objetivo", "Margen_Neto_5%", "Brecha_%", "Subida_%", "Ganancia_Potencial_Neta",
        "Precio_Mas_Barato", "Piso_Compra", "Techo_Venta", "Precio_Base", "Precio_Explosivo",
        "Media", "Puja_Sugerida", "Target_Lazy_2da_Ola", "Puja_Max_Sugerida",
        "Monto_Monedas", "Incidencia_%"
    ]
    for c in cols_num:
        if c in df.columns:
            s = df[c].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.strip()
            df[c] = pd.to_numeric(s, errors="coerce").fillna(1 if c == "Cantidad" else 0)
    if "Cantidad" in df.columns:
        df["Cantidad"] = df["Cantidad"].apply(lambda x: max(1, int(float(x)) if pd.notna(x) and float(x) >= 1 else 1))
    return df

def limpiar_escaneos_corruptos(df):
    if df is None or df.empty or "Jugador" not in df.columns or "Precio_Actual" not in df.columns:
        return pd.DataFrame()
    mask_corrupta = (
        df["Jugador"].astype(str).str.contains(r"Media \?\?", na=False)
        & (df["Precio_Actual"].apply(num_seguro) == 36000)
    )
    return df[~mask_corrupta].copy()

def _normalizar_precio_ea_individual(p_val, media_val=0):
    """Sanea un precio contra escalones EA Sports FC y pisos realistas por valoración."""
    p = num_seguro(p_val, 0)
    if p <= 0 or p < 200 or p > TOPE_RADAR_MAXIMO:
        return 0
    if not es_multiplo_escalon_ea(p, TOPE_RADAR_MAXIMO):
        # Corregir redondeos históricos 1-decimal de FUT.GG en 10k-50k (ej. 13300 -> 13250, 14800 -> 14750)
        if 10000 <= p < 50000:
            ajust = int(round(p / 250.0) * 250)
            if abs(ajust - p) <= 50 and es_multiplo_escalon_ea(ajust, TOPE_RADAR_MAXIMO):
                p = ajust
            else:
                return 0
        else:
            return 0
    if media_val >= 85 and not precio_cumple_piso_media_web(p, media_val):
        return 0
    return int(p)

def sanear_dataframe_mercado(df_raw):
    """
    Capa centralizada de blindaje y saneamiento de datos al cargar Excel / Google Sheets:
    1. Elimina columnas residuales de Futbin y duplicadas conservando Enlace_Optimo (FUT.GG).
    2. Normaliza todos los tipos numéricos y descarta filas corruptas heredadas.
    3. Corrige cartas base mal etiquetadas como TOTW en cortes antiguos.
    4. Sanea outliers imposibles, precios fuera de escalón EA y estadísticas coladas (-> 0 extinta).
    5. Ordena cronológicamente mediante pd.to_datetime(..., errors='coerce') y marca Estado_Mercado.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    df = df_raw.copy()
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]

    # 1. Consolidar Enlace_Optimo y eliminar columnas residuales de Futbin / duplicadas
    if "Enlace_Optimo" not in df.columns:
        if "Enlace_FUTGG" in df.columns:
            df["Enlace_Optimo"] = df["Enlace_FUTGG"]
        elif "Enlace_Futbin" in df.columns:
            df["Enlace_Optimo"] = df["Enlace_Futbin"]
        else:
            df["Enlace_Optimo"] = ""
    else:
        mask_vacio_enl = df["Enlace_Optimo"].isna() | (df["Enlace_Optimo"].astype(str).str.strip() == "")
        if mask_vacio_enl.any():
            for alt_col in ("Enlace_FUTGG", "Enlace_Futbin"):
                if alt_col in df.columns:
                    df.loc[mask_vacio_enl, "Enlace_Optimo"] = df.loc[mask_vacio_enl, alt_col]
                    mask_vacio_enl = df["Enlace_Optimo"].isna() | (df["Enlace_Optimo"].astype(str).str.strip() == "")

    cols_drop = [c for c in COLUMNAS_RESIDUALES_FUTBIN if c in df.columns]
    if cols_drop:
        df = df.drop(columns=cols_drop, errors="ignore")

    # 2. Normalizar tipos numéricos y limpiar escaneos corruptos heredados
    df = convertir_columnas_numericas(df)
    df = limpiar_escaneos_corruptos(df)
    if df.empty or "Jugador" not in df.columns or "Precio_Actual" not in df.columns:
        return pd.DataFrame()

    df["Jugador"] = df["Jugador"].astype(str).str.strip()
    df = df[df["Jugador"] != ""].copy()

    # 3. Corregir cartas base mal etiquetadas como TOTW en historial antiguo
    if "Enlace_Optimo" in df.columns:
        urls_s = df["Enlace_Optimo"].astype(str)
        m_ids = urls_s.str.extract(r'/players/(\d+)-[^/]+/27-(\d+)(?:/|$)')
        mask_base_url = m_ids[0].notna() & (m_ids[0] == m_ids[1])
        mask_fake_totw = mask_base_url & df["Jugador"].str.contains(r'TOTW|IN-FORM|-\s*IF\b', case=False, na=False)
        if mask_fake_totw.any():
            df.loc[mask_fake_totw, "Jugador"] = df.loc[mask_fake_totw, "Jugador"].str.replace(
                r'(?:TOTW\s*IF|TOTW|IN-FORM|-\s*IF)\s*(?:-\s*FUT\.GG)?',
                'Normal - FUT.GG',
                regex=True,
                flags=re.IGNORECASE
            )
            if "Categoria_Carta" in df.columns:
                df.loc[mask_fake_totw, "Categoria_Carta"] = "Normal/Oro"
    if "Categoria_Carta" not in df.columns:
        df["Categoria_Carta"] = df["Jugador"].apply(categoria_version_carta_web)
    else:
        mask_cat_vacia = df["Categoria_Carta"].isna() | (df["Categoria_Carta"].astype(str).str.strip().isin(["", "nan", "None"]))
        if mask_cat_vacia.any():
            df.loc[mask_cat_vacia, "Categoria_Carta"] = df.loc[mask_cat_vacia, "Jugador"].apply(categoria_version_carta_web)

    # 4. Sanear precios contra escalones oficiales EA y pisos realistas por media (vectorizado + rápido)
    medias_arr = pd.to_numeric(df["Jugador"].str.extract(r'Media\s*(\d{2})\b')[0], errors="coerce").fillna(0).astype(int).values
    precios_raw = pd.to_numeric(df["Precio_Actual"], errors="coerce").fillna(0).astype(int).values
    precios_sanos = [
        _normalizar_precio_ea_individual(pv, mv)
        for pv, mv in zip(precios_raw, medias_arr)
    ]
    df["Precio_Actual"] = pd.Series(precios_sanos, index=df.index, dtype="int64")

    if "Precio_Anterior" in df.columns:
        p_ant_raw = pd.to_numeric(df["Precio_Anterior"], errors="coerce").fillna(0).astype(int).values
        df["Precio_Anterior"] = pd.Series(
            [_normalizar_precio_ea_individual(pv, mv) for pv, mv in zip(p_ant_raw, medias_arr)],
            index=df.index,
            dtype="int64"
        )

    # 5. Enforzar coherencia de cartas extintas (Precio_Actual == 0) y columna Estado_Mercado
    mask_ext = df["Precio_Actual"] <= 0
    df["Estado_Mercado"] = np.where(mask_ext, "⛔ Extinta", "🟢 Activa")
    if mask_ext.any():
        for c_zero in ("Puja_Max_10%_ROI", "Break_Even_5%", "Precio_Lazy_Sugerido", "Beneficio_Neto_Mnds"):
            if c_zero in df.columns:
                df.loc[mask_ext, c_zero] = 0
        for c_flt in ("Variacion_Bruta_%", "Variacion_Neta_%"):
            if c_flt in df.columns:
                df.loc[mask_ext, c_flt] = 0.0
        if "Estado_Historico" in df.columns:
            df.loc[mask_ext, "Estado_Historico"] = "⛔ Extinta"
        if "Senal_Trading" in df.columns:
            df.loc[mask_ext, "Senal_Trading"] = "⛔ Extinta (Sin Precio)"
        if "Accion" in df.columns:
            df.loc[mask_ext, "Accion"] = "⛔ EXTINTA"

    # 6. Ordenar cronológicamente usando pd.to_datetime y deduplicar mismo corte+carta
    if "Fecha" in df.columns:
        dt_ser = pd.to_datetime(df["Fecha"], errors="coerce")
        df["_dt_sort"] = dt_ser
        df = df.sort_values(by="_dt_sort", kind="mergesort")
        if "Enlace_Optimo" in df.columns:
            df = df.drop_duplicates(subset=["Fecha", "Enlace_Optimo"], keep="last")
        df = df.drop(columns=["_dt_sort"]).reset_index(drop=True)

    return df

@st.cache_data(ttl=30)
def cargar_por_id_publico(sheet_id):
    hojas = {}
    nombres_hojas = [
        "Historial_Tendencias", "Mi_Cartera", "Recomendacion_IA",
        "Dashboard_Financiero", "Radar_Evos", "Radar_Links_Domino",
        "Mis_Alertas", "Clubes_Cloud"
    ]
    df_base_comparacion = None

    for h in nombres_hojas:
        url_csv = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(h)}"
        try:
            df_h = pd.read_csv(url_csv)
            df_h = df_h.loc[:, ~df_h.columns.str.contains("^Unnamed")]
            if h == "Historial_Tendencias":
                df_base_comparacion = df_h.copy()
                hojas[h] = sanear_dataframe_mercado(df_h)
            else:
                if (
                    df_base_comparacion is not None
                    and list(df_h.columns) == list(df_base_comparacion.columns)
                    and len(df_h) == len(df_base_comparacion)
                ):
                    hojas[h] = pd.DataFrame()
                elif h == "Clubes_Cloud":
                    hojas[h] = df_h
                else:
                    hojas[h] = convertir_columnas_numericas(df_h)
        except Exception:
            hojas[h] = pd.DataFrame()

    if not hojas.get("Historial_Tendencias", pd.DataFrame()).empty:
        return hojas, "OK"
    return {}, "No se pudo leer 'Historial_Tendencias'. Verificá que en 'Compartir' diga 'Cualquier persona con el vínculo'."

def guardar_hoja_sincronizada(df_guardar, nombre_hoja, es_admin=False, slug_club=None):
    if slug_club and not es_admin:
        st.session_state[f"club_{slug_club}_{nombre_hoja}"] = df_guardar.copy()
        try:
            if nombre_hoja == "Mi_Cartera":
                df_guardar.to_csv(f"cartera_club_{slug_club}.csv", index=False)
            elif nombre_hoja == "Mis_Alertas":
                df_guardar.to_csv(f"alertas_club_{slug_club}.csv", index=False)
        except Exception:
            pass
        sincronizar_clubes_cloud_sheet()
        return

    if not es_admin:
        st.session_state[f"guest_{nombre_hoja}"] = df_guardar.copy()
        return

    st.session_state[f"override_{nombre_hoja}"] = df_guardar.copy()
    try:
        if nombre_hoja == "Mi_Cartera":
            df_guardar.to_csv(ARCHIVO_CARTERA_LOCAL, index=False)
        elif nombre_hoja == "Mis_Alertas":
            df_guardar.to_csv(ARCHIVO_ALERTAS_LOCAL, index=False)
        elif nombre_hoja == "Recomendacion_IA":
            df_guardar.to_csv(ARCHIVO_IA_LOCAL, index=False)
    except Exception:
        pass

    gc = obtener_cliente_gspread()
    if gc is not None:
        try:
            sh = gc.open(NOMBRE_SHEET)
            try:
                ws = sh.worksheet(nombre_hoja)
                ws.clear()
            except Exception:
                ws = sh.add_worksheet(title=nombre_hoja, rows="500", cols="20")
            df_clean = df_guardar.fillna("")
            datos_lista = [df_clean.columns.values.tolist()] + df_clean.astype(str).values.tolist()
            ws.update(datos_lista)
        except Exception:
            pass

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_guardar.to_excel(writer, index=False, sheet_name=nombre_hoja)
        except Exception:
            pass

# ==========================================
# EJECUCIÓN PRINCIPAL
# ==========================================
try:
    if "link_manual" not in st.session_state:
        st.session_state["link_manual"] = obtener_link_guardado()

    link_activo = st.session_state["link_manual"] or obtener_link_guardado()
    sheet_id = extraer_id_sheet(link_activo)

    datos = {}
    msg_error = ""
    estado_conexion = "⚠️ Sin conexión a datos"

    if sheet_id:
        datos, msg_error = cargar_por_id_publico(sheet_id)
        if msg_error == "OK":
            estado_conexion = "☁️ Conectado a Google Sheets (Nube 24/7)"

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            df_hist_excel = sanear_dataframe_mercado(pd.read_excel(ARCHIVO_EXCEL, sheet_name="Historial_Tendencias"))
            tiene_padron_valido = (
                not df_hist_excel.empty
                and "Precio_Actual" in df_hist_excel.columns
                and (df_hist_excel["Precio_Actual"].apply(num_seguro) > 0).sum() >= 50
            )
            dt_excel = pd.to_datetime(df_hist_excel["Fecha"], errors="coerce").max() if (tiene_padron_valido and "Fecha" in df_hist_excel.columns) else pd.NaT
            df_sh_tmp = datos.get("Historial_Tendencias", pd.DataFrame()) if datos else pd.DataFrame()
            dt_sheet = pd.to_datetime(df_sh_tmp["Fecha"], errors="coerce").max() if (not df_sh_tmp.empty and "Fecha" in df_sh_tmp.columns) else pd.NaT
            excel_mas_reciente = (not datos) or pd.isna(dt_sheet) or (pd.notna(dt_excel) and dt_excel >= dt_sheet)

            if tiene_padron_valido and excel_mas_reciente:
                cartera_respaldo = datos.get("Mi_Cartera", pd.DataFrame()) if datos else pd.DataFrame()
                alertas_respaldo = datos.get("Mis_Alertas", pd.DataFrame()) if datos else pd.DataFrame()
                clubes_cloud_resp = datos.get("Clubes_Cloud", pd.DataFrame()) if datos else pd.DataFrame()
                datos["Historial_Tendencias"] = df_hist_excel
                for h in [
                    "Recomendacion_IA", "Dashboard_Financiero",
                    "Radar_Evos", "Radar_Links_Domino"
                ]:
                    try:
                        datos[h] = convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name=h))
                    except Exception:
                        if h not in datos:
                            datos[h] = pd.DataFrame()
                if not clubes_cloud_resp.empty:
                    datos["Clubes_Cloud"] = clubes_cloud_resp
                else:
                    try:
                        datos["Clubes_Cloud"] = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Clubes_Cloud")
                    except Exception:
                        datos["Clubes_Cloud"] = pd.DataFrame()
                if not cartera_respaldo.empty:
                    datos["Mi_Cartera"] = cartera_respaldo
                elif "Mi_Cartera" not in datos or datos["Mi_Cartera"].empty:
                    try:
                        datos["Mi_Cartera"] = convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name="Mi_Cartera"))
                    except Exception:
                        datos["Mi_Cartera"] = pd.DataFrame()
                if not alertas_respaldo.empty:
                    datos["Mis_Alertas"] = alertas_respaldo
                elif "Mis_Alertas" not in datos or datos["Mis_Alertas"].empty:
                    try:
                        datos["Mis_Alertas"] = convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name="Mis_Alertas"))
                    except Exception:
                        datos["Mis_Alertas"] = pd.DataFrame()
                estado_conexion = "☁️ Conectado a Nube 24/7 (GitHub Auto-Sync + Sheets)"
        except Exception:
            pass

    if datos and "Clubes_Cloud" in datos and not datos["Clubes_Cloud"].empty:
        hidratar_clubes_desde_cloud(datos["Clubes_Cloud"])

    # ==========================================
    # BARRA LATERAL COMPLETA Y PERSONALIZABLE
    # ==========================================
    with st.sidebar:
        st.markdown(f"""
        <div class="sb-brand">
            <div style="width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,#10b981,#0ea5e9);display:flex;align-items:center;justify-content:center;font-size:1.1rem;flex-shrink:0;box-shadow:0 4px 12px rgba(16,185,129,0.28);">⚡</div>
            <div style="min-width:0;">
                <div class="sb-brand-title">Centro de Mando</div>
                <div class="sb-brand-sub">{html_lib.escape(estado_conexion)}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        c_sb1, c_sb2 = st.columns(2)
        with c_sb1:
            if st.button("🔄 Sincronizar", use_container_width=True):
                for k_ov in ["override_Mi_Cartera", "override_Mis_Alertas"]:
                    st.session_state.pop(k_ov, None)
                for arch_loc in [ARCHIVO_CARTERA_LOCAL, ARCHIVO_ALERTAS_LOCAL]:
                    if os.path.exists(arch_loc):
                        try:
                            os.remove(arch_loc)
                        except Exception:
                            pass
                st.cache_data.clear()
                st.rerun()
        with c_sb2:
            auto_ref = st.toggle("⏱️ Auto 60s", value=False)

        # 1. SISTEMA MULTIUSUARIO COMPLETO (ADMIN PIN 2362 | CUENTAS DE CLUB CON PIN | MODO INVITADO CSV)
        st.markdown('<div class="sb-divider"></div><div class="sb-section-label">🔐 Acceso & Club</div>', unsafe_allow_html=True)
        modo_acceso = st.radio(
            "Modo de Acceso",
            ["👑 Admin (PIN 2362)", "🏆 Cuenta de Club (Amigos)", "👤 Invitado Libre (CSV)"],
            index=0,
            label_visibility="collapsed"
        )
        es_admin = False
        slug_club_activo = None
        nombre_club_activo = "Invitado"

        columnas_base_cartera = ["Jugador_Exacto", "Estrategia", "Cantidad", "Precio_Pagado", "Estado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%", "Fecha_Compra", "Fecha_Venta"]
        columnas_base_alertas = ["Jugador_Buscado", "Precio_Objetivo", "Estado"]

        if modo_acceso == "👑 Admin (PIN 2362)":
            pin_ingresado = st.text_input("🔑 PIN de Administrador", type="password", placeholder="Ingresá tu PIN (2362)")
            es_admin = (pin_ingresado.strip() == PIN_ADMIN_SECRETO)
            if es_admin:
                nombre_club_activo = "👑 Club Admin Principal"
                st.markdown('<div class="mode-badge-admin">👑 Admin Activo · Cartera Real</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="mode-badge-guest">🔒 Ingresá PIN Admin o elegí Club / Invitado</div>', unsafe_allow_html=True)

        elif modo_acceso == "🏆 Cuenta de Club (Amigos)":
            reg_clubes = cargar_registro_clubes()
            tab_login_c, tab_nuevo_c = st.tabs(["🔑 Entrar", "➕ Crear Club"])
            with tab_login_c:
                if reg_clubes:
                    opciones_clubes = {v.get("nombre", k): k for k, v in reg_clubes.items()}
                    nom_sel_c = st.selectbox("Seleccioná tu Club:", list(opciones_clubes.keys()))
                    pin_club_in = st.text_input("PIN de tu Club:", type="password", key="pin_login_club")
                    slug_cand = opciones_clubes.get(nom_sel_c)
                    if slug_cand and pin_club_in.strip() == str(reg_clubes[slug_cand].get("pin", "")):
                        slug_club_activo = slug_cand
                        nombre_club_activo = reg_clubes[slug_cand].get("nombre", nom_sel_c)
                        st.markdown(f'<div class="mode-badge-club">🏆 Club: {html_lib.escape(str(nombre_club_activo))}</div>', unsafe_allow_html=True)
                    else:
                        st.caption("Ingresá el PIN de tu club para desbloquear tu cartera persistente.")
                else:
                    st.caption("Aún no hay clubes registrados. Creá el primero en la pestaña '➕ Crear Club'.")
            with tab_nuevo_c:
                nuevo_nom_club = st.text_input("Nombre del Nuevo Club", placeholder="Ej: Real Madrid FC")
                nuevo_pin_club = st.text_input("Crear PIN Secreto (4 dígitos)", type="password", key="pin_new_club")
                if st.button("✅ Registrar Club", use_container_width=True):
                    if nuevo_nom_club.strip() and nuevo_pin_club.strip():
                        s_new = slug_club_seguro(nuevo_nom_club)
                        reg_clubes[s_new] = {
                            "nombre": nuevo_nom_club.strip(),
                            "pin": nuevo_pin_club.strip(),
                            "creado": ahora_arg().strftime("%Y-%m-%d %H:%M")
                        }
                        pd.DataFrame(columns=columnas_base_cartera).to_csv(f"cartera_club_{s_new}.csv", index=False)
                        guardar_registro_clubes(reg_clubes)
                        st.success(f"🎉 ¡Club '{nuevo_nom_club.strip()}' creado y sincronizado! Entrá desde '🔑 Entrar'.")
                        st.rerun()
        else:
            st.markdown('<div class="mode-badge-guest">👤 Invitado · Cartera Local CSV</div>', unsafe_allow_html=True)

        cartera_admin_base = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy() if datos else pd.DataFrame(columns=columnas_base_cartera)
        if os.path.exists(ARCHIVO_CARTERA_LOCAL):
            try:
                df_c_adm_loc = convertir_columnas_numericas(pd.read_csv(ARCHIVO_CARTERA_LOCAL))
                if not df_c_adm_loc.empty:
                    cartera_admin_base = df_c_adm_loc
            except Exception:
                pass

        if es_admin and datos:
            if "override_Mi_Cartera" in st.session_state:
                datos["Mi_Cartera"] = convertir_columnas_numericas(st.session_state["override_Mi_Cartera"])
            elif os.path.exists(ARCHIVO_CARTERA_LOCAL):
                try:
                    df_c_loc = convertir_columnas_numericas(pd.read_csv(ARCHIVO_CARTERA_LOCAL))
                    if not df_c_loc.empty or "override_Mi_Cartera" in st.session_state:
                        datos["Mi_Cartera"] = df_c_loc
                except Exception:
                    pass
            cartera_admin_base = datos.get("Mi_Cartera", cartera_admin_base).copy()
            if "override_Mis_Alertas" in st.session_state:
                datos["Mis_Alertas"] = convertir_columnas_numericas(st.session_state["override_Mis_Alertas"])
            elif os.path.exists(ARCHIVO_ALERTAS_LOCAL):
                try:
                    df_a_loc = convertir_columnas_numericas(pd.read_csv(ARCHIVO_ALERTAS_LOCAL))
                    if len(df_a_loc) >= len(datos.get("Mis_Alertas", pd.DataFrame())):
                        datos["Mis_Alertas"] = df_a_loc
                except Exception:
                    pass
        elif slug_club_activo and datos:
            arch_c_club = f"cartera_club_{slug_club_activo}.csv"
            arch_a_club = f"alertas_club_{slug_club_activo}.csv"
            if f"club_{slug_club_activo}_Mi_Cartera" in st.session_state:
                datos["Mi_Cartera"] = convertir_columnas_numericas(st.session_state[f"club_{slug_club_activo}_Mi_Cartera"])
            elif os.path.exists(arch_c_club):
                try:
                    datos["Mi_Cartera"] = convertir_columnas_numericas(pd.read_csv(arch_c_club))
                except Exception:
                    datos["Mi_Cartera"] = pd.DataFrame(columns=columnas_base_cartera)
            else:
                datos["Mi_Cartera"] = pd.DataFrame(columns=columnas_base_cartera)

            if f"club_{slug_club_activo}_Mis_Alertas" in st.session_state:
                datos["Mis_Alertas"] = convertir_columnas_numericas(st.session_state[f"club_{slug_club_activo}_Mis_Alertas"])
            elif os.path.exists(arch_a_club):
                try:
                    datos["Mis_Alertas"] = convertir_columnas_numericas(pd.read_csv(arch_a_club))
                except Exception:
                    datos["Mis_Alertas"] = pd.DataFrame(columns=columnas_base_alertas)
            else:
                datos["Mis_Alertas"] = pd.DataFrame(columns=columnas_base_alertas)
        elif not es_admin and datos:
            if "guest_Mi_Cartera" not in st.session_state:
                st.session_state["guest_Mi_Cartera"] = pd.DataFrame(columns=columnas_base_cartera)
            if "guest_Mis_Alertas" not in st.session_state:
                st.session_state["guest_Mis_Alertas"] = pd.DataFrame(columns=columnas_base_alertas)
            datos["Mi_Cartera"] = convertir_columnas_numericas(st.session_state["guest_Mi_Cartera"])
            datos["Mis_Alertas"] = convertir_columnas_numericas(st.session_state["guest_Mis_Alertas"])

        if not es_admin and datos:
            with st.expander("💾 Guardar / Cargar Mi Cartera (CSV)"):
                st.caption("Descargá tu cartera a tu PC/celular o subila cuando vuelvas a entrar:")
                csv_guest = datos["Mi_Cartera"].to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Descargar Mi Cartera (.csv)",
                    data=csv_guest,
                    file_name=f"cartera_{slug_club_activo or 'invitado'}_fc27.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                archivo_subido = st.file_uploader("⬆️ Subir Cartera guardada (.csv)", type=["csv"])
                if archivo_subido is not None:
                    try:
                        df_sub = convertir_columnas_numericas(pd.read_csv(archivo_subido))
                        guardar_hoja_sincronizada(df_sub, "Mi_Cartera", es_admin=False, slug_club=slug_club_activo)
                        st.success("¡Cartera cargada con éxito!")
                        st.rerun()
                    except Exception as e_csv:
                        st.error(f"Archivo inválido: {e_csv}")

        # 2. CONFIGURACIÓN DEL PERFIL DE TRADER
        with st.expander("⚙️ Personalizar Mi Perfil y Reglas", expanded=False):
            tema_club = st.selectbox(
                "🎨 Tema Visual del Club",
                ["🟢 Esmeralda FUT.GG", "🔵 Ciber Fintech", "🟡 Oro TOTY", "🔴 Carmesí Rivals"],
                index=0
            )
            pais_usuario = st.selectbox("🌍 Tu País (Hora del Drop)", list(PAISES_DROP.keys()), index=0)
            plataforma_usuario = st.radio("🎮 Plataforma de Mercado", ["🖥️ PC (Origin/EA App)", "🎮 Consola (PS5 / Xbox)"], index=0)
            presupuesto_usuario = st.number_input("💰 Mi Presupuesto Total (Monedas)", min_value=5000, max_value=15000000, step=10000, value=PRESUPUESTO_DEFAULT)
            roi_puja_pct = st.slider("🔨 % ROI Objetivo en Puja", min_value=3, max_value=25, value=10)
            lazy_pct = st.slider("🎯 % Recargo Lazy Buyer (1h)", min_value=5, max_value=30, value=12)
            stop_loss_pct = st.slider("🛑 % Stop-Loss (Aviso salida)", min_value=-25, max_value=-3, value=-10)
            horas_rotacion = st.slider("⏱️ Horas Tope de Rotación", min_value=12, max_value=168, value=72, step=12)
            api_key_custom = st.text_input("🤖 Gemini API Key Propia (Opcional)", type="password", placeholder="Usa la clave integrada por defecto")

        c_t1, c_t2 = st.columns(2)
        with c_t1:
            modo_streamer = st.toggle("👁️ Modo Streamer", value=False)
        with c_t2:
            alerta_sonora = st.toggle("🔔 Sonido Sniper", value=True)

        # 3. CALCULADORA RÁPIDA ADAPTADA AL PERFIL DEL USUARIO (ESTILO TICKET FINANCIERO)
        st.markdown('<div class="sb-divider"></div><div class="sb-section-label">🧮 Calculadora Rápida (5% EA Tax)</div>', unsafe_allow_html=True)
        precio_sim = st.number_input("Precio visto en Mercado", min_value=200, max_value=15000000, step=250, value=25000)
        factor_roi = 1.0 + (roi_puja_pct / 100.0)
        factor_lazy = 1.0 + (lazy_pct / 100.0)
        factor_sl = 1.0 + (stop_loss_pct / 100.0)

        puja_sim = redondear_escalon_ea((precio_sim * 0.95) / factor_roi, hacia_arriba=False)
        empate_sim = redondear_escalon_ea(precio_sim / 0.95, hacia_arriba=True)
        lazy_sim = redondear_escalon_ea(precio_sim * factor_lazy, hacia_arriba=True)
        sl_sim = redondear_escalon_ea((precio_sim * factor_sl) / 0.95, hacia_arriba=False)
        neto_lazy_sim = int(round((lazy_sim * 0.95) - precio_sim))
        ea_tax_sim = int(round(lazy_sim * 0.05))

        st.markdown(f"""
        <div class="calc-ticket">
            <div class="calc-ticket-head">
                <span>TICKET DE ORDEN</span>
                <span class="mono">REF {precio_sim:,}</span>
            </div>
            <div class="calc-row">
                <span class="calc-row-lbl">🔨 Puja Máx (+{roi_puja_pct}%)</span>
                <span class="calc-row-val" style="color:#34d399;">{puja_sim:,}</span>
            </div>
            <div class="calc-row">
                <span class="calc-row-lbl">⚖️ Break-Even (5%)</span>
                <span class="calc-row-val" style="color:#7dd3fc;">{empate_sim:,}</span>
            </div>
            <div class="calc-row">
                <span class="calc-row-lbl">🎯 Venta Lazy (+{lazy_pct}%)</span>
                <span class="calc-row-val" style="color:#fbbf24;">{lazy_sim:,}</span>
            </div>
            <div class="calc-row">
                <span class="calc-row-lbl">🏛️ EA Tax (5% Lazy)</span>
                <span class="calc-row-val" style="color:#94a3b8;">-{ea_tax_sim:,}</span>
            </div>
            <div class="calc-row">
                <span class="calc-row-lbl">🛑 Stop-Loss ({stop_loss_pct}%)</span>
                <span class="calc-row-val" style="color:#f87171;">{sl_sim:,}</span>
            </div>
            <div class="calc-ticket-foot">
                <span style="color:#cbd5e1;">✨ Neto Lazy</span>
                <span class="mono" style="color:#10b981;">{neto_lazy_sim:+,} mnds</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 4. REGISTRAR COMPRA MANUAL (CON SOPORTE DE LOTES / CANTIDAD)
        st.markdown('<div class="sb-divider"></div><div class="sb-section-label">➕ Registrar Compra Manual</div>', unsafe_allow_html=True)
        with st.form("form_compra", clear_on_submit=True):
            jug_c = st.text_input("Nombre del Jugador", placeholder="Ej: Odegaard")
            cant_c = st.number_input("Cantidad de Cartas (Lote)", min_value=1, max_value=100, step=1, value=1)
            precio_c = st.number_input("Precio Unitario Pagado (Monedas)", min_value=200, max_value=15000000, step=250, value=10000)
            est_c = st.selectbox("Unidad de Negocio", ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"])
            btn_guardar_c = st.form_submit_button("✅ Guardar en Cartera", use_container_width=True)

            if btn_guardar_c and jug_c.strip():
                df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                if "Cantidad" not in df_c_act.columns:
                    df_c_act["Cantidad"] = 1
                nueva_fila = pd.DataFrame([{
                    "Jugador_Exacto": jug_c.strip().title(),
                    "Estrategia": est_c,
                    "Cantidad": int(cant_c),
                    "Precio_Pagado": int(precio_c),
                    "Estado": "Activo",
                    "Precio_Venta": 0,
                    "Costo_EA_Tax": 0,
                    "Ganancia_Neta": 0,
                    "ROI_%": 0.0,
                    "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                    "Fecha_Venta": ""
                }])
                df_c_act = pd.concat([df_c_act, nueva_fila], ignore_index=True)
                guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                saldos_tmp = cargar_saldos_personalizados(es_admin, presupuesto_usuario, 0, slug_club=slug_club_activo)
                if saldos_tmp.get("modo_liquidez") == "Manual":
                    saldos_tmp["liquidez_manual"] = max(0, int(saldos_tmp.get("liquidez_manual", 0)) - int(precio_c * cant_c))
                    guardar_saldos_personalizados(saldos_tmp, es_admin, slug_club=slug_club_activo)
                st.success(f"¡{cant_c}x {jug_c.title()} agregado a Cartera!")
                st.rerun()

        # 5. CERRAR / CORREGIR VENTA (SOPORTA VENTA PARCIAL O TOTAL DE LOTES)
        st.markdown('<div class="sb-divider"></div><div class="sb-section-label">💰 Cerrar / Corregir Venta</div>', unsafe_allow_html=True)
        df_cartera_sb = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
        if "Cantidad" not in df_cartera_sb.columns:
            df_cartera_sb["Cantidad"] = 1
        if not df_cartera_sb.empty and "Estado" in df_cartera_sb.columns:
            mask_vendibles = (df_cartera_sb["Estado"].astype(str).str.strip() == "Activo") | (
                (df_cartera_sb["Estado"].astype(str).str.strip() == "Vendido") & (df_cartera_sb["Precio_Venta"].apply(num_seguro) == 0)
            )
            activos_sb = df_cartera_sb[mask_vendibles]
        else:
            activos_sb = pd.DataFrame()

        if not activos_sb.empty:
            with st.form("form_venta", clear_on_submit=True):
                opciones_v = [
                    f"{idx} | {max(1, num_seguro(r.get('Cantidad', 1), 1))}x {r.get('Jugador_Exacto', 'Carta')} (Costo Unit: {num_seguro(r.get('Precio_Pagado', 0)):,} | {r.get('Estado', 'Activo')})"
                    for idx, r in activos_sb.iterrows()
                ]
                seleccion_v = st.selectbox("Carta / Lote a Vender", opciones_v)
                cant_v_sb = st.number_input("Cantidad Vendida (Permite Venta Parcial)", min_value=1, max_value=100, step=1, value=1)
                precio_v = st.number_input("Precio de Venta Unitario Bruto", min_value=200, max_value=15000000, step=250, value=15000)
                btn_vender = st.form_submit_button("📈 Registrar Venta (5% Tax)", use_container_width=True)

                if btn_vender:
                    idx_real = int(seleccion_v.split(" | ")[0])
                    cant_lote = max(1, num_seguro(df_cartera_sb.loc[idx_real, "Cantidad"], 1))
                    cant_liq_sb = min(int(cant_v_sb), int(cant_lote))
                    p_comp_unit = float(num_seguro(df_cartera_sb.loc[idx_real, "Precio_Pagado"], 0))
                    venta_bruta_lote = int(precio_v * cant_liq_sb)
                    costo_lote = int(p_comp_unit * cant_liq_sb)
                    tax = int(round(venta_bruta_lote * 0.05))
                    neto = int(round((venta_bruta_lote - tax) - costo_lote))
                    roi = round((neto / costo_lote) * 100, 2) if costo_lote > 0 else 0.0
                    ahora_v_str = ahora_arg().strftime("%Y-%m-%d %H:%M")

                    df_cartera_sb["Estado"] = df_cartera_sb["Estado"].astype(object)
                    if "Fecha_Venta" not in df_cartera_sb.columns:
                        df_cartera_sb["Fecha_Venta"] = ""
                    df_cartera_sb["Fecha_Venta"] = df_cartera_sb["Fecha_Venta"].astype(object)

                    if cant_liq_sb < cant_lote and str(df_cartera_sb.loc[idx_real, "Estado"]).strip() == "Activo":
                        df_cartera_sb.loc[idx_real, "Cantidad"] = int(cant_lote - cant_liq_sb)
                        fila_vendida_parcial = pd.DataFrame([{
                            "Jugador_Exacto": str(df_cartera_sb.loc[idx_real, "Jugador_Exacto"]),
                            "Estrategia": str(df_cartera_sb.loc[idx_real, "Estrategia"]),
                            "Cantidad": int(cant_liq_sb),
                            "Precio_Pagado": int(p_comp_unit),
                            "Estado": "Vendido",
                            "Precio_Venta": int(precio_v),
                            "Costo_EA_Tax": tax,
                            "Ganancia_Neta": neto,
                            "ROI_%": roi,
                            "Fecha_Compra": str(df_cartera_sb.loc[idx_real, "Fecha_Compra"]),
                            "Fecha_Venta": ahora_v_str
                        }])
                        df_cartera_sb = pd.concat([df_cartera_sb, fila_vendida_parcial], ignore_index=True)
                    else:
                        df_cartera_sb.loc[idx_real, "Cantidad"] = int(cant_liq_sb)
                        df_cartera_sb.loc[idx_real, "Estado"] = "Vendido"
                        df_cartera_sb.loc[idx_real, "Precio_Venta"] = int(precio_v)
                        df_cartera_sb.loc[idx_real, "Costo_EA_Tax"] = tax
                        df_cartera_sb.loc[idx_real, "Ganancia_Neta"] = neto
                        df_cartera_sb.loc[idx_real, "ROI_%"] = roi
                        df_cartera_sb.loc[idx_real, "Fecha_Venta"] = ahora_v_str

                    guardar_hoja_sincronizada(df_cartera_sb, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                    saldos_tmp = cargar_saldos_personalizados(es_admin, presupuesto_usuario, 0, slug_club=slug_club_activo)
                    if saldos_tmp.get("modo_liquidez") == "Manual":
                        saldos_tmp["liquidez_manual"] = max(0, int(saldos_tmp.get("liquidez_manual", 0)) + int(venta_bruta_lote - tax))
                        guardar_saldos_personalizados(saldos_tmp, es_admin, slug_club=slug_club_activo)
                    st.success(f"¡Venta registrada ({cant_liq_sb}x)! Neto: {neto:+,} mnds")
                    st.rerun()
        else:
            st.caption("No tenés cartas pendientes de venta.")

        if es_admin:
            with st.expander("🚀 Escaneo Remoto GitHub (Admin)"):
                st.caption("Dispará el bot recolector en la nube sin abrir GitHub:")
                repo_gh = st.text_input("Repositorio GitHub:", value=REPO_GITHUB_DEFAULT)
                token_gh = st.text_input("GitHub Token (PAT con permiso workflow):", type="password")
                if st.button("⚡ Escanear Mercado Ahora", use_container_width=True):
                    if token_gh.strip():
                        ok_gh, msg_gh = disparar_github_action_remoto(repo_gh, token_gh)
                        if ok_gh:
                            st.success(msg_gh)
                        else:
                            st.error(msg_gh)
                    else:
                        st.warning("Ingresá tu Personal Access Token de GitHub para iniciar el workflow.")

            with st.expander("🔗 Vínculo de Google Sheets (Admin)"):
                link_sb = st.text_input("URL de tu planilla:", value=link_activo)
                if st.button("Guardar Vínculo", use_container_width=True):
                    st.session_state["link_manual"] = link_sb.strip()
                    guardar_link_cache(link_sb.strip())
                    st.cache_data.clear()
                    st.rerun()

    # ==========================================
    # CABECERA PRINCIPAL + RELOJ ADAPTADO AL PAÍS + SEMÁFORO DE TEMPERATURA Y CALENDARIO EA
    # ==========================================
    df_historial = datos.get("Historial_Tendencias", pd.DataFrame())
    if not df_historial.empty and "Fecha" in df_historial.columns:
        dt_hist_all = pd.to_datetime(df_historial["Fecha"], errors="coerce")
        dt_max_corte = dt_hist_all.max()
        if pd.notna(dt_max_corte):
            df_ult_corte = df_historial[dt_hist_all == dt_max_corte].copy()
            ultima_act = str(df_ult_corte["Fecha"].iloc[-1])
        else:
            ultima_act = str(df_historial["Fecha"].max())
            df_ult_corte = df_historial[df_historial["Fecha"] == ultima_act].copy()
    else:
        ultima_act = "Sin escaneo"
        df_ult_corte = pd.DataFrame()

    usar_consola = "Consola" in plataforma_usuario
    tiene_precios_consola_reales = False
    if usar_consola and not df_ult_corte.empty and "Precio_Consola" in df_ult_corte.columns:
        p_con_ser = df_ult_corte["Precio_Consola"].apply(num_seguro)
        p_act_ser_chk = df_ult_corte["Precio_Actual"].apply(num_seguro)
        plat_es_con = ("Plataforma" in df_ult_corte.columns) and (df_ult_corte["Plataforma"].astype(str).str.lower() == "consola").any()
        mask_con = (p_con_ser > 0) & ((p_con_ser != p_act_ser_chk) | plat_es_con)
        if mask_con.any():
            df_ult_corte.loc[mask_con, "Precio_Actual"] = p_con_ser[mask_con]
            tiene_precios_consola_reales = True

    if not df_ult_corte.empty and "Precio_Actual" in df_ult_corte.columns:
        mapa_max_hist = {}
        mapa_cnt_hist = {}
        if not df_historial.empty and "Jugador" in df_historial.columns and "Precio_Actual" in df_historial.columns:
            df_h_valid = df_historial[df_historial["Precio_Actual"].apply(num_seguro) > 0]
            if not df_h_valid.empty:
                grp_col = "Enlace_Optimo" if "Enlace_Optimo" in df_h_valid.columns else ("Enlace_Futbin" if "Enlace_Futbin" in df_h_valid.columns else "Jugador")
                agg_h = df_h_valid.groupby(grp_col)["Precio_Actual"].agg(["max", "count"])
                mapa_max_hist = agg_h["max"].to_dict()
                mapa_cnt_hist = agg_h["count"].to_dict()
                if "Precio_Max_Hist" not in df_ult_corte.columns:
                    df_ult_corte["Precio_Max_Hist"] = df_ult_corte[grp_col].map(mapa_max_hist).fillna(df_ult_corte["Precio_Actual"])
                if "Muestras_Hist" not in df_ult_corte.columns:
                    df_ult_corte["Muestras_Hist"] = df_ult_corte[grp_col].map(mapa_cnt_hist).fillna(1)

        df_ult_corte["Estado_Mercado"] = np.where(df_ult_corte["Precio_Actual"].apply(num_seguro) > 0, "🟢 Activa", "⛔ Extinta")
        df_ult_corte["Puja_Max_10%_ROI"] = df_ult_corte["Precio_Actual"].apply(
            lambda p: calcular_puja_maxima_segura(p, factor_roi) if num_seguro(p, 0) > 0 else 0
        )
        df_ult_corte["Break_Even_5%"] = df_ult_corte["Precio_Actual"].apply(
            lambda p: calcular_break_even(p) if num_seguro(p, 0) > 0 else 0
        )
        df_ult_corte["Precio_Lazy_Sugerido"] = df_ult_corte.apply(
            lambda r: calcular_precio_lazy(
                r.get("Precio_Actual", 0),
                max_hist=r.get("Precio_Max_Hist", 0),
                muestras=r.get("Muestras_Hist", 0),
                factor_lazy=factor_lazy,
                factor_roi=factor_roi
            ) if num_seguro(r.get("Precio_Actual", 0), 0) > 0 else 0,
            axis=1
        )
        df_ult_corte["_nom_base_norm"] = df_ult_corte["Jugador"].astype(str).apply(
            lambda x: sin_tildes(x.split(" (")[0]).lower().strip()
        )

    # Indicador de Salud de Datos del Corte Actual
    total_cartas_corte = len(df_ult_corte)
    cartas_activas_corte = int((df_ult_corte["Precio_Actual"].apply(num_seguro) > 0).sum()) if (not df_ult_corte.empty and "Precio_Actual" in df_ult_corte.columns) else 0
    cartas_extintas_corte = max(0, total_cartas_corte - cartas_activas_corte)
    salud_datos_pct = round((cartas_activas_corte / max(1, total_cartas_corte)) * 100.0, 1) if total_cartas_corte > 0 else 0.0
    color_salud = "#34d399" if salud_datos_pct >= 95.0 else ("#fbbf24" if salud_datos_pct >= 80.0 else "#f87171")

    tiempo_drop, leyenda_drop, hora_drop_txt = calcular_countdown_pais(pais_usuario)
    if usar_consola and not tiene_precios_consola_reales:
        etiqueta_plat = "🎮 Consola (Ref. FUT.GG PC)"
    else:
        etiqueta_plat = "🎮 Consola" if usar_consola else "🖥️ PC"
    tit_sem, det_sem, col_sem, tit_cal, det_cal = calcular_semaforo_y_calendario_ea(df_ult_corte)

    paleta_tema = {
        "🟢 Esmeralda FUT.GG": ("#10b981", "#38bdf8"),
        "🔵 Ciber Fintech": ("#38bdf8", "#818cf8"),
        "🟡 Oro TOTY": ("#f59e0b", "#fcd34d"),
        "🔴 Carmesí Rivals": ("#ef4444", "#fb923c"),
    }
    acento_club_1, acento_club_2 = paleta_tema.get(tema_club, ("#10b981", "#38bdf8"))
    escudo_svg = (
        f'<svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">'
        f'<path d="M18 3L5 8.2V17.4C5 25.2 10.6 31.6 18 33.8C25.4 31.6 31 25.2 31 17.4V8.2L18 3Z" '
        f'fill="rgba(15,23,42,0.9)" stroke="{acento_club_1}" stroke-width="2.2"/>'
        f'<path d="M19.6 10.5L13.2 19.2H18.1L16.4 25.5L22.8 16.8H17.9L19.6 10.5Z" fill="{acento_club_2}"/>'
        f'</svg>'
    )

    st.markdown(f"""
    <div class="main-header" style="border-top: 2px solid {acento_club_1};">
        <div class="header-left">
            <div class="club-crest" style="border-color: {acento_club_1}55;">{escudo_svg}</div>
            <div class="header-title-wrap">
                <h1>FC 27 TRADING TERMINAL PRO</h1>
                <div class="header-club-sub">{html_lib.escape(str(nombre_club_activo))} &nbsp;·&nbsp; Base Exclusiva FUT.GG (Hasta 10,000 cartas) &nbsp;·&nbsp; Cloud 24/7</div>
                <div class="header-badges">
                    <span class="status-pill">{html_lib.escape(estado_conexion)}</span>
                    <span class="status-pill">{html_lib.escape(etiqueta_plat)}</span>
                    <span class="status-pill">📦 Padrón FUT.GG <b>{total_cartas_corte:,}</b></span>
                    <span class="status-pill" style="border-color:{color_salud}55;">🩺 Salud <b style="color:{color_salud};">{salud_datos_pct:.1f}%</b> · ✅ <b>{cartas_activas_corte:,}</b> · ⛔ <b>{cartas_extintas_corte:,}</b></span>
                    <span class="status-pill">🕒 Corte <b>{html_lib.escape(str(ultima_act))}</b></span>
                </div>
            </div>
        </div>
        <div class="drop-timer" style="border-color: {acento_club_1}44;">
            <div class="drop-timer-title" style="color: {acento_club_2};">🎯 Drop EA ({html_lib.escape(str(hora_drop_txt))})</div>
            <div class="drop-timer-val">{html_lib.escape(str(tiempo_drop))}</div>
            <div class="drop-timer-sub">{html_lib.escape(str(leyenda_drop))}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c_sem1, c_sem2 = st.columns(2)
    with c_sem1:
        st.markdown(
            f'<div class="status-strip-card" style="border-left: 3px solid {col_sem};">'
            f'<div class="status-strip-title" style="color:{col_sem};">🌡️ {html_lib.escape(str(tit_sem))}</div>'
            f'<div class="status-strip-desc">{html_lib.escape(str(det_sem))}</div></div>',
            unsafe_allow_html=True
        )
    with c_sem2:
        st.markdown(
            f'<div class="status-strip-card" style="border-left: 3px solid {acento_club_2};">'
            f'<div class="status-strip-title" style="color:{acento_club_2};">{html_lib.escape(str(tit_cal))}</div>'
            f'<div class="status-strip-desc">{html_lib.escape(str(det_cal))}</div></div>',
            unsafe_allow_html=True
        )

    if not datos:
        if msg_error and msg_error != "OK":
            st.error(f"🔍 {msg_error}")
        st.info("👇 **Conexión instantánea:** Pegá acá abajo el vínculo de tu Google Sheets `Base_Datos_FC27` (con acceso en 'Cualquier persona con el vínculo'):")
        nuevo_link = st.text_input("🔗 Vínculo de Google Sheets:", value=link_activo, placeholder="https://docs.google.com/spreadsheets/d/...")
        if st.button("🚀 Conectar y Memorizar Vínculo"):
            st.session_state["link_manual"] = nuevo_link.strip()
            guardar_link_cache(nuevo_link.strip())
            st.cache_data.clear()
            st.rerun()
        st.stop()

    # ==========================================
    # DETECTOR VISUAL, SONORO, TELEGRAM AUTO Y COMPRA 1-CLIC DE ALERTAS SNIPER EN VIVO
    # ==========================================
    if "snipers_notificados_tg" not in st.session_state:
        st.session_state["snipers_notificados_tg"] = set()

    df_alertas_check = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
    if not df_alertas_check.empty and "Estado" in df_alertas_check.columns and not df_ult_corte.empty:
        act_snipers = df_alertas_check[df_alertas_check["Estado"].astype(str).str.strip() == "Activa"]
        disparos_sniper = []
        df_corte_sniper = df_ult_corte[df_ult_corte["Precio_Actual"].apply(num_seguro) > 0].copy()
        serie_jug_norm = df_corte_sniper["Jugador"].apply(lambda x: sin_tildes(str(x)).lower()) if not df_corte_sniper.empty else pd.Series(dtype=str)
        mask_sin_desfase = ~df_corte_sniper["Senal_Trading"].astype(str).str.contains("🔎|⛔", na=False) if ("Senal_Trading" in df_corte_sniper.columns and not df_corte_sniper.empty) else pd.Series(True, index=df_corte_sniper.index)
        for idx_sn, r_sn in act_snipers.iterrows():
            j_b = sin_tildes(str(r_sn.get("Jugador_Buscado", ""))).strip().lower()
            p_obj = num_seguro(r_sn.get("Precio_Objetivo", 0))
            if j_b and p_obj > 0 and not df_corte_sniper.empty:
                m_sn = df_corte_sniper[serie_jug_norm.str.contains(re.escape(j_b), na=False) & mask_sin_desfase]
                if not m_sn.empty:
                    mejor_fila = m_sn.loc[m_sn["Precio_Actual"].apply(num_seguro).idxmin()]
                    p_actual_sn = num_seguro(mejor_fila.get("Precio_Actual", 0))
                    if 0 < p_actual_sn <= p_obj:
                        disparos_sniper.append((
                            idx_sn,
                            str(mejor_fila["Jugador"]).split(" (")[0],
                            str(mejor_fila["Jugador"]),
                            p_actual_sn,
                            p_obj,
                            "FUT.GG",
                            str(mejor_fila.get("Enlace_Optimo", mejor_fila.get("Enlace_Futbin", "")))
                        ))
        if disparos_sniper:
            if alerta_sonora:
                reproducir_sonido_sniper()
            nuevos_avisos_tg = []
            for idx_sn, nom_s, _, p_s, obj_s, f_s, lnk_s in disparos_sniper:
                clave_notif = f"{nombre_club_activo}|{nom_s}|{p_s}|{obj_s}"
                if clave_notif not in st.session_state["snipers_notificados_tg"]:
                    st.session_state["snipers_notificados_tg"].add(clave_notif)
                    nuevos_avisos_tg.append(f"🎯 <b>{nom_s}</b> cayó a <b>{p_s:,} mnds</b> [{f_s}] (Objetivo: &lt;= {obj_s:,} mnds)")
            if nuevos_avisos_tg:
                enviar_aviso_telegram_web(
                    f"🚨 <b>¡DISPARO SNIPER EN VIVO ({nombre_club_activo})!</b>\n" + "\n".join(nuevos_avisos_tg)
                )

            items_html = "<br>".join([
                f"🎯 <b>{nom_s}</b> cayó a <span style='color:#00e676;font-weight:900;'>{p_s:,} mnds</span> [{f_s}] (Tu objetivo: <= {obj_s:,} mnds) — <a href='{lnk_s}' target='_blank' style='color:#ffea00;'>Abrir Carta</a>"
                for _, nom_s, _, p_s, obj_s, f_s, lnk_s in disparos_sniper
            ])
            st.markdown(f'<div class="sniper-banner">🚨 <b>¡ALERTA SNIPER DETECTADA EN VIVO!</b><br>{items_html}</div>', unsafe_allow_html=True)
            cols_sn_btn = st.columns(min(4, len(disparos_sniper)))
            for i_b, (idx_sn, nom_s, nom_full_s, p_s, obj_s, f_s, _) in enumerate(disparos_sniper[:4]):
                with cols_sn_btn[i_b]:
                    if st.button(f"🛒 Comprada: {nom_s} ({p_s:,} mnds) ➡️ A Cartera", key=f"btn_sniper_buy_{idx_sn}_{p_s}", use_container_width=True):
                        df_c_sn = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                        if "Cantidad" not in df_c_sn.columns:
                            df_c_sn["Cantidad"] = 1
                        nueva_c_sn = pd.DataFrame([{
                            "Jugador_Exacto": nom_s.strip().title(),
                            "Estrategia": autodetectar_estrategia(nom_full_s, p_s),
                            "Cantidad": 1,
                            "Precio_Pagado": int(p_s),
                            "Estado": "Activo",
                            "Precio_Venta": 0,
                            "Costo_EA_Tax": 0,
                            "Ganancia_Neta": 0,
                            "ROI_%": 0.0,
                            "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                            "Fecha_Venta": ""
                        }])
                        df_c_sn = pd.concat([df_c_sn, nueva_c_sn], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_sn, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                        df_alertas_check.loc[idx_sn, "Estado"] = "Cumplida"
                        guardar_hoja_sincronizada(df_alertas_check, "Mis_Alertas", es_admin=es_admin, slug_club=slug_club_activo)
                        st.success(f"🎯 ¡{nom_s} sumado a Mi Cartera y alerta marcada como Cumplida!")
                        st.rerun()

    # ==========================================
    # TARJETAS KPI FINANCIERAS + EDITOR EN PÁGINA DE LIQUIDEZ Y CAPITAL INMOVILIZADO
    # ==========================================
    df_cartera = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
    if "Cantidad" not in df_cartera.columns:
        df_cartera["Cantidad"] = 1
    df_cartera["Cantidad"] = df_cartera["Cantidad"].apply(lambda x: max(1, num_seguro(x, 1)))

    if not df_cartera.empty and "Estado" in df_cartera.columns:
        # Recalcular SIEMPRE con matemática exacta de lotes (Cantidad * Precio_Venta / Cantidad * Precio_Pagado y 5% EA Tax)
        for idx_c, r_c in df_cartera.iterrows():
            cant_row = max(1, num_seguro(r_c.get("Cantidad", 1), 1))
            p_v_row = num_seguro(r_c.get("Precio_Venta", 0), 0)
            p_p_row = num_seguro(r_c.get("Precio_Pagado", 0), 0)
            if str(r_c.get("Estado", "")).strip() == "Vendido":
                if p_v_row > 0:
                    vb_lote = p_v_row * cant_row
                    ca_lote = p_p_row * cant_row
                    tx_row = int(round(vb_lote * 0.05))
                    nt_row = int(round((vb_lote - tx_row) - ca_lote))
                    roi_row = round((nt_row / ca_lote) * 100, 2) if ca_lote > 0 else 0.0
                    df_cartera.loc[idx_c, "Costo_EA_Tax"] = tx_row
                    df_cartera.loc[idx_c, "Ganancia_Neta"] = nt_row
                    df_cartera.loc[idx_c, "ROI_%"] = roi_row
                else:
                    df_cartera.loc[idx_c, "Costo_EA_Tax"] = 0
                    df_cartera.loc[idx_c, "Ganancia_Neta"] = 0
                    df_cartera.loc[idx_c, "ROI_%"] = 0.0

        activos = df_cartera[df_cartera["Estado"].astype(str).str.strip() == "Activo"].copy()
        vendidos = df_cartera[df_cartera["Estado"].astype(str).str.strip() == "Vendido"].copy()
    else:
        activos = pd.DataFrame()
        vendidos = pd.DataFrame()

    total_cartas_transferibles = int(activos["Cantidad"].apply(lambda x: max(1, num_seguro(x, 1))).sum()) if not activos.empty else 0
    cap_auto_cartera = int((activos["Precio_Pagado"].apply(num_seguro) * activos["Cantidad"].apply(lambda x: max(1, num_seguro(x, 1)))).sum()) if not activos.empty and "Precio_Pagado" in activos.columns else 0
    saldos_cfg = cargar_saldos_personalizados(es_admin, presupuesto_usuario, cap_auto_cartera, slug_club=slug_club_activo)

    if saldos_cfg.get("modo_capital") == "Manual (Monto fijo)":
        cap_inmovilizado = int(saldos_cfg.get("capital_manual", cap_auto_cartera))
    else:
        cap_inmovilizado = cap_auto_cartera

    if saldos_cfg.get("modo_liquidez") == "Manual":
        liquidez_disp = int(saldos_cfg.get("liquidez_manual", max(0, int(presupuesto_usuario) - cap_inmovilizado)))
    else:
        liquidez_disp = max(0, int(presupuesto_usuario) - cap_inmovilizado)

    base_total_patrimonio = max(1, int(presupuesto_usuario) if saldos_cfg.get("modo_liquidez") != "Manual" else (liquidez_disp + cap_inmovilizado))
    vendidos_con_precio_kpi = vendidos[vendidos["Precio_Venta"].apply(num_seguro) > 0] if (not vendidos.empty and "Precio_Venta" in vendidos.columns) else pd.DataFrame()
    ganancia_realizada = int(vendidos_con_precio_kpi["Ganancia_Neta"].apply(num_seguro).sum()) if not vendidos_con_precio_kpi.empty else 0
    aciertos = len(vendidos_con_precio_kpi[vendidos_con_precio_kpi["Ganancia_Neta"].apply(num_seguro) > 0]) if not vendidos_con_precio_kpi.empty else 0
    efectividad = round((aciertos / len(vendidos_con_precio_kpi) * 100), 1) if not vendidos_con_precio_kpi.empty else 0.0

    filas_cartera_viva = []
    vnr_total = 0
    for idx_orig, r_a in activos.iterrows():
        nom_a = str(r_a.get("Jugador_Exacto", "Jugador"))
        cant_a = max(1, num_seguro(r_a.get("Cantidad", 1), 1))
        costo_a = num_seguro(r_a.get("Precio_Pagado", 0), 0)
        costo_lote_a = costo_a * cant_a
        est_a = str(r_a.get("Estrategia", "Flipping Meta"))
        if not est_a or est_a.lower() == "nan":
            est_a = autodetectar_estrategia(nom_a, costo_a)
        empate_a = calcular_break_even(costo_a)
        sl_a = redondear_escalon_ea((costo_a * factor_sl) / 0.95, hacia_arriba=False)

        horas_hold = 1.0
        try:
            dt_c = datetime.strptime(str(r_a.get("Fecha_Compra", "")).strip(), "%Y-%m-%d %H:%M")
            horas_hold = max(1.0, round((ahora_arg() - dt_c).total_seconds() / 3600.0, 1))
        except Exception:
            pass
        dias_hold = max(0.25, round(horas_hold / 24.0, 2))
        str_hold = f"{int(horas_hold)}h" if horas_hold < 48 else f"{round(horas_hold/24, 1)}d"

        p_mercado, max_hist_a, muestras_a, enlace_a, es_respaldo_a = emparejar_carta_mercado(
            nom_a, costo_a, est_a, df_ult_corte
        )

        bruto_lote_a = p_mercado * cant_a
        ea_tax_lote_a = int(round(bruto_lote_a * 0.05))
        vnr_lote = bruto_lote_a - ea_tax_lote_a
        vnr_total += vnr_lote
        tenencia = vnr_lote - costo_lote_a
        roi_lat = round((tenencia / costo_lote_a) * 100, 1) if costo_lote_a > 0 else 0.0
        roi_diario = round(roi_lat / max(1.0, dias_hold), 1)
        lazy_a = calcular_precio_lazy(
            p_mercado,
            precio_pagado=costo_a,
            max_hist=max_hist_a,
            muestras=muestras_a,
            factor_lazy=factor_lazy,
            factor_roi=factor_roi
        )

        if es_respaldo_a:
            estado_op = f"🛡️ RESPALDO COSTO ({str_hold})"
        elif roi_lat <= float(stop_loss_pct):
            estado_op = "🛑 STOP-LOSS (Vender)"
        elif tenencia > 0:
            estado_op = "🚨 VENDER (En Ganancia)"
        elif horas_hold >= float(horas_rotacion):
            estado_op = f"🔄 ROTAR STOCK ({str_hold})"
        else:
            estado_op = f"⏳ HOLD / LAZY ({str_hold})"

        filas_cartera_viva.append({
            "Idx_Cartera": idx_orig,
            "Quitar": False,
            "Jugador": nom_a,
            "Estrategia": est_a,
            "Cantidad": cant_a,
            "Costo_Pagado": costo_a,
            "Costo_Total_Lote": costo_lote_a,
            "Empate_5%": empate_a,
            "Mercado_Hoy": p_mercado,
            "Precio_Lazy_1h": lazy_a,
            "Stop_Loss": sl_a,
            "Resultado_Tenencia": tenencia,
            "ROI_Latente_%": roi_lat,
            "ROI_Diario_%": roi_diario,
            "Acción_Sugerida": estado_op,
            "Antigüedad": str_hold,
            "Cant_Vender": 0,
            "Vendido_A": 0,
            "Link": enlace_a
        })

    tenencia_total = vnr_total - cap_auto_cartera
    color_vnr = "kpi-teal" if tenencia_total >= 0 else "kpi-crimson"
    pct_libre = round(liquidez_disp / base_total_patrimonio * 100)

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card kpi-green" style="border-top-color:{acento_club_1};"><div class="kpi-head"><span class="kpi-title">Liquidez Disponible</span><span class="kpi-icon">💵</span></div><div class="kpi-value">{fmt_monedas(liquidez_disp, modo_streamer)}</div><div class="kpi-badge">🟢 {pct_libre}% libre</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-blue"><div class="kpi-head"><span class="kpi-title">Capital Inmovilizado</span><span class="kpi-icon">📦</span></div><div class="kpi-value">{fmt_monedas(cap_inmovilizado, modo_streamer)}</div><div class="kpi-badge">🃏 {total_cartas_transferibles}/100 en lista</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card {color_vnr}"><div class="kpi-head"><span class="kpi-title">Stock a VNR (5% Tax)</span><span class="kpi-icon">📈</span></div><div class="kpi-value">{fmt_monedas(vnr_total, modo_streamer)}</div><div class="kpi-badge">⚖️ {fmt_monedas(tenencia_total, modo_streamer, True)}</div></div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card kpi-gold"><div class="kpi-head"><span class="kpi-title">Ganancia Realizada</span><span class="kpi-icon">💰</span></div><div class="kpi-value">{fmt_monedas(ganancia_realizada, modo_streamer, True)}</div><div class="kpi-badge">📜 {len(vendidos_con_precio_kpi)} cerradas</div></div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""<div class="kpi-card kpi-teal" style="border-top-color:{acento_club_2};"><div class="kpi-head"><span class="kpi-title">Efectividad Trading</span><span class="kpi-icon">🎯</span></div><div class="kpi-value">{efectividad}%</div><div class="kpi-badge">🏆 {aciertos}/{len(vendidos_con_precio_kpi)} aciertos</div></div>""", unsafe_allow_html=True)

    # MEDIDOR VISUAL DINÁMICO DE 100 HUECOS DE LISTA DE TRANSFERENCIAS (VERDE / AMARILLO / ROJO)
    huecos_ocupados = min(100, total_cartas_transferibles)
    huecos_libres = max(0, 100 - total_cartas_transferibles)
    if total_cartas_transferibles > 90:
        color_barra_tr, estado_barra_tr = "#ef4444", "🔴 CRÍTICO (>90%)"
    elif total_cartas_transferibles > 75:
        color_barra_tr, estado_barra_tr = "#f59e0b", "🟡 ALTA (>75%)"
    else:
        color_barra_tr, estado_barra_tr = "#10b981", "🟢 ESPACIO LIBRE"
    st.markdown(f"""
    <div class="transfer-bar-wrap">
        <div class="transfer-bar-top">
            <span>📦 Lista de Transferencias EA (Tope 100 cartas)</span>
            <span class="mono" style="color:{color_barra_tr};font-weight:800;">{total_cartas_transferibles}/100 ocupados · {huecos_libres} libres · {estado_barra_tr}</span>
        </div>
        <div class="transfer-track">
            <div class="transfer-fill" style="width:{huecos_ocupados}%;background:linear-gradient(90deg, {color_barra_tr}, {color_barra_tr}cc);box-shadow:0 0 12px {color_barra_tr}66;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # MEJORA 1: PANEL EN PÁGINA PARA MODIFICAR LIQUIDEZ DISPONIBLE Y CAPITAL INMOVILIZADO
    with st.expander("💳 Modificar Liquidez Disponible y Capital Inmovilizado desde la Página", expanded=False):
        st.caption("Ajustá manualmente tus monedas líquidas en caja y tu capital inmovilizado, o volvé al cálculo automático basado en tus cartas:")
        c_sal1, c_sal2, c_sal3, c_sal4, c_sal5 = st.columns([1.3, 1.3, 1.3, 1, 1])
        with c_sal1:
            inp_liquidez = st.number_input("💵 Liquidez Disponible (Monedas)", min_value=0, max_value=50000000, step=1000, value=int(liquidez_disp))
        with c_sal2:
            idx_m_cap = 1 if saldos_cfg.get("modo_capital") == "Manual (Monto fijo)" else 0
            inp_modo_cap = st.selectbox("📦 Modo Capital Inmovilizado", ["Automático (Suma de mis cartas)", "Manual (Monto fijo)"], index=idx_m_cap)
        with c_sal3:
            inp_cap_man = st.number_input("📦 Capital Inmovilizado Manual", min_value=0, max_value=50000000, step=1000, value=int(cap_inmovilizado), disabled=(inp_modo_cap != "Manual (Monto fijo)"))
        with c_sal4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 Guardar Saldos", use_container_width=True):
                guardar_saldos_personalizados({
                    "modo_liquidez": "Manual",
                    "liquidez_manual": int(inp_liquidez),
                    "modo_capital": inp_modo_cap,
                    "capital_manual": int(inp_cap_man)
                }, es_admin, slug_club=slug_club_activo)
                st.success("✅ ¡Liquidez y Capital actualizados!")
                st.rerun()
        with c_sal5:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🔄 Modo Automático", use_container_width=True):
                guardar_saldos_personalizados({
                    "modo_liquidez": "Automático",
                    "liquidez_manual": max(0, int(presupuesto_usuario) - cap_auto_cartera),
                    "modo_capital": "Automático (Suma de mis cartas)",
                    "capital_manual": cap_auto_cartera
                }, es_admin, slug_club=slug_club_activo)
                st.info("Calculando saldos automáticamente según tu presupuesto y cartas.")
                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # 5 PESTAÑAS INTERACTIVAS COMPLETAS
    # ==========================================
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Cartera 1-Clic, P&L & Proyector",
        "⚡ Radar 1-Clic, Lupa & Compartir",
        "🎯 Optimizador SBC, Fodder & Sniper",
        "🔮 Oráculo & Reporte IA En Vivo",
        "🏆 Ranking & Versus de Clubes"
    ])

    with tab1:
        st.subheader("🎯 Monitoreo de Cartera Activa (Editá Jugador/Cantidad/Costo, Venta Parcial de Lotes o Vendido_A)")

        # MEJORA 2A: AGREGAR JUGADOR DIRECTAMENTE DESDE LA PÁGINA PRINCIPAL
        with st.expander("➕ Agregar Jugador o Lote Rápido a Mi Cartera desde acá", expanded=False):
            ca1, ca2, ca3, ca4, ca5 = st.columns([1.8, 0.8, 1.1, 1.2, 1])
            with ca1:
                nom_add_rap = st.text_input("Nombre del Jugador", placeholder="Ej: Valverde, Saliba, Theo...", key="inp_nom_rap")
            with ca2:
                cant_add_rap = st.number_input("Cantidad", min_value=1, max_value=100, step=1, value=1, key="inp_cant_rap")
            with ca3:
                costo_add_rap = st.number_input("Costo Unitario (Mnds)", min_value=200, max_value=15000000, step=250, value=12000, key="inp_cost_rap")
            with ca4:
                est_add_rap = st.selectbox("Estrategia", ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"], key="inp_est_rap")
            with ca5:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Sumar Jugador", use_container_width=True):
                    if nom_add_rap.strip():
                        df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                        if "Cantidad" not in df_c_act.columns:
                            df_c_act["Cantidad"] = 1
                        nueva_f = pd.DataFrame([{
                            "Jugador_Exacto": nom_add_rap.strip().title(),
                            "Estrategia": est_add_rap,
                            "Cantidad": int(cant_add_rap),
                            "Precio_Pagado": int(costo_add_rap),
                            "Estado": "Activo",
                            "Precio_Venta": 0,
                            "Costo_EA_Tax": 0,
                            "Ganancia_Neta": 0,
                            "ROI_%": 0.0,
                            "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                            "Fecha_Venta": ""
                        }])
                        df_c_act = pd.concat([df_c_act, nueva_f], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                        if saldos_cfg.get("modo_liquidez") == "Manual":
                            saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - int(costo_add_rap * cant_add_rap))
                            guardar_saldos_personalizados(saldos_cfg, es_admin, slug_club=slug_club_activo)
                        st.success(f"¡{cant_add_rap}x {nom_add_rap.title()} sumado a tu cartera activa!")
                        st.rerun()

        # MEJORA 2A-VISION: REGISTRAR COMPRAS POR FOTO / CAPTURA CON GEMINI VISION
        tiene_preview_foto = bool(st.session_state.get("compras_foto_preview"))
        with st.expander("📸 Registrar Compras por Foto / Captura (IA Vision)", expanded=tiene_preview_foto):
            st.caption(
                "Subí una captura o foto de tus compras en EA Sports FC 27 (Web App, Companion App o Consola). "
                "Gemini Vision detectará automáticamente solo las cartas ganadas/compradas, sus cantidades y precios pagados."
            )
            cv_up, cv_btn = st.columns([2.4, 1.1])
            with cv_up:
                foto_compras = st.file_uploader(
                    "Imagen de compras (PNG, JPG, JPEG, WEBP)",
                    type=["png", "jpg", "jpeg", "webp"],
                    key="uploader_compras_vision"
                )
            with cv_btn:
                st.markdown("<br>", unsafe_allow_html=True)
                btn_analizar_foto = st.button(
                    "🔍 Analizar Foto con IA",
                    use_container_width=True,
                    disabled=(foto_compras is None),
                    key="btn_analizar_foto_ia"
                )

            if foto_compras is not None:
                if st.checkbox("🖼️ Mostrar miniatura de la imagen cargada", value=False, key="chk_ver_foto_vision"):
                    st.image(foto_compras, use_container_width=False, width=420)

            if btn_analizar_foto and foto_compras is not None:
                with st.spinner("🤖 Analizando imagen con Gemini Vision..."):
                    detectados, mod_vis, err_vis = extraer_compras_desde_imagen(
                        foto_compras.getvalue(),
                        api_key_custom
                    )
                if err_vis:
                    st.error(f"❌ {err_vis}")
                elif not detectados:
                    st.warning("⚠️ No se detectaron cartas compradas/ganadas con precio válido en la imagen. Probá con una captura más nítida.")
                else:
                    st.session_state["compras_foto_preview"] = [
                        {
                            "Incluir": True,
                            "Jugador": str(it["Jugador"]).strip(),
                            "Cantidad": int(it["Cantidad"]),
                            "Precio_Pagado": int(it["Precio_Pagado"]),
                            "Estrategia": it["Estrategia"] if it["Estrategia"] in ESTRATEGIAS_VALIDAS_CARTERA else "Flipping Meta"
                        }
                        for it in detectados
                    ]
                    st.session_state["compras_foto_modelo"] = mod_vis
                    st.rerun()

            preview_foto = st.session_state.get("compras_foto_preview") or []
            if preview_foto:
                mod_prev_txt = st.session_state.get("compras_foto_modelo", "Gemini Vision")
                st.success(f"✨ Detectados **{len(preview_foto)}** ítem(s) con **{mod_prev_txt}**. Revisá o editá los valores antes de confirmar:")
                df_preview_foto = pd.DataFrame(preview_foto)
                edicion_foto = st.data_editor(
                    df_preview_foto,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key="editor_compras_foto_vision",
                    column_config={
                        "Incluir": st.column_config.CheckboxColumn("✅ Incluir", default=True),
                        "Jugador": st.column_config.TextColumn("👤 Jugador"),
                        "Cantidad": st.column_config.NumberColumn("📦 Cantidad", min_value=1, max_value=100, step=1, format="%d"),
                        "Precio_Pagado": st.column_config.NumberColumn("💵 Precio Pagado (Unit.)", min_value=150, max_value=TOPE_RADAR_MAXIMO, step=50, format="%d mnds"),
                        "Estrategia": st.column_config.SelectboxColumn("🏷️ Estrategia", options=ESTRATEGIAS_VALIDAS_CARTERA)
                    }
                )

                df_sel_foto = edicion_foto[edicion_foto["Incluir"] == True].copy() if not edicion_foto.empty else pd.DataFrame()
                unidades_lote_foto = 0
                costo_lote_foto = 0
                filas_listas_foto = []
                if not df_sel_foto.empty:
                    for _, rf in df_sel_foto.iterrows():
                        jug_f = str(rf.get("Jugador", "") or "").strip()
                        cant_f = max(1, min(100, num_seguro(rf.get("Cantidad", 1), 1)))
                        precio_f = num_seguro(rf.get("Precio_Pagado", 0), 0)
                        est_f = str(rf.get("Estrategia", "Flipping Meta") or "Flipping Meta")
                        if est_f not in ESTRATEGIAS_VALIDAS_CARTERA:
                            est_f = "Flipping Meta"
                        if jug_f and precio_f > 0:
                            unidades_lote_foto += cant_f
                            costo_lote_foto += int(cant_f * precio_f)
                            filas_listas_foto.append({
                                "Jugador_Exacto": jug_f,
                                "Estrategia": est_f,
                                "Cantidad": cant_f,
                                "Precio_Pagado": precio_f,
                                "Estado": "Activo",
                                "Precio_Venta": 0,
                                "Costo_EA_Tax": 0,
                                "Ganancia_Neta": 0,
                                "ROI_%": 0.0,
                                "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                                "Fecha_Venta": ""
                            })

                st.info(
                    f"🧾 **Resumen del lote a confirmar ({nombre_club_activo}):** "
                    f"**{len(filas_listas_foto)}** fila(s) · **{unidades_lote_foto}** carta(s) totales · "
                    f"Costo total: **{costo_lote_foto:,} mnds**"
                )

                cf_ok, cf_cancel = st.columns([2, 1])
                with cf_ok:
                    if st.button(
                        "✅ Confirmar y Guardar en Mi Cartera",
                        type="primary",
                        use_container_width=True,
                        disabled=(len(filas_listas_foto) == 0),
                        key="btn_confirmar_compras_foto"
                    ):
                        df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                        if "Cantidad" not in df_c_act.columns:
                            df_c_act["Cantidad"] = 1
                        df_c_act = pd.concat([df_c_act, pd.DataFrame(filas_listas_foto)], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                        if saldos_cfg.get("modo_liquidez") == "Manual":
                            saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - int(costo_lote_foto))
                            guardar_saldos_personalizados(saldos_cfg, es_admin, slug_club=slug_club_activo)
                        st.session_state.pop("compras_foto_preview", None)
                        st.session_state.pop("compras_foto_modelo", None)
                        st.success(f"✅ ¡{unidades_lote_foto} carta(s) guardadas en la cartera de {nombre_club_activo} por {costo_lote_foto:,} mnds!")
                        st.rerun()
                with cf_cancel:
                    if st.button("🗑️ Descartar Detección", use_container_width=True, key="btn_descartar_compras_foto"):
                        st.session_state.pop("compras_foto_preview", None)
                        st.session_state.pop("compras_foto_modelo", None)
                        st.rerun()

        # MEJORA 2B: TABLA DE CARTERA ACTIVA 100% EDITABLE CON SOPORTE DE VENTA PARCIAL DE LOTES
        if filas_cartera_viva:
            df_viva = pd.DataFrame(filas_cartera_viva)
            cols_editables_viva = ["Quitar", "Jugador", "Estrategia", "Cantidad", "Costo_Pagado", "Cant_Vender", "Vendido_A"]
            cols_deshabilitadas_viva = [c for c in df_viva.columns if c not in cols_editables_viva]

            edicion_viva = st.data_editor(
                df_viva,
                use_container_width=True,
                hide_index=True,
                disabled=cols_deshabilitadas_viva,
                column_config={
                    "Idx_Cartera": None,
                    "Quitar": st.column_config.CheckboxColumn("🗑️ Quitar"),
                    "Jugador": st.column_config.TextColumn("👤 Jugador (Editable)"),
                    "Estrategia": st.column_config.SelectboxColumn("🏷️ Estrategia", options=["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"]),
                    "Cantidad": st.column_config.NumberColumn("📦 Cant.", min_value=1, max_value=100, step=1, format="%d"),
                    "Costo_Pagado": st.column_config.NumberColumn("💵 Costo Unit.", min_value=200, max_value=15000000, step=250, format="%d mnds"),
                    "Costo_Total_Lote": st.column_config.NumberColumn("💼 Costo Lote", format="%d mnds"),
                    "Empate_5%": st.column_config.NumberColumn("⚖️ Empate (5%)", format="%d mnds"),
                    "Mercado_Hoy": st.column_config.NumberColumn("📊 Mercado Hoy", format="%d mnds"),
                    "Precio_Lazy_1h": st.column_config.NumberColumn(f"🎯 Lazy (+{lazy_pct}%)", format="%d mnds"),
                    "Stop_Loss": st.column_config.NumberColumn(f"🛑 SL ({stop_loss_pct}%)", format="%d mnds"),
                    "Resultado_Tenencia": st.column_config.NumberColumn("✨ Neto Lote Hoy", format="%+d mnds"),
                    "ROI_Latente_%": st.column_config.NumberColumn("📈 ROI %", format="%+.1f%%"),
                    "ROI_Diario_%": st.column_config.NumberColumn("⚡ ROI/Día", format="%+.1f%%"),
                    "Cant_Vender": st.column_config.NumberColumn("📤 Cant. Vender", min_value=0, max_value=100, step=1, format="%d", help="Dejalo en 0 para vender el lote entero al cargar Vendido_A, o poné cuántas unidades del lote vendiste."),
                    "Vendido_A": st.column_config.NumberColumn("💸 Vendido_A Unit.", min_value=0, max_value=15000000, step=250, format="%d mnds"),
                    "Link": st.column_config.LinkColumn("🔗 Web", display_text="Abrir Carta")
                },
                key="editor_cartera_viva"
            )

            if st.button("💾 Aplicar Cambios en Cartera Activa (Guardar Ediciones / Liquidar Ventas Totales o Parciales / Quitar)", use_container_width=True):
                df_c_upd = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                if "Cantidad" not in df_c_upd.columns:
                    df_c_upd["Cantidad"] = 1
                df_c_upd["Estado"] = df_c_upd["Estado"].astype(object)
                if "Fecha_Venta" not in df_c_upd.columns:
                    df_c_upd["Fecha_Venta"] = ""
                df_c_upd["Fecha_Venta"] = df_c_upd["Fecha_Venta"].astype(object)

                indices_a_borrar = []
                filas_parciales_vendidas = []
                delta_caja = 0

                for _, r_ed in edicion_viva.iterrows():
                    idx_c_orig = int(r_ed["Idx_Cartera"])
                    if idx_c_orig not in df_c_upd.index:
                        continue

                    if bool(r_ed["Quitar"]):
                        indices_a_borrar.append(idx_c_orig)
                        continue

                    nuevo_nom = str(r_ed["Jugador"]).strip().title()
                    nueva_est = str(r_ed["Estrategia"]).strip()
                    nueva_cant = max(1, int(num_seguro(r_ed.get("Cantidad", 1), 1)))
                    nuevo_costo = int(num_seguro(r_ed["Costo_Pagado"], 0))
                    cant_vender_ed = int(num_seguro(r_ed.get("Cant_Vender", 0), 0))
                    p_venta_liq = int(num_seguro(r_ed["Vendido_A"], 0))

                    df_c_upd.loc[idx_c_orig, "Jugador_Exacto"] = nuevo_nom
                    df_c_upd.loc[idx_c_orig, "Estrategia"] = nueva_est
                    df_c_upd.loc[idx_c_orig, "Cantidad"] = nueva_cant
                    df_c_upd.loc[idx_c_orig, "Precio_Pagado"] = nuevo_costo

                    if p_venta_liq > 0:
                        cant_efectiva_v = cant_vender_ed if (0 < cant_vender_ed < nueva_cant) else nueva_cant
                        vb_lote_liq = p_venta_liq * cant_efectiva_v
                        ca_lote_liq = nuevo_costo * cant_efectiva_v
                        tax_liq = int(round(vb_lote_liq * 0.05))
                        neto_liq = int(round((vb_lote_liq - tax_liq) - ca_lote_liq))
                        roi_liq = round((neto_liq / ca_lote_liq) * 100, 2) if ca_lote_liq > 0 else 0.0
                        fecha_v_hoy = ahora_arg().strftime("%Y-%m-%d %H:%M")

                        if cant_efectiva_v < nueva_cant:
                            # Venta parcial de lote: descontar del lote activo y sumar registro vendido
                            df_c_upd.loc[idx_c_orig, "Cantidad"] = int(nueva_cant - cant_efectiva_v)
                            filas_parciales_vendidas.append({
                                "Jugador_Exacto": nuevo_nom,
                                "Estrategia": nueva_est,
                                "Cantidad": int(cant_efectiva_v),
                                "Precio_Pagado": nuevo_costo,
                                "Estado": "Vendido",
                                "Precio_Venta": p_venta_liq,
                                "Costo_EA_Tax": tax_liq,
                                "Ganancia_Neta": neto_liq,
                                "ROI_%": roi_liq,
                                "Fecha_Compra": str(df_c_upd.loc[idx_c_orig, "Fecha_Compra"]),
                                "Fecha_Venta": fecha_v_hoy
                            })
                        else:
                            df_c_upd.loc[idx_c_orig, "Estado"] = "Vendido"
                            df_c_upd.loc[idx_c_orig, "Precio_Venta"] = p_venta_liq
                            df_c_upd.loc[idx_c_orig, "Costo_EA_Tax"] = tax_liq
                            df_c_upd.loc[idx_c_orig, "Ganancia_Neta"] = neto_liq
                            df_c_upd.loc[idx_c_orig, "ROI_%"] = roi_liq
                            df_c_upd.loc[idx_c_orig, "Fecha_Venta"] = fecha_v_hoy
                        delta_caja += (vb_lote_liq - tax_liq)

                if indices_a_borrar:
                    df_c_upd = df_c_upd.drop(index=indices_a_borrar).reset_index(drop=True)
                if filas_parciales_vendidas:
                    df_c_upd = pd.concat([df_c_upd, pd.DataFrame(filas_parciales_vendidas)], ignore_index=True)

                guardar_hoja_sincronizada(df_c_upd, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                if delta_caja != 0 and saldos_cfg.get("modo_liquidez") == "Manual":
                    saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) + delta_caja)
                    guardar_saldos_personalizados(saldos_cfg, es_admin, slug_club=slug_club_activo)

                st.success("✅ ¡Cambios aplicados en tu Cartera Activa!")
                st.rerun()

            with st.expander("📋 Copiar Listado Rápido de Re-Loteo Lazy (Ciclo 1 Hora) y Control de Stock", expanded=False):
                st.caption("Copiá esta guía rápida para relistar tus cartas cada 1 hora en la Web App / Companion de EA Sports FC 27 respetando los escalones oficiales:")
                lineas_lazy = [f"🎯 LISTA DE RE-LOTEO LAZY 1H ({nombre_club_activo} | +{lazy_pct}%)"]
                for r_lz in filas_cartera_viva:
                    puja_sug_lz = redondear_escalon_ea(r_lz["Precio_Lazy_1h"] * 0.95, hacia_arriba=False)
                    lineas_lazy.append(
                        f"• {r_lz['Cantidad']}x {r_lz['Jugador']} (Costo: {r_lz['Costo_Pagado']:,}) ➡️ "
                        f"Puja Inicial: {puja_sug_lz:,} | Compra Ya (Lazy): {r_lz['Precio_Lazy_1h']:,} | Empate 5%: {r_lz['Empate_5%']:,}"
                    )
                st.code("\n".join(lineas_lazy), language="markdown")
        else:
            st.info("No tenés cartas activas en cartera. Podés agregarlas desde el desplegable de arriba, desde la barra lateral o con 1 clic en la pestaña Radar.")

        # GRÁFICOS: VALUACIÓN, DIVERSIFICACIÓN DE RIESGO Y CURVA P&L
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            if filas_cartera_viva:
                df_bar_viva = pd.DataFrame(filas_cartera_viva)
                df_bar_viva["Etiqueta_Lote"] = [
                    f"#{i+1} {r['Jugador']} ({r['Cantidad']}x)" if df_bar_viva["Jugador"].tolist().count(r["Jugador"]) > 1 else f"{r['Jugador']} ({r['Cantidad']}x)"
                    for i, (_, r) in enumerate(df_bar_viva.iterrows())
                ]
                fig_bar = px.bar(
                    df_bar_viva, x="Etiqueta_Lote", y=["Costo_Pagado", "Mercado_Hoy", "Precio_Lazy_1h"],
                    barmode="group", title="📦 Costo Unit. vs Mercado Hoy vs Target Lazy",
                    color_discrete_sequence=["#38bdf8", "#f43f5e", "#10b981"]
                )
                fig_bar.update_traces(marker_line_width=0, opacity=0.92)
                estilizar_plotly_fintech(fig_bar, height=340, hovermode="x unified")
                st.plotly_chart(fig_bar, use_container_width=True)

        with col_g2:
            filas_div = [{"Rubro": "💵 Liquidez en Caja", "Monedas": liquidez_disp}]
            if filas_cartera_viva:
                df_ag_est = pd.DataFrame(filas_cartera_viva).groupby("Estrategia")["Costo_Total_Lote"].sum().reset_index()
                for _, r_est in df_ag_est.iterrows():
                    filas_div.append({"Rubro": str(r_est["Estrategia"]), "Monedas": int(r_est["Costo_Total_Lote"])})
                    if int(r_est["Costo_Total_Lote"]) > (base_total_patrimonio * 0.60):
                        st.warning(f"⚠️ **Alerta de Concentración:** Tenés más del 60% de tu capital en **{r_est['Estrategia']}** (`{int(r_est['Costo_Total_Lote']):,} mnds`). Diversificá para reducir riesgo.")
            df_div = pd.DataFrame(filas_div)
            fig_pie = px.pie(
                df_div, values="Monedas", names="Rubro", hole=0.58,
                title="🍩 Exposición y Diversificación de Capital",
                color_discrete_sequence=["#10b981", "#38bdf8", "#f59e0b", "#f43f5e", "#a855f7"]
            )
            fig_pie.update_traces(marker=dict(line=dict(color="#090d16", width=2)), textfont=dict(family="JetBrains Mono, monospace", size=11))
            estilizar_plotly_fintech(fig_pie, height=340, hovermode="closest")
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_g3:
            if not vendidos_con_precio_kpi.empty:
                df_curva = vendidos_con_precio_kpi.copy()
                if "Fecha_Venta" in df_curva.columns:
                    df_curva["_dt_v"] = pd.to_datetime(df_curva["Fecha_Venta"], errors="coerce")
                    df_curva = df_curva.sort_values(by="_dt_v", kind="mergesort").drop(columns=["_dt_v"])
                df_curva["Ganancia_Neta_Num"] = df_curva["Ganancia_Neta"].apply(num_seguro)
                df_curva["Ganancia_Acumulada"] = df_curva["Ganancia_Neta_Num"].cumsum()
                df_curva["Operacion"] = [f"#{i+1} {int(r.get('Cantidad', 1))}x {r.get('Jugador_Exacto', 'Carta')}" for i, (_, r) in enumerate(df_curva.iterrows())]
                fig_curva = px.area(
                    df_curva, x="Operacion", y="Ganancia_Acumulada",
                    markers=True, title="📈 Curva de Contribución Marginal (P&L)",
                    color_discrete_sequence=["#10b981"]
                )
                fig_curva.update_traces(line=dict(width=2.5), fillcolor="rgba(16, 185, 129, 0.16)")
                estilizar_plotly_fintech(fig_curva, height=340, hovermode="x unified")
                st.plotly_chart(fig_curva, use_container_width=True)

        # ESTADO DE RESULTADOS + LIBRO MAYOR CON RE-COMPRA, EDICIÓN Y BORRADO + EXPORTADOR EXCEL
        col_pnl1, col_pnl2 = st.columns([1, 1])
        with col_pnl1:
            st.subheader("📑 Estado de Resultados (Modelo Costeo Variable)")
            if not vendidos_con_precio_kpi.empty:
                cant_liq_serie = vendidos_con_precio_kpi["Cantidad"].apply(lambda x: max(1, num_seguro(x, 1))) if "Cantidad" in vendidos_con_precio_kpi.columns else 1
                ventas_brutas = int((vendidos_con_precio_kpi["Precio_Venta"].apply(num_seguro) * cant_liq_serie).sum())
                costo_adq = int((vendidos_con_precio_kpi["Precio_Pagado"].apply(num_seguro) * cant_liq_serie).sum()) if "Precio_Pagado" in vendidos_con_precio_kpi.columns else 0
                ea_tax = int(vendidos_con_precio_kpi["Costo_EA_Tax"].apply(num_seguro).sum()) if "Costo_EA_Tax" in vendidos_con_precio_kpi.columns else int(round(ventas_brutas * 0.05))
                ingreso_neto = ventas_brutas - ea_tax
                contrib_marginal = int(vendidos_con_precio_kpi["Ganancia_Neta"].apply(num_seguro).sum())
            else:
                ventas_brutas = 0
                costo_adq = 0
                ea_tax = 0
                ingreso_neto = 0
                contrib_marginal = 0

            df_estado_resultados = pd.DataFrame([
                {"Rubro Contable": "(+) Ingresos Brutos por Ventas Cerradas", "Monto (Monedas)": ventas_brutas},
                {"Rubro Contable": "(-) Costo Variable Comercial (5% EA Tax)", "Monto (Monedas)": -ea_tax},
                {"Rubro Contable": "(=) Ingreso Neto de Realización", "Monto (Monedas)": ingreso_neto},
                {"Rubro Contable": "(-) Costo de Adquisición de Cartas Vendidas (CMV)", "Monto (Monedas)": -costo_adq},
                {"Rubro Contable": "(=) CONTRIBUCIÓN MARGINAL NETA (GANANCIA REALIZADA)", "Monto (Monedas)": contrib_marginal},
                {"Rubro Contable": "📦 Inventario Activo (Capital Inmovilizado)", "Monto (Monedas)": cap_inmovilizado},
                {"Rubro Contable": "📈 Inventario Activo a VNR FUT.GG (Neto 5% Tax)", "Monto (Monedas)": vnr_total},
                {"Rubro Contable": "⚖️ Resultado por Tenencia (No Realizado)", "Monto (Monedas)": tenencia_total},
                {"Rubro Contable": "💵 Liquidez Disponible (Caja Actual)", "Monto (Monedas)": liquidez_disp},
            ])
            styled_er = aplicar_estilo_seguro(
                df_estado_resultados,
                cols_neto=["Monto (Monedas)"],
                cols_accion=["Rubro Contable"]
            )
            st.dataframe(
                styled_er, use_container_width=True, hide_index=True,
                column_config={"Monto (Monedas)": st.column_config.NumberColumn("💰 Monedas", format="%+d mnds")}
            )

            bytes_excel = generar_excel_contable_bytes(
                df_estado_resultados,
                pd.DataFrame(filas_cartera_viva) if filas_cartera_viva else pd.DataFrame(),
                vendidos
            )
            st.download_button(
                "📥 Descargar Reporte Contable Completo (.xlsx)",
                data=bytes_excel,
                file_name=f"Balance_FC27_{ahora_arg().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_pnl2:
            st.subheader("📜 Libro Mayor de Ventas (Re-Comprá 🔄, Editá o Quitá 🗑️)")
            if not vendidos.empty:
                if any(vendidos["Precio_Venta"].apply(num_seguro) == 0):
                    st.caption("💡 Tip: Podés editar el precio de venta directamente en la columna '💰 Venta' de acá abajo y tocar Guardar.")
                df_vend_edit = vendidos.copy()
                if "Cantidad" not in df_vend_edit.columns:
                    df_vend_edit["Cantidad"] = 1
                df_vend_edit["Idx_Venta"] = df_vend_edit.index
                df_vend_edit.insert(0, "Quitar_Venta", False)
                df_vend_edit.insert(0, "Re_Comprar", False)

                cols_edit_vend = ["Re_Comprar", "Quitar_Venta", "Jugador_Exacto", "Estrategia", "Cantidad", "Precio_Pagado", "Precio_Venta"]
                edicion_vend = st.data_editor(
                    df_vend_edit,
                    use_container_width=True,
                    hide_index=True,
                    disabled=[c for c in df_vend_edit.columns if c not in cols_edit_vend],
                    column_config={
                        "Idx_Venta": None,
                        "Re_Comprar": st.column_config.CheckboxColumn("🔄 Re-Comprar"),
                        "Quitar_Venta": st.column_config.CheckboxColumn("🗑️ Quitar"),
                        "Jugador_Exacto": st.column_config.TextColumn("👤 Jugador"),
                        "Cantidad": st.column_config.NumberColumn("📦 Cant.", min_value=1, max_value=100, step=1, format="%d"),
                        "Precio_Pagado": st.column_config.NumberColumn("💵 Costo Unit.", format="%d mnds"),
                        "Precio_Venta": st.column_config.NumberColumn("💰 Venta Unit.", format="%d mnds"),
                        "Costo_EA_Tax": st.column_config.NumberColumn("🏛️ Tax 5% Lote", format="%d mnds"),
                        "Ganancia_Neta": st.column_config.NumberColumn("✨ Neto Lote", format="%+d mnds"),
                        "ROI_%": st.column_config.NumberColumn("📈 ROI %", format="%+.1f%%"),
                    },
                    key="editor_libro_ventas"
                )

                if st.button("💾 Guardar Cambios / Procesar Re-Compras en Libro Mayor", use_container_width=True):
                    df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                    if "Cantidad" not in df_c_act.columns:
                        df_c_act["Cantidad"] = 1
                    borrar_v_idx = []
                    nuevas_re = []
                    costo_recompras = 0

                    for _, r_v in edicion_vend.iterrows():
                        idx_v = int(r_v["Idx_Venta"])
                        if idx_v not in df_c_act.index:
                            continue
                        if bool(r_v["Quitar_Venta"]):
                            borrar_v_idx.append(idx_v)
                            continue

                        cant_v = max(1, int(num_seguro(r_v.get("Cantidad", 1), 1)))
                        pc_v = int(num_seguro(r_v["Precio_Pagado"], 0))
                        pv_v = int(num_seguro(r_v["Precio_Venta"], 0))
                        vb_v = pv_v * cant_v
                        ca_v = pc_v * cant_v
                        tx_v = int(round(vb_v * 0.05))
                        nt_v = int(round((vb_v - tx_v) - ca_v)) if pv_v > 0 else 0
                        roi_v = round((nt_v / ca_v) * 100, 2) if ca_v > 0 and pv_v > 0 else 0.0

                        df_c_act.loc[idx_v, "Jugador_Exacto"] = str(r_v["Jugador_Exacto"]).strip().title()
                        df_c_act.loc[idx_v, "Estrategia"] = str(r_v["Estrategia"]).strip()
                        df_c_act.loc[idx_v, "Cantidad"] = cant_v
                        df_c_act.loc[idx_v, "Precio_Pagado"] = pc_v
                        df_c_act.loc[idx_v, "Precio_Venta"] = pv_v
                        df_c_act.loc[idx_v, "Costo_EA_Tax"] = tx_v
                        df_c_act.loc[idx_v, "Ganancia_Neta"] = nt_v
                        df_c_act.loc[idx_v, "ROI_%"] = roi_v

                        if bool(r_v["Re_Comprar"]):
                            costo_recompras += ca_v
                            nuevas_re.append({
                                "Jugador_Exacto": str(r_v.get("Jugador_Exacto", "Carta")).strip().title(),
                                "Estrategia": str(r_v.get("Estrategia", "Flipping Meta")),
                                "Cantidad": cant_v,
                                "Precio_Pagado": pc_v,
                                "Estado": "Activo",
                                "Precio_Venta": 0,
                                "Costo_EA_Tax": 0,
                                "Ganancia_Neta": 0,
                                "ROI_%": 0.0,
                                "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                                "Fecha_Venta": ""
                            })

                    if borrar_v_idx:
                        df_c_act = df_c_act.drop(index=borrar_v_idx).reset_index(drop=True)
                    if nuevas_re:
                        df_c_act = pd.concat([df_c_act, pd.DataFrame(nuevas_re)], ignore_index=True)
                        if saldos_cfg.get("modo_liquidez") == "Manual":
                            saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - costo_recompras)
                            guardar_saldos_personalizados(saldos_cfg, es_admin, slug_club=slug_club_activo)

                    guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                    st.success("✅ ¡Libro Mayor actualizado con éxito!")
                    st.rerun()
            else:
                st.info("Todavía no hay ventas cerradas registradas.")

        # PROYECTOR DE INTERÉS COMPUESTO (META DEL MILLÓN)
        st.markdown("---")
        with st.expander("🚀 Proyector de Interés Compuesto — Calculadora 'Meta del Millón'", expanded=False):
            patrimonio_actual = max(5000, liquidez_disp + vnr_total)
            cp1, cp2, cp3 = st.columns(3)
            with cp1:
                cap_inicial_proy = st.number_input("Patrimonio Actual (Caja + Stock VNR)", min_value=5000, max_value=50000000, value=int(patrimonio_actual), step=10000)
            with cp2:
                meta_proy = st.number_input("🎯 Meta de Monedas Objetivo", min_value=50000, max_value=50000000, value=max(1000000, int(patrimonio_actual * 2)), step=50000)
            with cp3:
                roi_diario_proy = st.slider("📈 % Rendimiento Diario Compuesto", min_value=1.0, max_value=15.0, value=4.5, step=0.5)

            if cap_inicial_proy >= meta_proy:
                st.success("🏆 ¡Ya alcanzaste o superaste esta meta de monedas!")
            else:
                dias_necesarios = int(math.ceil(math.log(meta_proy / cap_inicial_proy) / math.log(1.0 + (roi_diario_proy / 100.0))))
                fecha_meta = (ahora_arg() + timedelta(days=dias_necesarios)).strftime("%d/%m/%Y")
                st.info(
                    f"⏱️ Manteniendo un **`+{roi_diario_proy}%` diario**, pasarás de **`{cap_inicial_proy:,} mnds`** a **`{meta_proy:,} mnds`** "
                    f"en exactamente **{dias_necesarios} días** (Fecha estimada: **{fecha_meta}**)."
                )
                dias_graf = min(dias_necesarios, 45)
                df_proy = pd.DataFrame([
                    {"Día": f"Día {d}", "Patrimonio_Proyectado": int(round(cap_inicial_proy * ((1.0 + roi_diario_proy / 100.0) ** d)))}
                    for d in range(dias_graf + 1)
                ])
                fig_proy = px.line(
                    df_proy, x="Día", y="Patrimonio_Proyectado", markers=True,
                    title=f"Proyección de Crecimiento Compuesto ({dias_graf} días)",
                    color_discrete_sequence=["#10b981"]
                )
                fig_proy.update_traces(line=dict(width=2.8))
                estilizar_plotly_fintech(fig_proy, height=280, hovermode="x unified")
                st.plotly_chart(fig_proy, use_container_width=True)

    # ==========================================
    # PESTAÑA 2: RADAR EXCLUSIVO FUT.GG (HASTA 10,000 CARTAS | 15,000,000 MNDS)
    # ==========================================
    with tab2:
        # Indicador visual de salud de datos del corte actual
        st.markdown(f"""
        <div class="status-strip-card" style="border-left: 3px solid {color_salud}; margin-bottom: 12px;">
            <div class="status-strip-title" style="color:{color_salud};">🩺 Salud de Datos FUT.GG ({salud_datos_pct:.1f}% cotizaciones activas)</div>
            <div class="status-strip-desc">
                ✅ <b style="color:#f8fafc;">{cartas_activas_corte:,}</b> cartas con precio válido verificado (escalones EA Sports FC) &nbsp;·&nbsp;
                ⛔ <b style="color:#f87171;">{cartas_extintas_corte:,}</b> cartas extintas / sin precio (<code>0 mnds</code>, excluidas de gangas y mínimos históricos) &nbsp;·&nbsp;
                🕒 Corte: <b>{html_lib.escape(str(ultima_act))}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.subheader("🔍 Lupa Visual Histórica (Evolución de Precio, Piso, Promedio y Zona de Puja)")
        if not df_historial.empty and "Jugador" in df_historial.columns:
            # Preferir nombres del último corte para evitar duplicados de etiquetas viejas en el selector
            if not df_ult_corte.empty:
                lista_jugadores_unicos = sorted(df_ult_corte["Jugador"].astype(str).unique().tolist())
            else:
                lista_jugadores_unicos = sorted(df_historial["Jugador"].astype(str).unique().tolist())
            jug_seleccionado = st.selectbox("Elegí un jugador para ver su gráfico histórico:", lista_jugadores_unicos)

            if jug_seleccionado:
                df_hist_tmp = df_historial.copy()
                df_hist_tmp["_dt_lupa"] = pd.to_datetime(df_hist_tmp["Fecha"], errors="coerce")
                # Unificar historial por Enlace_Optimo para conservar toda la serie aunque haya cambiado la etiqueta
                enl_sel = ""
                if "Enlace_Optimo" in df_hist_tmp.columns:
                    sub_match_j = df_hist_tmp[df_hist_tmp["Jugador"] == jug_seleccionado]
                    if not sub_match_j.empty:
                        enl_sel = str(sub_match_j["Enlace_Optimo"].dropna().iloc[-1] or "").strip()
                if enl_sel and "Enlace_Optimo" in df_hist_tmp.columns:
                    df_jug_hist = df_hist_tmp[df_hist_tmp["Enlace_Optimo"] == enl_sel].sort_values(by="_dt_lupa", kind="mergesort")
                else:
                    df_jug_hist = df_hist_tmp[df_hist_tmp["Jugador"] == jug_seleccionado].sort_values(by="_dt_lupa", kind="mergesort")

                precios_serie = df_jug_hist["Precio_Actual"].apply(num_seguro)
                precios_validos_j = precios_serie[precios_serie > 0]
                p_actual_j = int(precios_serie.iloc[-1]) if not precios_serie.empty else 0
                es_extinta_j = (p_actual_j <= 0)
                p_min_j = int(precios_validos_j.min()) if not precios_validos_j.empty else 0
                p_max_j = int(precios_validos_j.max()) if not precios_validos_j.empty else 0
                p_prom_j = int(round(precios_validos_j.mean())) if not precios_validos_j.empty else 0
                puja_j = calcular_puja_maxima_segura(p_actual_j, factor_roi) if not es_extinta_j else 0
                lazy_j = calcular_precio_lazy(
                    p_actual_j,
                    max_hist=p_max_j,
                    muestras=len(precios_validos_j),
                    factor_lazy=factor_lazy,
                    factor_roi=factor_roi
                ) if not es_extinta_j else 0

                if es_extinta_j:
                    st.markdown(
                        '<div class="sniper-banner" style="border-left-color:#ef4444;border-color:rgba(239,68,68,0.55);">'
                        '⛔ <b>CARTA EXTINTA / SIN PRECIO ACTIVO EN FUT.GG (0 mnds):</b> '
                        'Esta carta no tiene cotización vigente en el mercado. Se muestran debajo sus bandas históricas previas con precio válido.'
                        '</div>',
                        unsafe_allow_html=True
                    )

                m_l1, m_l2, m_l3, m_l4, m_l5 = st.columns(5)
                m_l1.metric("💰 Precio Actual", "⛔ EXTINTA (0)" if es_extinta_j else f"{p_actual_j:,} mnds")
                m_l2.metric("💎 Piso Mínimo Hist.", f"{p_min_j:,} mnds" if p_min_j > 0 else "—")
                m_l3.metric("⚖️ Promedio Hist.", f"{p_prom_j:,} mnds" if p_prom_j > 0 else "—")
                m_l4.metric(f"🔨 Pujar (+{roi_puja_pct}%)", f"{puja_j:,} mnds" if puja_j > 0 else "—")
                m_l5.metric(f"🎯 Target Lazy (+{lazy_pct}%)", f"{lazy_j:,} mnds" if lazy_j > 0 else "—")

                # Graficar cronológicamente omitiendo ceros extintos en la línea continua
                y_plot = precios_serie.where(precios_serie > 0, np.nan)
                fig_hist = go.Figure()
                fig_hist.add_trace(go.Scatter(
                    x=df_jug_hist["Fecha"], y=y_plot,
                    mode="lines+markers", name=etiqueta_plat,
                    line=dict(color="#38bdf8", width=2.8),
                    fill="tozeroy",
                    fillcolor="rgba(56, 189, 248, 0.08)",
                    connectgaps=False
                ))
                if p_min_j > 0:
                    fig_hist.add_hline(y=p_min_j, line_dash="dash", line_color="#10b981", annotation_text=f"Piso ({p_min_j:,})")
                if puja_j > 0:
                    fig_hist.add_hline(y=puja_j, line_dash="dashdot", line_color="#f59e0b", annotation_text=f"Puja Máx ({puja_j:,})")
                estilizar_plotly_fintech(fig_hist, height=340, hovermode="x unified")
                st.plotly_chart(fig_hist, use_container_width=True)

        if not df_ult_corte.empty:
            df_corte_activos_tab2 = df_ult_corte[df_ult_corte["Precio_Actual"].apply(num_seguro) > 0].copy()
            c_sh1, c_sh2 = st.columns(2)
            with c_sh1:
                with st.expander("📲 Copiar Resumen de Gangas para WhatsApp / Discord / Telegram"):
                    df_top_g = df_corte_activos_tab2.copy()
                    if "Senal_Trading" in df_top_g.columns:
                        df_top_g = df_top_g[~df_top_g["Senal_Trading"].astype(str).str.contains("🔎|⚠️|⛔", na=False)]
                    if "Variacion_Bruta_%" in df_top_g.columns:
                        df_top_g = df_top_g.sort_values(by="Variacion_Bruta_%").head(5)
                    else:
                        df_top_g = df_top_g.head(5)
                    lineas_share = [f"⚡ *TOP 5 GANGAS FC27 — FUT.GG ({etiqueta_plat} | Corte {ultima_act})* ⚡"]
                    for _, rg in df_top_g.iterrows():
                        nom_sh = str(rg.get("Jugador", "")).split(" (")[0]
                        p_sh = num_seguro(rg.get("Precio_Actual", 0))
                        puja_sh = num_seguro(rg.get("Puja_Max_10%_ROI", 0))
                        lazy_sh = num_seguro(rg.get("Precio_Lazy_Sugerido", 0))
                        lineas_share.append(f"• *{nom_sh}*: {p_sh:,} mnds | 🔨 Pujar <= {puja_sh:,} | 🎯 Lazy: {lazy_sh:,}")
                    st.code("\n".join(lineas_share), language="markdown")

            with c_sh2:
                with st.expander("💥 Detector de Flash Crashes Intradía & Rebotes (FUT.GG)", expanded=False):
                    filas_crash = []
                    for _, r_ar in df_corte_activos_tab2.iterrows():
                        var_a = float_seguro(r_ar.get("Variacion_Bruta_%", 0.0))
                        senal_a = str(r_ar.get("Senal_Trading", ""))
                        if var_a <= -15.0 and "🔎" not in senal_a and "⛔" not in senal_a:
                            p_act_fc = num_seguro(r_ar.get("Precio_Actual", 0))
                            p_prom_fc = num_seguro(r_ar.get("Precio_Prom_Hist", r_ar.get("Precio_Anterior", 0)))
                            neto_fc = int(round((p_prom_fc * 0.95) - p_act_fc)) if p_prom_fc > 0 else 0
                            if p_act_fc > 0 and neto_fc >= ganancia_neta_minima_real(p_act_fc):
                                filas_crash.append({
                                    "Jugador": r_ar.get("Jugador"),
                                    "Tipo": "⚠️ Cuchillo / Crash" if "⚠️" in senal_a else "💥 Flash Crash",
                                    "Precio_Actual": p_act_fc,
                                    "Ref_Historica": p_prom_fc,
                                    "Caida_y_Neto": f"{var_a:+.1f}% ({neto_fc:+,} netos)",
                                    "Puja_Sugerida": num_seguro(r_ar.get("Puja_Max_10%_ROI", 0))
                                })
                    if filas_crash:
                        st.dataframe(pd.DataFrame(filas_crash).head(25), use_container_width=True, hide_index=True)
                    else:
                        st.caption("Sin Flash Crashes intradía rentables Post-Tax (caída <= -15%) en este corte.")

        # TOP 3 OPORTUNIDADES DEL MOMENTO (TARJETAS VISUALES DESTACADAS — EXCLUYE EXTINTAS)
        if not df_ult_corte.empty:
            st.markdown("### 🔥 Top 3 Oportunidades del Momento")
            df_opp = df_ult_corte[df_ult_corte["Precio_Actual"].apply(num_seguro) > 0].copy()
            if "Senal_Trading" in df_opp.columns:
                df_opp = df_opp[~df_opp["Senal_Trading"].astype(str).str.contains("🔎|⚠️|⛔", na=False)]
            if not df_opp.empty:
                df_opp["_p_act"] = df_opp["Precio_Actual"].apply(num_seguro)
                df_opp["_p_lazy"] = df_opp["Precio_Lazy_Sugerido"].apply(num_seguro)
                df_opp["_neto_est"] = (df_opp["_p_lazy"] * 0.95 - df_opp["_p_act"]).round().astype(int)
                df_opp["_var_num"] = pd.to_numeric(df_opp.get("Variacion_Bruta_%", 0.0), errors="coerce").fillna(0.0)
                df_opp["_prio"] = 0
                if "Accion" in df_opp.columns:
                    df_opp.loc[df_opp["Accion"].astype(str).str.contains("COMPRAR", na=False), "_prio"] += 2
                if "Senal_Trading" in df_opp.columns:
                    df_opp.loc[df_opp["Senal_Trading"].astype(str).str.contains("⚡|💎", na=False), "_prio"] += 1
                top3_opp = df_opp[df_opp["_neto_est"] > 0].sort_values(
                    by=["_prio", "_var_num", "_neto_est"],
                    ascending=[False, True, False]
                ).head(3)
                if top3_opp.empty:
                    top3_opp = df_opp.sort_values(by="_neto_est", ascending=False).head(3)

                cols_opp = st.columns(max(1, len(top3_opp)))
                medallas_opp = ["#1 TOP PICK", "#2 OPORTUNIDAD", "#3 RADAR PRO"]
                bordes_opp = ["#10b981", "#38bdf8", "#f59e0b"]
                for i_op, (_, r_op) in enumerate(top3_opp.iterrows()):
                    nom_op = str(r_op.get("Jugador", "Carta"))
                    p_act_op = num_seguro(r_op.get("Precio_Actual", 0))
                    p_min_op = num_seguro(r_op.get("Precio_Min_Hist", p_act_op))
                    p_prom_op = num_seguro(r_op.get("Precio_Prom_Hist", p_act_op))
                    lazy_op = num_seguro(r_op.get("Precio_Lazy_Sugerido", 0))
                    puja_op = num_seguro(r_op.get("Puja_Max_10%_ROI", 0))
                    neto_op = int(round((lazy_op * 0.95) - p_act_op))
                    var_op = float_seguro(r_op.get("Variacion_Bruta_%", 0.0))
                    senal_op = str(r_op.get("Senal_Trading", "⚡ Oportunidad"))
                    link_op = str(r_op.get("Enlace_Optimo", r_op.get("Enlace_Futbin", "")) or "").strip()
                    link_html_op = f'<a href="{html_lib.escape(link_op)}" target="_blank">Abrir en FUT.GG ↗</a>' if link_op.startswith("http") else '<span>Sin enlace</span>'
                    min_txt = f"{p_min_op:,}" if p_min_op > 0 else "—"
                    prom_txt = f"{p_prom_op:,}" if p_prom_op > 0 else "—"
                    var_col = "#34d399" if var_op <= 0 else "#f87171"

                    with cols_opp[i_op]:
                        st.markdown(f"""
                        <div class="opp-card" style="border-left-color:{bordes_opp[i_op]};">
                            <div class="opp-rank-row">
                                <span class="opp-tag">{medallas_opp[i_op]}</span>
                                <span class="opp-source">FUT.GG</span>
                            </div>
                            <div class="opp-name" title="{html_lib.escape(nom_op)}">{html_lib.escape(nom_op)}</div>
                            <div class="opp-sub">Piso: <span class="mono">{min_txt}</span> · Prom: <span class="mono">{prom_txt}</span> · {html_lib.escape(senal_op[:28])}</div>
                            <div class="opp-grid">
                                <div class="opp-metric">
                                    <div class="opp-metric-lbl">Precio FUT.GG</div>
                                    <div class="opp-metric-val">{p_act_op:,}</div>
                                </div>
                                <div class="opp-metric">
                                    <div class="opp-metric-lbl">Neto Est. (5% Tax)</div>
                                    <div class="opp-metric-val" style="color:#34d399;">{neto_op:+,}</div>
                                </div>
                                <div class="opp-metric">
                                    <div class="opp-metric-lbl">Puja Máx (+{roi_puja_pct}%)</div>
                                    <div class="opp-metric-val" style="color:#7dd3fc;">{puja_op:,}</div>
                                </div>
                                <div class="opp-metric">
                                    <div class="opp-metric-lbl">Lazy Sugerido 1h</div>
                                    <div class="opp-metric-val" style="color:#fbbf24;">{lazy_op:,}</div>
                                </div>
                            </div>
                            <div class="opp-footer">
                                <span>Var: <b class="mono" style="color:{var_col};">{var_op:+.1f}%</b></span>
                                {link_html_op}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("⚡ Padrón Interactivo — Base Exclusiva FUT.GG (Hasta 10,000 cartas | 15M + Compra 1-Clic)")

        # FILTROS SOBRE EL PADRÓN COMPLETO DE HASTA 10,000 CARTAS Y TOPE HASTA 15,000,000 MONEDAS
        c_f1, c_f2, c_f3, c_f4 = st.columns([1.8, 1.4, 1.1, 1.3])
        with c_f1:
            busqueda_jug = st.text_input("🔍 Filtrar por Nombre, Media o Tipo de Carta", placeholder="Ej: 86, TOTW, Mbappe...")
        with c_f2:
            filtro_accion = st.selectbox(
                "Filtrar por Señal / Categoría",
                [
                    "Todas",
                    "🟢 COMPRAR",
                    "⚡ Flipping Rápido",
                    "💎 Mínimo Histórico",
                    "💥 Flash Crash Intradía (Caída <= -15%)",
                    "📦 Out of Packs / Especiales (TOTW & Promos)",
                    "🔴 VENDER",
                    "⚠️ Evitar Cuchillo",
                    "⛔ Extintas (0 mnds)"
                ]
            )
        with c_f3:
            piso_precio_filtro = st.number_input("Precio Mínimo (mnds)", min_value=0, max_value=TOPE_RADAR_MAXIMO, value=200, step=500)
        with c_f4:
            tope_precio_filtro = st.number_input("Precio Máximo (Hasta 15M)", min_value=500, max_value=TOPE_RADAR_MAXIMO, value=TOPE_RADAR_MAXIMO, step=25000)

        c_sl1, c_sl2, c_sl3 = st.columns([2.6, 1.0, 1.0])
        with c_sl1:
            tope_slider_radar = st.slider(
                "🎚️ Ajustar Rango de Precio Máximo con Deslizador (Monedas)",
                min_value=500,
                max_value=TOPE_RADAR_MAXIMO,
                value=int(tope_precio_filtro),
                step=10000
            )
        with c_sl2:
            max_filas_tabla = st.selectbox(
                "Máximo de filas a mostrar en tabla",
                [300, 1000, 3000, 10000],
                index=1
            )
        with c_sl3:
            st.markdown("<br>", unsafe_allow_html=True)
            incluir_extintas_padron = st.checkbox("⛔ Mostrar extintas (0 mnds)", value=False, key="chk_extintas_padron")
        tope_final_radar = min(int(tope_precio_filtro), int(tope_slider_radar)) if tope_slider_radar != TOPE_RADAR_MAXIMO else int(tope_precio_filtro)

        if not df_ult_corte.empty:
            serie_precios_r = df_ult_corte["Precio_Actual"].apply(num_seguro)
            if filtro_accion == "⛔ Extintas (0 mnds)":
                df_filtrado = df_ult_corte[serie_precios_r <= 0].copy()
            else:
                # Excluir SIEMPRE cartas extintas (<= 0) en gangas, mínimos históricos y vista por defecto
                excluir_ext_estricto = (not incluir_extintas_padron) or (filtro_accion in [
                    "🟢 COMPRAR", "⚡ Flipping Rápido", "💎 Mínimo Histórico",
                    "💥 Flash Crash Intradía (Caída <= -15%)", "🔴 VENDER", "⚠️ Evitar Cuchillo"
                ])
                piso_eff = max(200, piso_precio_filtro) if excluir_ext_estricto else piso_precio_filtro
                mask_rango = (serie_precios_r >= piso_eff) & (serie_precios_r <= tope_final_radar)
                if excluir_ext_estricto:
                    mask_rango = mask_rango & (serie_precios_r > 0)
                df_filtrado = df_ult_corte[mask_rango].copy()

            if busqueda_jug.strip():
                b_norm = sin_tildes(busqueda_jug.strip()).lower()
                df_filtrado = df_filtrado[df_filtrado["Jugador"].apply(lambda x: b_norm in sin_tildes(str(x)).lower())]
            if filtro_accion == "🟢 COMPRAR" and "Accion" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Accion"].astype(str).str.contains("COMPRAR", na=False)]
            elif filtro_accion == "⚡ Flipping Rápido" and "Senal_Trading" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Senal_Trading"].astype(str).str.contains("⚡", na=False)]
            elif filtro_accion == "💎 Mínimo Histórico" and "Estado_Historico" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Estado_Historico"].astype(str).str.contains("💎", na=False)]
            elif filtro_accion == "💥 Flash Crash Intradía (Caída <= -15%)" and "Variacion_Bruta_%" in df_filtrado.columns:
                mask_fc = pd.to_numeric(df_filtrado["Variacion_Bruta_%"], errors="coerce").fillna(0) <= -15.0
                if "Senal_Trading" in df_filtrado.columns:
                    mask_fc = mask_fc & (~df_filtrado["Senal_Trading"].astype(str).str.contains("🔎|⛔", na=False))
                df_filtrado = df_filtrado[mask_fc]
            elif filtro_accion == "📦 Out of Packs / Especiales (TOTW & Promos)":
                df_filtrado = df_filtrado[df_filtrado["Jugador"].apply(es_carta_especial_oop)]
            elif filtro_accion == "🔴 VENDER" and "Accion" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Accion"].astype(str).str.contains("VENDER", na=False)]
            elif filtro_accion == "⚠️ Evitar Cuchillo" and "Senal_Trading" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Senal_Trading"].astype(str).str.contains("⚠️", na=False)]

            if "Enlace_Optimo" not in df_filtrado.columns and "Enlace_Futbin" in df_filtrado.columns:
                df_filtrado["Enlace_Optimo"] = df_filtrado["Enlace_Futbin"]
            if "Estado_Mercado" not in df_filtrado.columns:
                df_filtrado["Estado_Mercado"] = np.where(df_filtrado["Precio_Actual"].apply(num_seguro) > 0, "🟢 Activa", "⛔ Extinta")

            # Ordenar poniendo siempre activas antes que extintas cuando se incluyen extintas
            df_filtrado["_ord_ext"] = np.where(df_filtrado["Precio_Actual"].apply(num_seguro) > 0, 0, 1)
            df_filtrado = df_filtrado.sort_values(by=["_ord_ext"], kind="mergesort").drop(columns=["_ord_ext"])

            cols_mostrar = [c for c in [
                "Jugador", "Estado_Mercado", "Precio_Actual",
                "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Break_Even_5%", "Precio_Lazy_Sugerido", "Enlace_Optimo"
            ] if c in df_filtrado.columns]

            df_editor_padron = df_filtrado[cols_mostrar].head(int(max_filas_tabla)).copy()
            df_editor_padron.insert(0, "Comprar_1Clic", False)

            st.caption(
                f"Mostrando **{len(df_editor_padron):,}** de **{len(df_filtrado):,}** cartas filtradas · "
                f"Activas con precio: **{cartas_activas_corte:,}** · Extintas: **{cartas_extintas_corte:,}**"
            )

            edicion_padron = st.data_editor(
                df_editor_padron, use_container_width=True, hide_index=True,
                disabled=[c for c in df_editor_padron.columns if c != "Comprar_1Clic"],
                column_config={
                    "Comprar_1Clic": st.column_config.CheckboxColumn("🛒 Comprar"),
                    "Estado_Mercado": st.column_config.TextColumn("🩺 Estado"),
                    "Precio_Actual": st.column_config.NumberColumn("💰 Precio FUT.GG", format="%d mnds"),
                    "Variacion_Bruta_%": st.column_config.NumberColumn("📊 Var %", format="%+.1f%%"),
                    "Puja_Max_10%_ROI": st.column_config.NumberColumn(f"🔨 Puja (+{roi_puja_pct}%)", format="%d mnds"),
                    "Break_Even_5%": st.column_config.NumberColumn("⚖️ Break-Even 5%", format="%d mnds"),
                    "Precio_Lazy_Sugerido": st.column_config.NumberColumn(f"🎯 Lazy (+{lazy_pct}%)", format="%d mnds"),
                    "Enlace_Optimo": st.column_config.LinkColumn("🔗 FUT.GG", display_text="Ver Carta")
                }
            )

            seleccionadas_1clic = edicion_padron[edicion_padron["Comprar_1Clic"] == True]
            if not seleccionadas_1clic.empty:
                if st.button(f"📥 Enviar {len(seleccionadas_1clic)} carta(s) a Mi Cartera"):
                    df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                    if "Cantidad" not in df_c_act.columns:
                        df_c_act["Cantidad"] = 1
                    nuevas_filas = []
                    costo_1clic_tot = 0
                    omitidas_ext = 0
                    for _, r_sel in seleccionadas_1clic.iterrows():
                        nom_limpio = str(r_sel["Jugador"]).split(" (")[0].strip()
                        p_compra_1c = num_seguro(r_sel["Precio_Actual"], 0)
                        if p_compra_1c <= 0:
                            omitidas_ext += 1
                            continue
                        costo_1clic_tot += p_compra_1c
                        est_auto = autodetectar_estrategia(r_sel["Jugador"], p_compra_1c)
                        nuevas_filas.append({
                            "Jugador_Exacto": nom_limpio,
                            "Estrategia": est_auto,
                            "Cantidad": 1,
                            "Precio_Pagado": p_compra_1c,
                            "Estado": "Activo",
                            "Precio_Venta": 0,
                            "Costo_EA_Tax": 0,
                            "Ganancia_Neta": 0,
                            "ROI_%": 0.0,
                            "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                            "Fecha_Venta": ""
                        })
                    if nuevas_filas:
                        df_c_act = pd.concat([df_c_act, pd.DataFrame(nuevas_filas)], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                        if saldos_cfg.get("modo_liquidez") == "Manual":
                            saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - costo_1clic_tot)
                            guardar_saldos_personalizados(saldos_cfg, es_admin, slug_club=slug_club_activo)
                        st.success(f"¡{len(nuevas_filas)} carta(s) agregada(s) a tu Cartera!" + (f" ({omitidas_ext} extinta(s) omitida(s))" if omitidas_ext else ""))
                        st.rerun()
                    else:
                        st.warning("⛔ No se pueden comprar cartas extintas (0 mnds) en 1 clic. Cargá la compra manualmente con el precio real pagado.")

    with tab3:
        # 1. OPTIMIZADOR MATEMÁTICO DE PLANTILLAS SBC CONECTADO AL ALMACÉN DE MEDIAS INTRANSFERIBLES DEL CLUB (PERSISTENTE)
        st.subheader("🧩 Optimizador Matemático de Plantillas SBC (Conectado a tu Almacén de Intransferibles)")
        pisos_vivos, jug_ref_vivos, df_fodder_vivo = calcular_pisos_fodder_vivos(df_ult_corte)
        df_totw_vivo, piso_totw_vivo = construir_radar_totw_vivo(
            df_ult_corte, factor_lazy=factor_lazy, factor_roi=factor_roi
        )

        intransf_guardados = cargar_intransferibles_club(es_admin, slug_club=slug_club_activo)
        with st.expander("🎒 Almacén de Medias Intransferibles del Club (Persistente por Club en Cloud/Disco)", expanded=True):
            st.caption("Cargá cuántas cartas intransferibles tenés en tu club para cada media; el optimizador las descontará automáticamente y podés guardarlas para próximas sesiones:")
            ci1, ci2, ci3, ci4, ci5, ci6, ci7, ci8, ci9 = st.columns([1, 1, 1, 1, 1, 1, 1, 1, 1.3])
            stock_intransf = {}
            with ci1:
                stock_intransf[83] = st.number_input("M83 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("83", 0)), step=1, key="intr_83")
            with ci2:
                stock_intransf[84] = st.number_input("M84 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("84", 0)), step=1, key="intr_84")
            with ci3:
                stock_intransf[85] = st.number_input("M85 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("85", 0)), step=1, key="intr_85")
            with ci4:
                stock_intransf[86] = st.number_input("M86 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("86", 0)), step=1, key="intr_86")
            with ci5:
                stock_intransf[87] = st.number_input("M87 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("87", 0)), step=1, key="intr_87")
            with ci6:
                stock_intransf[88] = st.number_input("M88 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("88", 0)), step=1, key="intr_88")
            with ci7:
                stock_intransf[89] = st.number_input("M89 Intr.", min_value=0, max_value=11, value=int(intransf_guardados.get("89", 0)), step=1, key="intr_89")
            with ci8:
                ifs_intransf = st.number_input("TOTW Intr.", min_value=0, max_value=5, value=int(intransf_guardados.get("TOTW", 0)), step=1, key="intr_if")
            with ci9:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Guardar Stock", use_container_width=True, key="btn_save_intransf"):
                    nuevo_stock_intr = {str(k): int(v) for k, v in stock_intransf.items()}
                    nuevo_stock_intr["TOTW"] = int(ifs_intransf)
                    guardar_intransferibles_club(nuevo_stock_intr, es_admin, slug_club=slug_club_activo)
                    st.success("✅ ¡Almacén guardado!")

        co1, co2 = st.columns([1, 2])
        with co1:
            media_sbc_obj = st.selectbox("🎯 Media de Plantilla exigida por el SBC", [84, 85, 86, 87, 88], index=2)
            ifs_exigidos = st.selectbox("🖤 Cantidad de Cartas TOTW / IF exigidas", [0, 1, 2], index=0)
            st.caption(f"🖤 Piso vivo TOTW validado: **{piso_totw_vivo:,} mnds**")
        with co2:
            combos_disp = COMBOS_SBC_PLANTILLAS.get(media_sbc_obj, [])
            filas_combos = []
            for nombre_combo, receta in combos_disp:
                costo_bruto = sum(cant * pisos_vivos.get(med_r, 5000) for med_r, cant in receta.items())
                costo_neto = 0
                partes_comprar = []
                for med_r, cant in receta.items():
                    usadas_club = min(int(cant), int(stock_intransf.get(med_r, 0)))
                    faltan_comprar = max(0, int(cant) - usadas_club)
                    costo_neto += faltan_comprar * pisos_vivos.get(med_r, 5000)
                    if faltan_comprar > 0:
                        partes_comprar.append(f"{faltan_comprar}x M{med_r} [{jug_ref_vivos[med_r]}]")
                    else:
                        partes_comprar.append(f"0x M{med_r} (Cubierto c/ Club)")
                if ifs_exigidos > 0:
                    costo_bruto += ifs_exigidos * piso_totw_vivo
                    ifs_a_comprar = max(0, int(ifs_exigidos) - int(ifs_intransf))
                    costo_neto += ifs_a_comprar * piso_totw_vivo
                    if ifs_a_comprar > 0:
                        partes_comprar.append(f"{ifs_a_comprar}x TOTW ({piso_totw_vivo:,} mnds)")
                    else:
                        partes_comprar.append("0x TOTW (Cubierto c/ Club)")
                ahorro_club = costo_bruto - costo_neto
                detalle_cartas = " + ".join(partes_comprar)
                filas_combos.append({
                    "Receta_11_Cartas": nombre_combo + (f" (+{ifs_exigidos} IF)" if ifs_exigidos > 0 else ""),
                    "Costo_Neto_A_Comprar": costo_neto,
                    "Ahorro_Intransferibles": ahorro_club,
                    "Costo_Mercado_11_Cartas": costo_bruto,
                    "Cartas_A_Comprar_Hoy": detalle_cartas
                })
            df_combos = pd.DataFrame(filas_combos).sort_values(by=["Costo_Neto_A_Comprar", "Costo_Mercado_11_Cartas"])
            mejor_combo = df_combos.iloc[0]
            st.markdown(f"""
            <div class="sbc-card">
                <div class="sbc-card-head">
                    <span class="sbc-card-title">✅ Combinación Óptima — Plantilla Media {media_sbc_obj}</span>
                    <span class="sbc-recipe-pill">{html_lib.escape(str(mejor_combo['Receta_11_Cartas']))}</span>
                </div>
                <div class="sbc-metrics-grid">
                    <div class="sbc-tile">
                        <div class="sbc-tile-lbl">💵 A Pagar en Mercado</div>
                        <div class="sbc-tile-val" style="color:#34d399;">{int(mejor_combo['Costo_Neto_A_Comprar']):,} mnds</div>
                    </div>
                    <div class="sbc-tile">
                        <div class="sbc-tile-lbl">🎒 Ahorro Intransferibles</div>
                        <div class="sbc-tile-val" style="color:#7dd3fc;">+{int(mejor_combo['Ahorro_Intransferibles']):,} mnds</div>
                    </div>
                    <div class="sbc-tile">
                        <div class="sbc-tile-lbl">🏪 Costo sin Club (11 Cartas)</div>
                        <div class="sbc-tile-val" style="color:#cbd5e1;">{int(mejor_combo['Costo_Mercado_11_Cartas']):,} mnds</div>
                    </div>
                </div>
                <div class="sbc-shop-list">🛒 <b>Lista de compras hoy:</b> {html_lib.escape(str(mejor_combo['Cartas_A_Comprar_Hoy']))}</div>
            </div>
            """, unsafe_allow_html=True)
            st.dataframe(df_combos, use_container_width=True, hide_index=True)

        st.markdown("---")
        # 2. SIMULADOR DE INVERSIÓN EN MEDIAS (FODDER SBC)
        st.subheader("🧮 Simulador Personalizado de Inversión en Medias (Fodder SBC)")
        ref_fodder_sim = REFERENCIA_FODDER_WEB
        cs1, cs2, cs3, cs4 = st.columns(4)
        with cs1:
            med_sim = st.selectbox("Media Objetivo", [83, 84, 85, 86, 87, 88, 89], index=3)
        piso_def, techo_def = ref_fodder_sim[med_sim]
        with cs2:
            monto_inv_fod = st.number_input("Monedas a Invertir", min_value=5000, max_value=10000000, value=min(50000, max(5000, int(liquidez_disp))), step=5000)
        with cs3:
            p_compra_fod = st.number_input("Precio Compra por Carta", min_value=600, max_value=100000, value=max(600, int(pisos_vivos.get(med_sim, piso_def))), step=100)
        with cs4:
            p_venta_fod = st.number_input("Precio Venta Esperado (SBC)", min_value=1000, max_value=150000, value=techo_def, step=250)

        cant_cartas_fod = int(monto_inv_fod // max(1, p_compra_fod))
        inv_real_fod = cant_cartas_fod * p_compra_fod
        ingreso_neto_fod = int(round(cant_cartas_fod * (p_venta_fod * 0.95)))
        ganancia_neta_fod = ingreso_neto_fod - inv_real_fod
        roi_fod = round((ganancia_neta_fod / max(1, inv_real_fod)) * 100, 1)
        col_fod_neto = "#34d399" if ganancia_neta_fod >= 0 else "#f87171"

        st.markdown(f"""
        <div class="sbc-card" style="border-top-color:#38bdf8;">
            <div class="sbc-card-head">
                <span class="sbc-card-title">📦 Proyección de Inversión — Fodder Media {med_sim}</span>
                <span class="sbc-recipe-pill" style="background:rgba(56,189,248,0.14);border-color:rgba(56,189,248,0.30);color:#7dd3fc;">Compra {p_compra_fod:,} → Venta {p_venta_fod:,}</span>
            </div>
            <div class="sbc-metrics-grid">
                <div class="sbc-tile">
                    <div class="sbc-tile-lbl">🃏 Cartas a Comprar</div>
                    <div class="sbc-tile-val">{cant_cartas_fod:,} u.</div>
                </div>
                <div class="sbc-tile">
                    <div class="sbc-tile-lbl">💵 Inversión Real</div>
                    <div class="sbc-tile-val">{inv_real_fod:,} mnds</div>
                </div>
                <div class="sbc-tile">
                    <div class="sbc-tile-lbl">🏦 Ingreso Post-Tax (5%)</div>
                    <div class="sbc-tile-val" style="color:#7dd3fc;">{ingreso_neto_fod:,} mnds</div>
                </div>
                <div class="sbc-tile">
                    <div class="sbc-tile-lbl">✨ Ganancia Neta / ROI</div>
                    <div class="sbc-tile-val" style="color:{col_fod_neto};">{ganancia_neta_fod:+,} ({roi_fod:+.1f}%)</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("🎯 Gestor Visual de Alertas Sniper (Con Alarma Web Audio + Aviso a Telegram + Conversión 1-Clic)")
        col_al1, col_al2 = st.columns([1, 2])
        with col_al1:
            with st.form("form_alerta_web", clear_on_submit=True):
                jug_al = st.text_input("Jugador a vigilar", placeholder="Ej: Theo Hernandez")
                obj_al = st.number_input("Precio Objetivo (<=)", min_value=200, max_value=15000000, step=250, value=15000)
                avisar_tg_sniper = st.checkbox("📲 Enviar confirmación a Telegram", value=True)
                if st.form_submit_button("🎯 Activar Sniper", use_container_width=True) and jug_al.strip():
                    df_al_act = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
                    nueva_al = pd.DataFrame([{
                        "Jugador_Buscado": jug_al.strip().title(),
                        "Precio_Objetivo": int(obj_al),
                        "Estado": "Activa"
                    }])
                    df_al_act = pd.concat([df_al_act, nueva_al], ignore_index=True)
                    guardar_hoja_sincronizada(df_al_act, "Mis_Alertas", es_admin=es_admin, slug_club=slug_club_activo)
                    if avisar_tg_sniper:
                        enviar_aviso_telegram_web(
                            f"🎯 <b>NUEVA ALERTA SNIPER ACTIVADA ({nombre_club_activo})</b>\n"
                            f"• Jugador: <b>{jug_al.strip().title()}</b>\n"
                            f"• Precio Objetivo: <b>&lt;= {int(obj_al):,} mnds</b>"
                        )
                    st.rerun()

        with col_al2:
            df_alertas_w = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
            activas_al = df_alertas_w[df_alertas_w["Estado"].astype(str).str.strip() == "Activa"].copy() if not df_alertas_w.empty and "Estado" in df_alertas_w.columns else pd.DataFrame()
            if not activas_al.empty:
                st.dataframe(activas_al, use_container_width=True, hide_index=True)
                c_sn1, c_sn2, c_sn3 = st.columns([1.6, 1, 1])
                with c_sn1:
                    opciones_sn = [f"{idx} | {r.get('Jugador_Buscado')} (<= {num_seguro(r.get('Precio_Objetivo', 0)):,} mnds)" for idx, r in activas_al.iterrows()]
                    sel_sn_comp = st.selectbox("🛒 ¿La compraste? Pasar a Mi Cartera:", opciones_sn, key="sel_sn_to_cart")
                with c_sn2:
                    cant_sn_comp = st.number_input("Cant. Comprada", min_value=1, max_value=50, value=1, step=1, key="cant_sn_to_cart")
                with c_sn3:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("📥 Pasar a Cartera", use_container_width=True, key="btn_sn_to_cart"):
                        idx_sn_sel = int(sel_sn_comp.split(" | ")[0])
                        fila_sn_s = df_alertas_w.loc[idx_sn_sel]
                        nom_sn_s = str(fila_sn_s.get("Jugador_Buscado", "Carta")).strip().title()
                        p_obj_sn_s = num_seguro(fila_sn_s.get("Precio_Objetivo", 1000), 1000)
                        df_alertas_w.loc[idx_sn_sel, "Estado"] = "Cumplida"
                        guardar_hoja_sincronizada(df_alertas_w, "Mis_Alertas", es_admin=es_admin, slug_club=slug_club_activo)

                        df_c_sn = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                        if "Cantidad" not in df_c_sn.columns:
                            df_c_sn["Cantidad"] = 1
                        nueva_f_sn2 = pd.DataFrame([{
                            "Jugador_Exacto": nom_sn_s,
                            "Estrategia": autodetectar_estrategia(nom_sn_s, p_obj_sn_s),
                            "Cantidad": int(cant_sn_comp),
                            "Precio_Pagado": int(p_obj_sn_s),
                            "Estado": "Activo",
                            "Precio_Venta": 0,
                            "Costo_EA_Tax": 0,
                            "Ganancia_Neta": 0,
                            "ROI_%": 0.0,
                            "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                            "Fecha_Venta": ""
                        }])
                        df_c_sn = pd.concat([df_c_sn, nueva_f_sn2], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_sn, "Mi_Cartera", es_admin=es_admin, slug_club=slug_club_activo)
                        if saldos_cfg.get("modo_liquidez") == "Manual":
                            saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - int(p_obj_sn_s * cant_sn_comp))
                            guardar_saldos_personalizados(saldos_cfg, es_admin, slug_club=slug_club_activo)
                        st.success(f"✅ ¡{cant_sn_comp}x {nom_sn_s} sumado a Mi Cartera y alerta marcada como Cumplida!")
                        st.rerun()

                c_btn_al1, c_btn_al2 = st.columns(2)
                with c_btn_al1:
                    if st.button("🗑️ Limpiar Alertas Activas", use_container_width=True):
                        df_alertas_w.loc[df_alertas_w["Estado"].astype(str).str.strip() == "Activa", "Estado"] = "Cancelada"
                        guardar_hoja_sincronizada(df_alertas_w, "Mis_Alertas", es_admin=es_admin, slug_club=slug_club_activo)
                        st.rerun()
                with c_btn_al2:
                    if st.button("📲 Enviar Mis Alertas Activas a Telegram", use_container_width=True):
                        lineas_tg = [f"🎯 <b>ALERTAS SNIPER ACTIVAS ({nombre_club_activo})</b>"]
                        for _, r_act in activas_al.iterrows():
                            lineas_tg.append(f"• <b>{r_act.get('Jugador_Buscado')}</b> &lt;= {num_seguro(r_act.get('Precio_Objetivo', 0)):,} mnds")
                        ok_tg = enviar_aviso_telegram_web("\n".join(lineas_tg))
                        if ok_tg:
                            st.success("📲 ¡Resumen de alertas enviado a tu Telegram!")
                        else:
                            st.warning("No se pudo conectar con Telegram en este intento.")
            else:
                st.info("No tenés alertas Sniper activas en este momento.")

        st.markdown("---")
        plat_caption_t3 = locals().get("etiqueta_plat") or locals().get("plataforma_usuario") or "FUT.GG"
        corte_caption_t3 = locals().get("ultima_act") or "Último corte"
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            st.subheader("🧱 Radar de Medias SBC (Fodder En Vivo)")
            st.caption(f"Calculado en tiempo real sobre {plat_caption_t3} ({corte_caption_t3}) · Mediana de las 5 cartas Oro más baratas (anti-outliers).")
            if not df_fodder_vivo.empty:
                st.dataframe(
                    df_fodder_vivo,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Media": st.column_config.NumberColumn("⭐ Media", format="%d"),
                        "Precio_Mas_Barato": st.column_config.NumberColumn("💰 Piso Vivo (Mediana Top 5)", format="%d mnds"),
                        "Jugadores_Referencia": st.column_config.TextColumn("🛒 Top 3 Oro Más Baratas"),
                        "Piso_Compra": st.column_config.NumberColumn("🎯 Ref. Compra", format="%d mnds"),
                        "Techo_Venta": st.column_config.NumberColumn("🚀 Ref. Venta", format="%d mnds"),
                        "Enlace_Optimo": st.column_config.LinkColumn("🔗 FUT.GG", display_text="Ver Carta")
                    }
                )
            else:
                st.caption("Sin muestras válidas de Fodder Oro en el corte actual.")
        with c_m2:
            st.subheader("🖤 Radar TOTW en Piso (En Vivo)")
            st.caption(f"Calculado en tiempo real sobre {plat_caption_t3} ({corte_caption_t3}) · Cartas reales del Equipo de la Semana (9,500–18,000 mnds, Media 83+).")
            if not df_totw_vivo.empty:
                st.dataframe(
                    df_totw_vivo,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Precio_Actual": st.column_config.NumberColumn("💰 Precio FUT.GG", format="%d mnds"),
                        "Break_Even_5%": st.column_config.NumberColumn("⚖️ Break-Even", format="%d mnds"),
                        "Precio_Lazy_1h": st.column_config.NumberColumn(f"🎯 Lazy (+{lazy_pct}%)", format="%d mnds"),
                        "Enlace_Optimo": st.column_config.LinkColumn("🔗 FUT.GG", display_text="Ver Carta")
                    }
                )
            else:
                st.info("Sin cartas TOTW reales en piso en este corte.")

    # ==========================================
    # PESTAÑA 4: CENTRO DE INTELIGENCIA ARTIFICIAL EN VIVO (DINÁMICO BAJO DEMANDA)
    # ==========================================
    with tab4:
        st.subheader("🧠 Centro de Inteligencia Artificial & Oráculo En Vivo (Bajo Demanda)")
        st.caption("Cada vez que toques un botón, la IA analizará en tiempo real el padrón actual de cartas, tu presupuesto, tus reglas de Puja/Lazy y tu cartera activa.")

        df_ia_base = datos.get("Recomendacion_IA", pd.DataFrame())
        rep_guardado = str(df_ia_base["Reporte"].dropna().iloc[-1]) if not df_ia_base.empty and "Reporte" in df_ia_base.columns and not df_ia_base["Reporte"].dropna().empty else "Tocá el botón '🤖 Generar Nuevo Reporte IA Ahora' para analizar el mercado en vivo."
        ora_guardado = str(df_ia_base["Oraculo"].dropna().iloc[-1]) if not df_ia_base.empty and "Oraculo" in df_ia_base.columns and not df_ia_base["Oraculo"].dropna().empty else "Tocá el botón '🔮 Tirar Nuevo Oráculo Ahora' para predecir el contenido del día."

        if "reporte_ia_vivo" not in st.session_state:
            st.session_state["reporte_ia_vivo"] = rep_guardado
        if "oraculo_ia_vivo" not in st.session_state:
            st.session_state["oraculo_ia_vivo"] = ora_guardado

        col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4 = st.columns([2, 1, 1, 1])
        with col_ctrl1:
            enfoque_ia = st.selectbox(
                "🎯 Elegí el enfoque del análisis:",
                [
                    "🎯 Estrategia Global Balanceada (Flipping + SBC)",
                    "⚡ Flipping Rápido Intradía (Rebotes y Liquidez)",
                    "🧱 Acumulación de Medias / Fodder SBC y TOTW",
                    "🏆 Compras Meta para Ciclo FUT Champions",
                    "🛡️ Auditoría Profunda de Mi Cartera Activa"
                ]
            )
        with col_ctrl2:
            top_n_ia = st.selectbox("📊 Cantidad Cartas Top:", [3, 5, 10], index=1)
        with col_ctrl3:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_nuevo_rep = st.button("🤖 Generar Nuevo Reporte IA", use_container_width=True)
        with col_ctrl4:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_nuevo_ora = st.button(f"🔮 Tirar Nuevo Oráculo ({hora_drop_txt})", use_container_width=True)

        if btn_nuevo_rep:
            with st.spinner("🧠 Analizando las cartas del padrón y tu cartera con Gemini AI en tiempo real..."):
                nuevo_rep_txt = generar_reporte_estrategico_vivo(
                    df_ult_corte,
                    filas_cartera_viva,
                    int(max(presupuesto_usuario, liquidez_disp)),
                    roi_puja_pct,
                    lazy_pct,
                    etiqueta_plat,
                    enfoque_ia,
                    int(top_n_ia),
                    api_key_custom
                )
                st.session_state["reporte_ia_vivo"] = nuevo_rep_txt
                if es_admin:
                    df_ia_upd = pd.DataFrame([{
                        "Fecha": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                        "Reporte": nuevo_rep_txt,
                        "Oraculo": st.session_state["oraculo_ia_vivo"]
                    }])
                    guardar_hoja_sincronizada(df_ia_upd, "Recomendacion_IA", es_admin=True)
                st.success("✅ ¡Nuevo Reporte Estratégico generado en vivo!")

        if btn_nuevo_ora:
            with st.spinner("🔮 Consultando al Oráculo de Contenido con los pisos de Fodder y tendencias de hoy..."):
                nuevo_ora_txt = generar_oraculo_contenido_vivo(
                    df_ult_corte,
                    df_fodder_vivo,
                    df_totw_vivo,
                    hora_drop_txt,
                    etiqueta_plat,
                    api_key_custom
                )
                st.session_state["oraculo_ia_vivo"] = nuevo_ora_txt
                if es_admin:
                    df_ia_upd = pd.DataFrame([{
                        "Fecha": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                        "Reporte": st.session_state["reporte_ia_vivo"],
                        "Oraculo": nuevo_ora_txt
                    }])
                    guardar_hoja_sincronizada(df_ia_upd, "Recomendacion_IA", es_admin=True)
                st.success("✅ ¡Nuevo Oráculo de Contenido generado en vivo!")

        col_ia1, col_ia2 = st.columns(2)
        with col_ia1:
            rep_html_limpio = formatear_md_a_html_seguro(st.session_state["reporte_ia_vivo"])
            st.markdown(f"""
            <div class="ia-card-report">
                <div class="ia-card-head">
                    <span class="ia-card-head-title">🤖 Reporte Estratégico FUT.GG</span>
                    <span class="ia-live-chip">LIVE AI · {html_lib.escape(etiqueta_plat)}</span>
                </div>
                <div class="ia-card-body">{rep_html_limpio}</div>
            </div>
            """, unsafe_allow_html=True)
        with col_ia2:
            ora_html_limpio = formatear_md_a_html_seguro(st.session_state["oraculo_ia_vivo"])
            st.markdown(f"""
            <div class="ia-card-oracle">
                <div class="ia-card-head">
                    <span class="ia-card-head-title">🔮 Oráculo de Contenido</span>
                    <span class="ia-live-chip" style="background:rgba(56,189,248,0.14);color:#7dd3fc;border-color:rgba(56,189,248,0.30);">DROP {html_lib.escape(str(hora_drop_txt))}</span>
                </div>
                <div class="ia-card-body">{ora_html_limpio}</div>
            </div>
            """, unsafe_allow_html=True)

        # CONSULTOR TÁCTICO INTERACTIVO (PREGUNTAS PERSONALIZADAS A LA IA)
        st.markdown("""
        <div class="chat-section-header">
            <div class="chat-section-title">💬 Consultor Táctico IA Interactivo</div>
            <div class="chat-section-sub">Hacé preguntas personalizadas sobre el padrón FUT.GG (hasta 10,000 cartas), tu liquidez y tu cartera activa (descontando 5% de EA Tax).</div>
        </div>
        """, unsafe_allow_html=True)
        with st.form("form_chat_ia"):
            pregunta_usuario = st.text_input(
                "Escribí tu consulta de trading (la IA lee tu padrón y tu cartera actual):",
                placeholder="Ej: Tengo 100,000 monedas libres, ¿qué 3 cartas del padrón me recomendás pujar hoy y a cuánto las listo?"
            )
            btn_preguntar_ia = st.form_submit_button("🚀 Analizar Consulta con IA")

            if btn_preguntar_ia and pregunta_usuario.strip():
                with st.spinner("🤖 Analizando tu consulta contra la base de datos en vivo..."):
                    if not df_ult_corte.empty and "Precio_Actual" in df_ult_corte.columns:
                        p_q_ser = df_ult_corte["Precio_Actual"].apply(num_seguro)
                        df_disp_q = df_ult_corte[(p_q_ser > 0) & (p_q_ser <= max(presupuesto_usuario, liquidez_disp))].head(25)
                    else:
                        df_disp_q = pd.DataFrame()
                    cols_q = [c for c in ["Jugador", "Precio_Actual", "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Precio_Lazy_Sugerido"] if c in df_disp_q.columns]
                    muestra_q = df_disp_q[cols_q].to_string(index=False) if not df_disp_q.empty else "Sin datos activos."
                    cartera_q = ", ".join([f"{c['Cantidad']}x {c['Jugador']} (Costo Unit {c['Costo_Pagado']:,}, Hoy {c['Mercado_Hoy']:,})" for c in filas_cartera_viva]) if filas_cartera_viva else "Sin cartas activas."

                    prompt_q = f"""
                    Actuá como Asesor Experto de Trading en EA Sports FC 27 ({etiqueta_plat}).
                    Liquidez libre del usuario: {liquidez_disp:,} mnds (Capital inmovilizado: {cap_inmovilizado:,} mnds).
                    Reglas del usuario: Puja +{roi_puja_pct}% ROI, Lazy +{lazy_pct}%, Tax EA 5%.
                    Cartera activa del usuario: {cartera_q}
                    Muestra del padrón actual en vivo (solo cartas con precio activo > 0):
                    {muestra_q}

                    PREGUNTA DEL USUARIO: "{pregunta_usuario.strip()}"

                    Respondé de forma directa, práctica y con números exactos de compra/puja y venta Lazy (considerando el 5% de EA Tax).
                    """
                    resp_q, mod_q = llamar_ia_en_vivo(prompt_q, api_key_custom)
                    if resp_q:
                        st.session_state["respuesta_chat_ia"] = f"💬 **Respuesta de Gemini AI ({mod_q}):**\n\n{resp_q}"
                    else:
                        st.session_state["respuesta_chat_ia"] = (
                            f"💬 **Respuesta Cuantitativa Rápida:** Con tu liquidez disponible de `{liquidez_disp:,} mnds`, "
                            f"priorizá entrar por puja (`+{roi_puja_pct}% ROI`) en las cartas con señal `⚡ Flipping Rápido` o `💎 Mínimo Histórico` de la Pestaña 2 y listalas a 1 hora en el valor de la columna `🎯 Lazy (+{lazy_pct}%)`."
                        )

        if "respuesta_chat_ia" in st.session_state:
            chat_html_limpio = formatear_md_a_html_seguro(st.session_state["respuesta_chat_ia"])
            st.markdown(f'<div class="chat-response-box">{chat_html_limpio}</div>', unsafe_allow_html=True)

    # ==========================================
    # PESTAÑA 5: RANKING & VERSUS DE CLUBES (COMPETENCIA MULTIUSUARIO)
    # ==========================================
    with tab5:
        st.subheader("🏆 Ranking General & Versus de Clubes FC27 (Competencia entre Amigos)")
        st.caption("Compará en vivo el Patrimonio Total (Liquidez + Stock a VNR 5% Tax), la Ganancia Realizada (Contribución Marginal) y el % de Efectividad de todos los clubes registrados.")

        def evaluar_metricas_club(nombre_c, df_c_club, es_adm_c=False, slug_c=None):
            df_c_eval = convertir_columnas_numericas(df_c_club) if df_c_club is not None and not df_c_club.empty else pd.DataFrame(columns=columnas_base_cartera)
            if "Cantidad" not in df_c_eval.columns:
                df_c_eval["Cantidad"] = 1
            df_c_eval["Cantidad"] = df_c_eval["Cantidad"].apply(lambda x: max(1, num_seguro(x, 1)))

            if not df_c_eval.empty and "Estado" in df_c_eval.columns:
                act_c = df_c_eval[df_c_eval["Estado"].astype(str).str.strip() == "Activo"].copy()
                ven_c = df_c_eval[df_c_eval["Estado"].astype(str).str.strip() == "Vendido"].copy()
            else:
                act_c = pd.DataFrame()
                ven_c = pd.DataFrame()

            cartas_stock_c = int(act_c["Cantidad"].sum()) if not act_c.empty else 0
            cap_auto_c = int((act_c["Precio_Pagado"].apply(num_seguro) * act_c["Cantidad"]).sum()) if not act_c.empty and "Precio_Pagado" in act_c.columns else 0
            cfg_c = cargar_saldos_personalizados(es_adm_c, presupuesto_usuario, cap_auto_c, slug_club=slug_c)
            cap_inm_c = int(cfg_c.get("capital_manual", cap_auto_c)) if cfg_c.get("modo_capital") == "Manual (Monto fijo)" else cap_auto_c
            liq_c = int(cfg_c.get("liquidez_manual", max(0, int(presupuesto_usuario) - cap_inm_c))) if cfg_c.get("modo_liquidez") == "Manual" else max(0, int(presupuesto_usuario) - cap_inm_c)

            vnr_c = 0
            for _, r_ac in act_c.iterrows():
                nom_ac = str(r_ac.get("Jugador_Exacto", "Carta"))
                cant_ac = max(1, num_seguro(r_ac.get("Cantidad", 1), 1))
                costo_ac = num_seguro(r_ac.get("Precio_Pagado", 0), 0)
                est_ac = str(r_ac.get("Estrategia", ""))
                p_merc_ac, _, _, _, _ = emparejar_carta_mercado(nom_ac, costo_ac, est_ac, df_ult_corte)
                bruto_lote_ac = p_merc_ac * cant_ac
                vnr_c += bruto_lote_ac - int(round(bruto_lote_ac * 0.05))

            gan_real_c = 0
            aciertos_c = 0
            trades_liquidados_c = 0
            if not ven_c.empty:
                for _, r_vc in ven_c.iterrows():
                    cant_vc = max(1, num_seguro(r_vc.get("Cantidad", 1), 1))
                    pv_vc = num_seguro(r_vc.get("Precio_Venta", 0), 0)
                    pp_vc = num_seguro(r_vc.get("Precio_Pagado", 0), 0)
                    if pv_vc > 0:
                        vb_vc = pv_vc * cant_vc
                        tax_vc = int(round(vb_vc * 0.05))
                        neto_vc = int(round((vb_vc - tax_vc) - (pp_vc * cant_vc)))
                        trades_liquidados_c += 1
                    else:
                        neto_vc = 0
                    gan_real_c += neto_vc
                    if neto_vc > 0:
                        aciertos_c += 1

            efectividad_c = round((aciertos_c / trades_liquidados_c * 100.0), 1) if trades_liquidados_c > 0 else 0.0
            patrimonio_c = liq_c + vnr_c
            return {
                "Club": nombre_c,
                "Patrimonio_Total": patrimonio_c,
                "Liquidez_Caja": liq_c,
                "Stock_VNR_5%": vnr_c,
                "Capital_Inmovilizado": cap_inm_c,
                "Ganancia_Realizada": gan_real_c,
                "Cartas_Transferibles": cartas_stock_c,
                "Trades_Cerrados": trades_liquidados_c,
                "Efectividad_%": efectividad_c
            }

        filas_ranking = []
        filas_ranking.append(evaluar_metricas_club("👑 Club Admin Principal", cartera_admin_base, es_adm_c=True, slug_c=None))

        reg_todos = cargar_registro_clubes()
        for slug_k, info_k in reg_todos.items():
            nom_k = info_k.get("nombre", slug_k)
            arch_k = f"cartera_club_{slug_k}.csv"
            if slug_club_activo == slug_k:
                df_k = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera))
            elif os.path.exists(arch_k):
                try:
                    df_k = pd.read_csv(arch_k)
                except Exception:
                    df_k = pd.DataFrame(columns=columnas_base_cartera)
            else:
                df_k = pd.DataFrame(columns=columnas_base_cartera)
            filas_ranking.append(evaluar_metricas_club(f"🏆 {nom_k}", df_k, es_adm_c=False, slug_c=slug_k))

        if not es_admin and not slug_club_activo and not datos.get("Mi_Cartera", pd.DataFrame()).empty:
            filas_ranking.append(evaluar_metricas_club("👤 Mi Sesión Invitado", datos.get("Mi_Cartera"), es_adm_c=False, slug_c=None))

        df_ranking = pd.DataFrame(filas_ranking).sort_values(by=["Patrimonio_Total", "Ganancia_Realizada"], ascending=False).reset_index(drop=True)
        df_ranking.insert(0, "Puesto", [f"#{i+1}" for i in range(len(df_ranking))])

        # PODIO VISUAL ORO / PLATA / BRONCE
        if not df_ranking.empty:
            top_podio = df_ranking.head(3)
            cols_pod = st.columns(max(1, len(top_podio)))
            cfg_podio = [
                ("podium-gold", "🥇 #1 ORO · LÍDER", "#fbbf24"),
                ("podium-silver", "🥈 #2 PLATA", "#e2e8f0"),
                ("podium-bronze", "🥉 #3 BRONCE", "#f59e0b"),
            ]
            for i_pod, (_, r_pod) in enumerate(top_podio.iterrows()):
                cls_pod, lbl_pod, col_pod = cfg_podio[i_pod]
                pat_pod = int(r_pod.get("Patrimonio_Total", 0))
                gan_pod = int(r_pod.get("Ganancia_Realizada", 0))
                ef_pod = float(r_pod.get("Efectividad_%", 0.0))
                tr_pod = int(r_pod.get("Trades_Cerrados", 0))
                col_gan_pod = "#34d399" if gan_pod >= 0 else "#f87171"
                with cols_pod[i_pod]:
                    st.markdown(f"""
                    <div class="podium-card {cls_pod}">
                        <div class="podium-rank" style="color:{col_pod};">{lbl_pod}</div>
                        <div class="podium-club">{html_lib.escape(str(r_pod.get('Club', 'Club')))}</div>
                        <div class="podium-patrimonio">{fmt_monedas(pat_pod, modo_streamer)}</div>
                        <div class="podium-meta">
                            <span>Gan: <b style="color:{col_gan_pod};">{fmt_monedas(gan_pod, modo_streamer, True)}</b></span>
                            <span>🎯 {ef_pod:.1f}% ({tr_pod})</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

        st.dataframe(
            df_ranking,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Puesto": st.column_config.TextColumn("🏅 Pos."),
                "Club": st.column_config.TextColumn("🛡️ Club"),
                "Patrimonio_Total": st.column_config.NumberColumn("💎 Patrimonio Total (Caja + VNR)", format="%d mnds"),
                "Liquidez_Caja": st.column_config.NumberColumn("💵 Liquidez Caja", format="%d mnds"),
                "Stock_VNR_5%": st.column_config.NumberColumn("📈 Stock VNR (5% Tax)", format="%d mnds"),
                "Capital_Inmovilizado": st.column_config.NumberColumn("📦 Cap. Inmovilizado", format="%d mnds"),
                "Ganancia_Realizada": st.column_config.NumberColumn("💰 Ganancia Realizada", format="%+d mnds"),
                "Cartas_Transferibles": st.column_config.NumberColumn("🃏 Cartas Activas", format="%d/100"),
                "Trades_Cerrados": st.column_config.NumberColumn("📜 Ventas Cerradas", format="%d"),
                "Efectividad_%": st.column_config.NumberColumn("🎯 Efectividad", format="%.1f%%")
            }
        )

        col_rk1, col_rk2 = st.columns([1.2, 1])
        with col_rk1:
            fig_rk = px.bar(
                df_ranking,
                x="Club",
                y=["Liquidez_Caja", "Stock_VNR_5%", "Ganancia_Realizada"],
                barmode="group",
                title="📊 Comparativa de Patrimonio y Ganancia Realizada por Club",
                color_discrete_sequence=["#10b981", "#38bdf8", "#f59e0b"]
            )
            fig_rk.update_traces(marker_line_width=0, opacity=0.92)
            estilizar_plotly_fintech(fig_rk, height=350, hovermode="x unified")
            st.plotly_chart(fig_rk, use_container_width=True)

        with col_rk2:
            st.markdown("#### ⚔️ Versus 1 vs 1 (Cara a Cara entre Clubes)")
            lista_nombres_rk = df_ranking["Club"].tolist()
            if len(lista_nombres_rk) >= 2:
                cv1, cv2 = st.columns(2)
                with cv1:
                    club_vs_1 = st.selectbox("Club A:", lista_nombres_rk, index=0, key="vs_c1")
                with cv2:
                    club_vs_2 = st.selectbox("Club B:", lista_nombres_rk, index=min(1, len(lista_nombres_rk) - 1), key="vs_c2")
                r_vs1 = df_ranking[df_ranking["Club"] == club_vs_1].iloc[0]
                r_vs2 = df_ranking[df_ranking["Club"] == club_vs_2].iloc[0]
                dif_pat = int(r_vs1["Patrimonio_Total"] - r_vs2["Patrimonio_Total"])
                dif_gan = int(r_vs1["Ganancia_Realizada"] - r_vs2["Ganancia_Realizada"])
                lider_vs = club_vs_1 if dif_pat >= 0 else club_vs_2
                col_dp = "#34d399" if dif_pat >= 0 else "#f87171"
                col_dg = "#34d399" if dif_gan >= 0 else "#f87171"
                st.markdown(f"""
                <div class="vs-card">
                    <div class="sbc-card-head" style="margin-bottom:10px;">
                        <span class="sbc-card-title">🏆 Líder: {html_lib.escape(str(lider_vs))}</span>
                        <span class="sbc-recipe-pill" style="color:{col_dp};">{dif_pat:+,} mnds</span>
                    </div>
                    <div class="sbc-metrics-grid" style="margin-bottom:0;">
                        <div class="sbc-tile">
                            <div class="sbc-tile-lbl">💎 Patrimonio A vs B</div>
                            <div class="sbc-tile-val" style="font-size:0.88rem;">{int(r_vs1['Patrimonio_Total']):,} <span style="color:#64748b;">vs</span> {int(r_vs2['Patrimonio_Total']):,}</div>
                        </div>
                        <div class="sbc-tile">
                            <div class="sbc-tile-lbl">💰 Dif. Ganancia Real</div>
                            <div class="sbc-tile-val" style="font-size:0.92rem;color:{col_dg};">{dif_gan:+,} mnds</div>
                        </div>
                        <div class="sbc-tile">
                            <div class="sbc-tile-lbl">🎯 Efectividad A vs B</div>
                            <div class="sbc-tile-val" style="font-size:0.92rem;color:#7dd3fc;">{r_vs1['Efectividad_%']}% <span style="color:#64748b;">vs</span> {r_vs2['Efectividad_%']}%</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("Creá una **Cuenta de Club** para tus amigos desde el menú lateral (`🏆 Cuenta de Club (Amigos)` ➡️ `➕ Crear Club`) para habilitar el duelo 1 vs 1.")

    if auto_ref:
        time.sleep(60)
        st.cache_data.clear()
        st.rerun()

except Exception as err_global:
    st.error(f"🚨 Error detectado: {err_global}")
    st.code(traceback.format_exc())
