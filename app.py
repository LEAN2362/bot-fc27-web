import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import io
import json
import math
import time
import wave
import base64
import struct
import requests
import traceback
from urllib.parse import quote
from datetime import datetime, timedelta, timezone

# ==========================================
# CONFIGURACIÓN DE PÁGINA Y ESTÉTICA FULL COLOR
# ==========================================
st.set_page_config(
    page_title="FC27 Trading Terminal Pro (Cloud)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

LINK_GOOGLE_SHEET = ""
PIN_ADMIN_SECRETO = "2362"

_G_PART1 = "AQ.Ab8RN6KGQFnimiJeqwsb"
_G_PART2 = "KKcbIWeCUq6SggjfhQ6kIN9aKIFafA"
DEFAULT_GEMINI_KEY = os.environ.get("GEMINI_API_KEY", _G_PART1 + _G_PART2)
MODELOS_PREFERIDOS_IA = ["models/gemini-2.5-flash", "gemini-2.5-flash", "models/gemini-3.5-flash", "gemini-3.5-flash", "gemini-2.0-flash"]

PRESUPUESTO_DEFAULT = 200000
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
    .stApp {
        background: radial-gradient(circle at top right, #1a103c 0%, #0b0f19 55%, #070a12 100%);
        color: #f0f6fc;
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #111625 0%, #191233 100%) !important;
        border-right: 1px solid #2e3c56;
    }
    section[data-testid="stSidebar"] * { color: #e6edf3 !important; }
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background-color: #1e2638 !important;
        color: #ffffff !important;
        border: 1px solid #3b4c6b !important;
    }
    div.stButton > button, div[data-testid="stFormSubmitButton"] > button, div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(90deg, #00c853 0%, #00e676 50%, #00b0ff 100%) !important;
        color: #070a12 !important;
        font-weight: 800 !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 10px 16px !important;
        box-shadow: 0 4px 15px rgba(0, 230, 118, 0.35) !important;
    }
    .main-header {
        background: linear-gradient(120deg, #0f2027 0%, #1c1042 50%, #004d40 100%);
        padding: 22px 28px;
        border-radius: 16px;
        border: 1px solid #00e676;
        box-shadow: 0 8px 25px rgba(0, 230, 118, 0.18);
        margin-bottom: 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
    }
    .main-header h1 {
        margin: 0;
        font-size: 1.95rem;
        font-weight: 900;
        background: linear-gradient(90deg, #00e676, #00e5ff, #ffea00);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .main-header p { margin: 6px 0 0 0; color: #e2e8f0; font-size: 0.96rem; }
    .drop-timer {
        background: rgba(0, 0, 0, 0.45);
        border: 1px solid #00e5ff;
        padding: 10px 18px;
        border-radius: 12px;
        text-align: right;
    }
    .drop-timer-title { font-size: 0.78rem; color: #84ffff; text-transform: uppercase; font-weight: 700; }
    .drop-timer-val { font-size: 1.35rem; font-weight: 900; color: #ffea00; }
    .kpi-card {
        padding: 16px 14px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 8px 20px rgba(0,0,0,0.35);
        margin-bottom: 10px;
    }
    .kpi-title { font-size: 0.80rem; font-weight: 700; text-transform: uppercase; opacity: 0.92; margin-bottom: 6px; }
    .kpi-value { font-size: 1.35rem; font-weight: 900; margin-bottom: 8px; white-space: nowrap; }
    .kpi-badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; background: rgba(0, 0, 0, 0.35); }
    .kpi-green { background: linear-gradient(135deg, #054f31 0%, #00a86b 100%); border: 1px solid #00e676; }
    .kpi-blue { background: linear-gradient(135deg, #0d2b6b 0%, #1e62d0 100%); border: 1px solid #00e5ff; }
    .kpi-purple { background: linear-gradient(135deg, #3b0d6b 0%, #7b2cbf 100%); border: 1px solid #e040fb; }
    .kpi-gold { background: linear-gradient(135deg, #6b4c0d 0%, #d48806 100%); border: 1px solid #ffd600; }
    .kpi-crimson { background: linear-gradient(135deg, #5c0d28 0%, #c9184a 100%); border: 1px solid #ff4d6d; }
    .calc-box {
        background: rgba(15, 32, 39, 0.75);
        border: 1px solid #00e676;
        border-radius: 10px;
        padding: 12px;
        margin-top: 8px;
        font-size: 0.88rem;
    }
    .ia-card-report {
        background: linear-gradient(135deg, rgba(5, 79, 49, 0.45) 0%, rgba(17, 22, 37, 0.9) 100%);
        border: 1px solid #00e676;
        border-radius: 14px;
        padding: 18px 22px;
        margin-top: 10px;
        line-height: 1.55;
    }
    .ia-card-oracle {
        background: linear-gradient(135deg, rgba(59, 13, 107, 0.45) 0%, rgba(17, 22, 37, 0.9) 100%);
        border: 1px solid #00e5ff;
        border-radius: 14px;
        padding: 18px 22px;
        margin-top: 10px;
        line-height: 1.55;
    }
    .mode-badge-admin {
        background: rgba(0, 230, 118, 0.2);
        border: 1px solid #00e676;
        color: #00e676;
        padding: 6px 10px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 8px;
    }
    .mode-badge-guest {
        background: rgba(0, 229, 255, 0.18);
        border: 1px solid #00e5ff;
        color: #84ffff;
        padding: 6px 10px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 8px;
    }
    div[data-baseweb="tab-list"] { gap: 8px; background-color: #111625; padding: 8px; border-radius: 12px; border: 1px solid #26334d; }
    button[data-baseweb="tab"] { font-size: 0.95rem !important; font-weight: 700 !important; color: #94a3b8 !important; border-radius: 8px !important; padding: 10px 16px !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background: linear-gradient(90deg, #6200ea, #00b0ff) !important; color: #ffffff !important; }
    h2, h3 { color: #38bdf8 !important; font-weight: 800 !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIONES MATEMÁTICAS, HORA Y SALDOS PERSONALIZABLES
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

def ahora_arg():
    tz_arg = timezone(timedelta(hours=-3))
    return datetime.now(tz_arg).replace(tzinfo=None)

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

def redondear_escalon_ea(precio, hacia_arriba=False):
    precio = float_seguro(precio, 0.0)
    if precio <= 0:
        return 0
    if precio < 1000:
        paso = 50
    elif precio < 10000:
        paso = 100
    elif precio < 50000:
        paso = 250
    elif precio < 100000:
        paso = 500
    else:
        paso = 1000
    if hacia_arriba:
        return int(math.ceil(precio / paso) * paso)
    return int(math.floor(precio / paso) * paso)

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

def cargar_saldos_personalizados(es_admin, cap_auto_cartera):
    """Carga la Liquidez Disponible y el Capital Inmovilizado configurados por el usuario."""
    defecto = {
        "modo_liquidez": "Auto",
        "liquidez_manual": max(0, PRESUPUESTO_DEFAULT - cap_auto_cartera),
        "modo_capital": "Auto (Suma de mis cartas)",
        "capital_manual": cap_auto_cartera
    }
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

def guardar_saldos_personalizados(nuevos_saldos, es_admin):
    if not es_admin:
        st.session_state["guest_saldos"] = nuevos_saldos
        return
    st.session_state["admin_saldos"] = nuevos_saldos
    try:
        with open(ARCHIVO_SALDOS_LOCAL, "w", encoding="utf-8") as f:
            json.dump(nuevos_saldos, f)
    except Exception:
        pass

# ==========================================
# HERRAMIENTAS DE MERCADO FC27 Y CONTROL REMOTO
# ==========================================
def generar_beep_wav_base64():
    buf = io.BytesIO()
    sample_rate = 22050
    duration = 0.35
    freq = 880.0
    n_samples = int(sample_rate * duration)
    with wave.open(buf, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        for i in range(n_samples):
            val = int(16000 * math.sin(2 * math.pi * freq * (i / sample_rate)))
            wf.writeframesraw(struct.pack('<h', val))
    return base64.b64encode(buf.getvalue()).decode('ascii')

def disparar_github_action_remoto(gh_token, repo_owner="LEAN2362", repo_name="bot-fc27-web", workflow_file="escanear.yml"):
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/actions/workflows/{workflow_file}/dispatches"
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {gh_token.strip()}",
        "X-GitHub-Api-Version": "2022-11-28"
    }
    try:
        r = requests.post(url, headers=headers, json={"ref": "main"}, timeout=10)
        if r.status_code in [200, 204]:
            return True, "🚀 ¡Orden enviada! El servidor de GitHub Actions ya está escaneando el mercado."
        return False, f"⚠️ GitHub devolvió código {r.status_code}: verificá que tu Token tenga permiso 'workflow'."
    except Exception as e:
        return False, f"⚠️ Error de conexión con GitHub: {e}"

def exportar_contabilidad_excel(df_er, df_viva, df_vendidos):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_er.to_excel(writer, index=False, sheet_name="Estado_Resultados")
        df_viva.to_excel(writer, index=False, sheet_name="Cartera_Activa")
        df_vendidos.to_excel(writer, index=False, sheet_name="Ventas_Cerradas")
    return output.getvalue()

def optimizar_plantilla_sbc(media_objetivo, df_corte):
    pisos_defecto = {82: 750, 83: 900, 84: 1200, 85: 2800, 86: 5500, 87: 9500, 88: 14500, 89: 22000, 90: 32000}
    precios_reales = {}
    jugadores_ref = {}

    for m in range(82, 91):
        if not df_corte.empty and "Jugador" in df_corte.columns:
            cands = df_corte[df_corte["Jugador"].astype(str).str.contains(f"Media {m}", na=False) & (df_corte["Precio_Actual"].apply(num_seguro) > 0)]
            if not cands.empty:
                fila_min = cands.loc[cands["Precio_Actual"].apply(num_seguro).idxmin()]
                precios_reales[m] = num_seguro(fila_min["Precio_Actual"])
                jugadores_ref[m] = f"{str(fila_min['Jugador']).split(' (')[0]} ({precios_reales[m]:,} mnds en {fila_min.get('Fuente_Mejor_Precio', 'PC')})"
                continue
        precios_reales[m] = pisos_defecto.get(m, 5000)
        jugadores_ref[m] = f"Media {m} (~{precios_reales[m]:,} mnds)"

    r = int(media_objetivo)
    recetas = [
        {"Nombre": f"Opción A (Pareja): 11x Media {r}", "Combo": {r: 11}},
        {"Nombre": f"Opción B (2 Altas + Base): 2x Media {r+2} + 9x Media {r-1}", "Combo": {r+2: 2, r-1: 9}},
        {"Nombre": f"Opción C (1 Alta + 2 Medias + Base): 1x {r+2} + 2x {r+1} + 8x {r-1}", "Combo": {r+2: 1, r+1: 2, r-1: 8}},
        {"Nombre": f"Opción D (Escalonada): 3x Media {r+1} + 4x Media {r} + 4x Media {r-1}", "Combo": {r+1: 3, r: 4, r-1: 4}},
    ]

    resultados = []
    for rec in recetas:
        costo_tot = sum(cant * precios_reales.get(med, 10000) for med, cant in rec["Combo"].items())
        detalle = " + ".join([f"{cant}x {jugadores_ref.get(med, f'Media {med}')}" for med, cant in rec["Combo"].items()])
        resultados.append({
            "Fórmula SBC (11 Cartas)": rec["Nombre"],
            "Costo_Total": int(costo_tot),
            "Cartas Más Baratas Hoy en Padrón": detalle
        })

    return pd.DataFrame(resultados).sort_values(by="Costo_Total")

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

def generar_reporte_estrategico_vivo(df_corte, filas_cartera, liquidez_disp, roi_pct, lazy_pct, plataforma, enfoque, api_key_usar):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_hoy = dias[ahora.weekday()]
    hora_str = ahora.strftime("%H:%M")

    tope_busq = max(50000, liquidez_disp)
    df_disp = df_corte[df_corte["Precio_Actual"].apply(num_seguro) <= tope_busq].copy() if not df_corte.empty else pd.DataFrame()
    if df_disp.empty:
        df_disp = df_corte.copy()

    df_compras = df_disp[df_disp["Accion"].astype(str).str.contains("COMPRAR", na=False)].sort_values(by="Variacion_Bruta_%").head(15) if "Accion" in df_disp.columns else df_disp.head(15)
    df_ventas = df_disp[df_disp["Accion"].astype(str).str.contains("VENDER", na=False)].sort_values(by="Variacion_Neta_%", ascending=False).head(10) if "Accion" in df_disp.columns else pd.DataFrame()
    df_muestra = pd.concat([df_compras, df_ventas, df_disp.head(15)]).drop_duplicates(subset=["Jugador"]).head(28)

    cols_res = [c for c in ["Jugador", "Precio_Actual", "Fuente_Mejor_Precio", "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Precio_Lazy_Sugerido", "Accion"] if c in df_muestra.columns]
    tabla_txt = df_muestra[cols_res].to_string(index=False) if not df_muestra.empty else "Sin datos en el corte actual."

    cartera_txt = "Sin cartas activas en cartera."
    if filas_cartera:
        lineas_c = [
            f"- {c['Jugador']} ({c['Estrategia']}): Costo {c['Costo_Pagado']:,} | Hoy {c['Mercado_Hoy']:,} | Lazy {c['Precio_Lazy_1h']:,} | Neto {c['Resultado_Tenencia']:+,} ({c['ROI_Latente_%']:+.1f}%)"
            for c in filas_cartera
        ]
        cartera_txt = "\n".join(lineas_c)

    prompt = f"""
    Actuá como Trader Cuantitativo Principal de EA Sports FC 27 para la plataforma {plataforma}.
    Momento del análisis: {dia_hoy} {hora_str} hs (ARG).
    Liquidez Disponible del usuario: {liquidez_disp:,} monedas | Objetivo Puja: +{roi_pct}% ROI | Recargo Lazy Buyer: +{lazy_pct}% | Tax EA: 5%.
    ENFOQUE SOLICITADO: {enfoque}.

    CARTERA ACTIVA DEL USUARIO:
    {cartera_txt}

    MUESTRA EN VIVO DEL MERCADO:
    {tabla_txt}

    Redactá un REPORTE ESTRATÉGICO EN VIVO claro, directo y accionable (usando viñetas, sin tablas '|'):
    1. 🎯 TOP 5 OPORTUNIDADES DE COMPRA AHORA ({enfoque}): Nombre, Precio Actual, Fuente (FB o GG), hasta cuánto PUJAR (+{roi_pct}% ROI) y a cuánto listar para LAZY BUYERS (+{lazy_pct}%).
    2. 💼 DIAGNÓSTICO DE TU CARTERA ACTIVA: Qué hacer hoy con cada carta activa (o cómo invertir la liquidez).
    3. 🧭 PLAN DE ACCIÓN PARA HOY ({dia_hoy}): Estrategia según el ciclo semanal de Rivals / FUT Champions.
    """

    texto_ia, modelo_usado = llamar_ia_en_vivo(prompt, api_key_usar)
    if texto_ia:
        return f"🤖 **REPORTE IA EN VIVO ({dia_hoy} {hora_str} hs | Enfoque: {enfoque} | Modelo: {modelo_usado})**\n\n" + texto_ia

    lineas_fb = [
        f"🤖 **REPORTE TÁCTICO CUANTITATIVO EN VIVO ({dia_hoy} {hora_str} hs | Enfoque: {enfoque})**",
        f"*(Generado en tiempo real sobre {len(df_corte):,} cartas con Liquidez: {liquidez_disp:,} mnds)*\n",
        "### 🎯 Top 5 Oportunidades Detectadas:"
    ]
    for _, r in df_muestra.head(5).iterrows():
        nom = str(r.get("Jugador", "Carta")).split(" (")[0]
        p_a = num_seguro(r.get("Precio_Actual", 0))
        f_a = str(r.get("Fuente_Mejor_Precio", "PC"))
        puja_a = num_seguro(r.get("Puja_Max_10%_ROI", 0))
        lazy_a = num_seguro(r.get("Precio_Lazy_Sugerido", 0))
        var_a = float_seguro(r.get("Variacion_Bruta_%", 0))
        lineas_fb.append(f"- **{nom}** (`{p_a:,} mnds` en `{f_a}` | `{var_a:+.1f}%`) ➡️ 🔨 Pujar hasta `{puja_a:,}` | 🎯 Listar Lazy a `{lazy_a:,}`")
    return "\n".join(lineas_fb)

def generar_oraculo_contenido_vivo(df_corte, hora_drop_txt, plataforma, api_key_usar):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_hoy = dias[ahora.weekday()]
    hora_str = ahora.strftime("%H:%M")
    pauta_dia = AGENDA_EA_SEMANAL.get(dia_hoy, "Contenido diario de SBCs y Evoluciones.")

    subidas = []
    fodder_lista = []
    if not df_corte.empty:
        df_sub = df_corte.sort_values(by="Variacion_Bruta_%", ascending=False).head(6) if "Variacion_Bruta_%" in df_corte.columns else pd.DataFrame()
        for _, r in df_sub.iterrows():
            if float_seguro(r.get("Variacion_Bruta_%", 0)) > 0:
                subidas.append(f"{str(r['Jugador']).split(' (')[0]} ({num_seguro(r['Precio_Actual']):,} mnds, +{float_seguro(r['Variacion_Bruta_%']):.1f}%)")

        for med in [84, 85, 86, 87, 88, 89]:
            cands = df_corte[df_corte["Jugador"].astype(str).str.contains(f"Media {med}", na=False) & (df_corte["Precio_Actual"].apply(num_seguro) > 0)]
            if not cands.empty:
                min_row = cands.loc[cands["Precio_Actual"].apply(num_seguro).idxmin()]
                fodder_lista.append(f"Media {med}: {str(min_row['Jugador']).split(' (')[0]} a {num_seguro(min_row['Precio_Actual']):,} mnds [{min_row.get('Fuente_Mejor_Precio', 'PC')}]")

    prompt = f"""
    Actuá como el Oráculo de Contenido, SBCs y Filtraciones (Leaks) de EA Sports FC 27 ({plataforma}).
    Hoy es {dia_hoy}, hora actual: {hora_str} hs (ARG). Horario del Drop diario: {hora_drop_txt}.
    Pauta de EA para {dia_hoy}: {pauta_dia}
    Cartas con mayor subida hoy: {', '.join(subidas) if subidas else 'Mercado estable'}
    Precios piso de Medias (Fodder SBC) ahora mismo: {', '.join(fodder_lista) if fodder_lista else 'Sin muestra'}

    Redactá un DICTAMEN DEL ORÁCULO DE CONTENIDO dinámico y directo (sin tablas '|'):
    1. 🔮 PREDICCIÓN DEL DROP DE HOY ({dia_hoy} - {hora_drop_txt}): Qué SBCs o Evoluciones salen hoy y qué requisitos subirán.
    2. 🛒 LAS 3 COMPRAS ANTICIPADAS DEL ORÁCULO: Qué 3 medias, TOTW o cartas conviene acumular AHORA MISMO y a qué precio máximo.
    3. ⏰ TIMING DE ENTRADA Y SALIDA HOY: A qué hora comprar en pánico y a qué hora vender.
    """

    texto_ia, modelo_usado = llamar_ia_en_vivo(prompt, api_key_usar)
    if texto_ia:
        return f"🔮 **ORÁCULO DE CONTENIDO EN VIVO ({dia_hoy} {hora_str} hs | Drop: {hora_drop_txt} | Modelo: {modelo_usado})**\n\n" + texto_ia

    lineas_or = [
        f"🔮 **ORÁCULO DE CONTENIDO EN VIVO ({dia_hoy} {hora_str} hs | Drop: {hora_drop_txt})**\n",
        f"### 📅 Pauta Oficial esperada para hoy ({dia_hoy}):\n- {pauta_dia}\n",
        "### 🧱 Piso Actual de Medias SBC (Fodder en Vivo):"
    ]
    for f_it in (fodder_lista or ["Consultá la pestaña Simulador Fodder para ver los pisos por media."]):
        lineas_or.append(f"- {f_it}")
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
    if "STOP-LOSS" in s or "🛑" in s or "⚠️" in s:
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
    cols_num = [
        "Precio_Pagado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%",
        "Precio_Anterior", "Precio_Actual", "Precio_Consola", "Precio_Futbin", "Precio_FUTGG",
        "Precio_Min_Hist", "Precio_Prom_Hist", "Puja_Max_10%_ROI", "Break_Even_5%",
        "Precio_Lazy_Sugerido", "Variacion_Bruta_%", "Beneficio_Neto_Mnds", "Variacion_Neta_%",
        "Precio_Objetivo", "Margen_Neto_5%", "Brecha_%", "Subida_%", "Ganancia_Potencial_Neta",
        "Precio_Mas_Barato", "Piso_Compra", "Techo_Venta", "Precio_Base", "Precio_Explosivo"
    ]
    for c in cols_num:
        if c in df.columns:
            s = df[c].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.strip()
            df[c] = pd.to_numeric(s, errors="coerce").fillna(0)
    return df

def limpiar_escaneos_corruptos(df):
    if df is None or df.empty or "Jugador" not in df.columns or "Precio_Actual" not in df.columns:
        return pd.DataFrame()
    mask_corrupta = (
        df["Jugador"].astype(str).str.contains(r"Media \?\?", na=False)
        & (df["Precio_Actual"].apply(num_seguro) == 36000)
    )
    return df[~mask_corrupta].copy()

@st.cache_data(ttl=30)
def cargar_por_id_publico(sheet_id):
    hojas = {}
    nombres_hojas = [
        "Historial_Tendencias", "Mi_Cartera", "Recomendacion_IA",
        "Dashboard_Financiero", "Radar_Fodder", "Arbitraje_FUTGG",
        "Radar_Evos", "Radar_TOTW", "Radar_Links_Domino", "Mis_Alertas"
    ]
    df_base_comparacion = None

    for h in nombres_hojas:
        url_csv = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(h)}"
        try:
            df_h = pd.read_csv(url_csv)
            df_h = df_h.loc[:, ~df_h.columns.str.contains("^Unnamed")]
            if h == "Historial_Tendencias":
                df_base_comparacion = df_h.copy()
                hojas[h] = limpiar_escaneos_corruptos(convertir_columnas_numericas(df_h))
            else:
                if (
                    df_base_comparacion is not None
                    and list(df_h.columns) == list(df_base_comparacion.columns)
                    and len(df_h) == len(df_base_comparacion)
                ):
                    hojas[h] = pd.DataFrame()
                else:
                    hojas[h] = convertir_columnas_numericas(df_h)
        except Exception:
            hojas[h] = pd.DataFrame()

    if not hojas.get("Historial_Tendencias", pd.DataFrame()).empty:
        return hojas, "OK"
    return {}, "No se pudo leer 'Historial_Tendencias'. Verificá que en 'Compartir' diga 'Cualquier persona con el vínculo'."

def guardar_hoja_sincronizada(df_guardar, nombre_hoja, es_admin=False):
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

    if os.path.exists(ARCHIVO_CRED):
        try:
            import gspread
            gc = gspread.service_account(filename=ARCHIVO_CRED)
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
            df_hist_excel = limpiar_escaneos_corruptos(
                convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name="Historial_Tendencias"))
            )
            tiene_futbin_valido = (
                not df_hist_excel.empty
                and "Precio_Futbin" in df_hist_excel.columns
                and (df_hist_excel["Precio_Futbin"].apply(num_seguro) > 0).sum() >= 50
            )
            fecha_excel = str(df_hist_excel["Fecha"].max()) if tiene_futbin_valido and "Fecha" in df_hist_excel.columns else ""
            fecha_sheet = str(datos.get("Historial_Tendencias", pd.DataFrame()).get("Fecha", pd.Series([""])).max()) if datos else ""

            if tiene_futbin_valido and (not datos or fecha_excel >= fecha_sheet):
                cartera_respaldo = datos.get("Mi_Cartera", pd.DataFrame()) if datos else pd.DataFrame()
                alertas_respaldo = datos.get("Mis_Alertas", pd.DataFrame()) if datos else pd.DataFrame()
                datos["Historial_Tendencias"] = df_hist_excel
                for h in [
                    "Recomendacion_IA", "Dashboard_Financiero",
                    "Radar_Fodder", "Arbitraje_FUTGG", "Radar_Evos", "Radar_TOTW", "Radar_Links_Domino"
                ]:
                    try:
                        datos[h] = convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name=h))
                    except Exception:
                        if h not in datos:
                            datos[h] = pd.DataFrame()
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

    # ==========================================
    # BARRA LATERAL COMPLETA Y PERSONALIZABLE
    # ==========================================
    with st.sidebar:
        st.markdown("## 🛠️ Centro de Mando")
        st.caption(estado_conexion)
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

        st.markdown("---")
        pin_ingresado = st.text_input("🔑 PIN de Administrador (Opcional)", type="password", placeholder="Dejar vacío para Modo Invitado")
        es_admin = (pin_ingresado.strip() == PIN_ADMIN_SECRETO)

        if es_admin:
            st.markdown('<div class="mode-badge-admin">👑 Modo Admin Activo (Tu Cartera Real)</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="mode-badge-guest">👤 Modo Invitado (Tu Propia Cartera)</div>', unsafe_allow_html=True)

        columnas_base_cartera = ["Jugador_Exacto", "Estrategia", "Precio_Pagado", "Estado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%", "Fecha_Compra", "Fecha_Venta"]
        columnas_base_alertas = ["Jugador_Buscado", "Precio_Objetivo", "Estado"]

        if es_admin and datos:
            if "override_Mi_Cartera" in st.session_state:
                datos["Mi_Cartera"] = convertir_columnas_numericas(st.session_state["override_Mi_Cartera"])
            elif os.path.exists(ARCHIVO_CARTERA_LOCAL):
                try:
                    datos["Mi_Cartera"] = convertir_columnas_numericas(pd.read_csv(ARCHIVO_CARTERA_LOCAL))
                except Exception:
                    pass
            if "override_Mis_Alertas" in st.session_state:
                datos["Mis_Alertas"] = convertir_columnas_numericas(st.session_state["override_Mis_Alertas"])
            elif os.path.exists(ARCHIVO_ALERTAS_LOCAL):
                try:
                    datos["Mis_Alertas"] = convertir_columnas_numericas(pd.read_csv(ARCHIVO_ALERTAS_LOCAL))
                except Exception:
                    pass
        elif not es_admin and datos:
            if "guest_Mi_Cartera" not in st.session_state:
                st.session_state["guest_Mi_Cartera"] = pd.DataFrame(columns=columnas_base_cartera)
            if "guest_Mis_Alertas" not in st.session_state:
                st.session_state["guest_Mis_Alertas"] = pd.DataFrame(columns=columnas_base_alertas)
            datos["Mi_Cartera"] = convertir_columnas_numericas(st.session_state["guest_Mi_Cartera"])
            datos["Mis_Alertas"] = convertir_columnas_numericas(st.session_state["guest_Mis_Alertas"])

            with st.expander("💾 Guardar / Cargar Mi Cartera (CSV)"):
                csv_guest = datos["Mi_Cartera"].to_csv(index=False).encode("utf-8")
                st.download_button("⬇️ Descargar Mi Cartera (.csv)", data=csv_guest, file_name="mi_cartera_fc27.csv", mime="text/csv", use_container_width=True)
                archivo_subido = st.file_uploader("⬆️ Subir Cartera (.csv)", type=["csv"])
                if archivo_subido is not None:
                    try:
                        st.session_state["guest_Mi_Cartera"] = convertir_columnas_numericas(pd.read_csv(archivo_subido))
                        st.success("¡Cartera cargada!")
                        st.rerun()
                    except Exception as e_csv:
                        st.error(f"Archivo inválido: {e_csv}")

        with st.expander("⚙️ Personalizar Mi Perfil y Reglas", expanded=False):
            pais_usuario = st.selectbox("🌍 Tu País (Hora del Drop)", list(PAISES_DROP.keys()), index=0)
            plataforma_usuario = st.radio("🎮 Plataforma de Mercado", ["🖥️ PC (Origin/EA App)", "🎮 Consola (PS5 / Xbox)"], index=0)
            presupuesto_usuario = st.number_input("💰 Presupuesto Base Referencia", min_value=5000, max_value=50000000, step=10000, value=PRESUPUESTO_DEFAULT)
            roi_puja_pct = st.slider("🔨 % ROI Objetivo en Puja", min_value=3, max_value=25, value=10)
            lazy_pct = st.slider("🎯 % Recargo Lazy Buyer (1h)", min_value=5, max_value=30, value=12)
            stop_loss_pct = st.slider("🛑 % Stop-Loss (Aviso salida)", min_value=-25, max_value=-3, value=-10)
            horas_rotacion = st.slider("⏱️ Horas Tope de Rotación", min_value=12, max_value=168, value=72, step=12)
            api_key_custom = st.text_input("🤖 Gemini API Key Propia (Opcional)", type="password")

        c_t1, c_t2 = st.columns(2)
        with c_t1:
            modo_streamer = st.toggle("👁️ Modo Streamer", value=False)
        with c_t2:
            sonido_sniper = st.toggle("🔔 Sonido Sniper", value=True)

        st.markdown("---")
        st.markdown("### 🧮 Calculadora Rápida (5% Tax)")
        precio_sim = st.number_input("Precio visto en Mercado", min_value=200, max_value=15000000, step=250, value=25000)
        factor_roi = 1.0 + (roi_puja_pct / 100.0)
        factor_lazy = 1.0 + (lazy_pct / 100.0)
        factor_sl = 1.0 + (stop_loss_pct / 100.0)

        puja_sim = redondear_escalon_ea((precio_sim * 0.95) / factor_roi, hacia_arriba=False)
        empate_sim = redondear_escalon_ea(precio_sim / 0.95, hacia_arriba=True)
        lazy_sim = redondear_escalon_ea(precio_sim * factor_lazy, hacia_arriba=True)
        sl_sim = redondear_escalon_ea((precio_sim * factor_sl) / 0.95, hacia_arriba=False)
        neto_lazy_sim = int(round((lazy_sim * 0.95) - precio_sim))

        st.markdown(f"""
        <div class="calc-box">
            🔨 <b>Pujar (+{roi_puja_pct}% ROI):</b> <span style="color:#00e676;font-weight:bold;">{puja_sim:,} mnds</span><br>
            ⚖️ <b>Empate (5% Tax):</b> <span style="color:#84ffff;">{empate_sim:,} mnds</span><br>
            🎯 <b>Listar Lazy (+{lazy_pct}%):</b> <span style="color:#ffea00;font-weight:bold;">{lazy_sim:,} mnds</span><br>
            ✨ <b>Ganancia Neta Lazy:</b> <span style="color:#00e676;">+{neto_lazy_sim:,} mnds</span><br>
            🛑 <b>Stop-Loss ({stop_loss_pct}%):</b> <span style="color:#ff5252;">{sl_sim:,} mnds</span>
        </div>
        """, unsafe_allow_html=True)

        if es_admin:
            st.markdown("---")
            with st.expander("🚀 Control Remoto Bot 24/7 (Admin)"):
                gh_token_in = st.text_input("GitHub Token (workflow):", type="password")
                if st.button("⚡ Escanear Mercado Ahora", use_container_width=True):
                    if gh_token_in.strip():
                        ok_gh, msg_gh = disparar_github_action_remoto(gh_token_in)
                        st.success(msg_gh) if ok_gh else st.error(msg_gh)
                    else:
                        st.warning("Ingresá tu GitHub Token.")

            with st.expander("🔗 Vínculo de Google Sheets (Admin)"):
                link_sb = st.text_input("URL de tu planilla:", value=link_activo)
                if st.button("Guardar Vínculo", use_container_width=True):
                    st.session_state["link_manual"] = link_sb.strip()
                    guardar_link_cache(link_sb.strip())
                    st.cache_data.clear()
                    st.rerun()

    # ==========================================
    # CABECERA PRINCIPAL
    # ==========================================
    df_historial = datos.get("Historial_Tendencias", pd.DataFrame())
    ultima_act = str(df_historial["Fecha"].max()) if not df_historial.empty and "Fecha" in df_historial.columns else "Sin escaneo"
    df_ult_corte = df_historial[df_historial["Fecha"] == ultima_act].copy() if not df_historial.empty and "Fecha" in df_historial.columns else pd.DataFrame()

    usar_consola = "Consola" in plataforma_usuario
    if usar_consola and not df_ult_corte.empty and "Precio_Consola" in df_ult_corte.columns:
        mask_con = df_ult_corte["Precio_Consola"].apply(num_seguro) > 0
        if mask_con.any():
            df_ult_corte.loc[mask_con, "Precio_Actual"] = df_ult_corte.loc[mask_con, "Precio_Consola"]

    if not df_ult_corte.empty and "Precio_Actual" in df_ult_corte.columns:
        df_ult_corte["Puja_Max_10%_ROI"] = df_ult_corte["Precio_Actual"].apply(
            lambda p: redondear_escalon_ea((num_seguro(p) * 0.95) / factor_roi, hacia_arriba=False)
        )
        df_ult_corte["Precio_Lazy_Sugerido"] = df_ult_corte["Precio_Actual"].apply(
            lambda p: redondear_escalon_ea(num_seguro(p) * factor_lazy, hacia_arriba=True)
        )

    tiempo_drop, leyenda_drop, hora_drop_txt = calcular_countdown_pais(pais_usuario)
    etiqueta_plat = "🎮 Consola" if usar_consola else "🖥️ PC"

    st.markdown(f"""
    <div class="main-header">
        <div>
            <h1>⚡ FC 27 TRADING TERMINAL PRO — CLOUD 24/7</h1>
            <p>{estado_conexion} &nbsp;|&nbsp; {etiqueta_plat} &nbsp;|&nbsp; 📦 Padrón: <b>{len(df_ult_corte):,} cartas</b> &nbsp;|&nbsp; 🕒 Corte: <b>{ultima_act}</b></p>
        </div>
        <div class="drop-timer">
            <div class="drop-timer-title">🎯 Drop EA ({hora_drop_txt})</div>
            <div class="drop-timer-val">⏰ Faltan {tiempo_drop}</div>
            <div style="font-size:0.78rem;color:#c9d1d9;">{leyenda_drop}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

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

    # DETECTOR EN VIVO DE ALERTAS SNIPER
    df_alertas_check = datos.get("Mis_Alertas", pd.DataFrame())
    if not df_alertas_check.empty and "Estado" in df_alertas_check.columns and not df_ult_corte.empty:
        activas_sniper = df_alertas_check[df_alertas_check["Estado"].astype(str).str.strip() == "Activa"]
        disparadas_hoy = []
        for _, r_sn in activas_sniper.iterrows():
            j_busq = str(r_sn.get("Jugador_Buscado", "")).strip().lower()
            p_obj = num_seguro(r_sn.get("Precio_Objetivo", 0))
            if j_busq and p_obj > 0:
                m_sn = df_ult_corte[df_ult_corte["Jugador"].astype(str).str.lower().str.contains(j_busq, na=False)]
                if not m_sn.empty:
                    p_actual_sn = num_seguro(m_sn.iloc[0].get("Precio_Actual", 0))
                    if 0 < p_actual_sn <= p_obj:
                        disparadas_hoy.append(f"🎯 **{m_sn.iloc[0]['Jugador']}** cayó a `{p_actual_sn:,} mnds` (Objetivo: `<={p_obj:,} mnds`)")
        if disparadas_hoy:
            st.error("🚨 **¡ALERTA SNIPER CUMPLIDA EN VIVO!**\n\n" + "\n".join([f"- {d}" for d in disparadas_hoy]))
            if sonido_sniper:
                b64_wav = generar_beep_wav_base64()
                st.markdown(f'<audio autoplay><source src="data:audio/wav;base64,{b64_wav}" type="audio/wav"></audio>', unsafe_allow_html=True)

    # ==========================================
    # CÁLCULO Y EDICIÓN DIRECTA DE LIQUIDEZ Y CAPITAL INMOVILIZADO
    # ==========================================
    df_cartera = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
    if not df_cartera.empty and "Estado" in df_cartera.columns:
        for idx_c, r_c in df_cartera.iterrows():
            p_v_row = num_seguro(r_c.get("Precio_Venta", 0), 0)
            p_p_row = num_seguro(r_c.get("Precio_Pagado", 0), 0)
            if str(r_c.get("Estado", "")).strip() == "Vendido" and p_v_row > 0 and num_seguro(r_c.get("Ganancia_Neta", 0), 0) == 0:
                tx_row = int(round(p_v_row * 0.05))
                nt_row = int(round((p_v_row - tx_row) - p_p_row))
                roi_row = round((nt_row / p_p_row) * 100, 2) if p_p_row > 0 else 0.0
                df_cartera.loc[idx_c, "Costo_EA_Tax"] = tx_row
                df_cartera.loc[idx_c, "Ganancia_Neta"] = nt_row
                df_cartera.loc[idx_c, "ROI_%"] = roi_row

        activos = df_cartera[df_cartera["Estado"].astype(str).str.strip() == "Activo"].copy()
        vendidos = df_cartera[df_cartera["Estado"].astype(str).str.strip() == "Vendido"].copy()
    else:
        activos = pd.DataFrame()
        vendidos = pd.DataFrame()

    cap_auto_cartera = int(activos["Precio_Pagado"].apply(num_seguro).sum()) if not activos.empty and "Precio_Pagado" in activos.columns else 0
    saldos_cfg = cargar_saldos_personalizados(es_admin, cap_auto_cartera)

    if saldos_cfg.get("modo_capital") == "Manual (Fijo)":
        cap_inmovilizado = int(saldos_cfg.get("capital_manual", cap_auto_cartera))
    else:
        cap_inmovilizado = cap_auto_cartera

    if saldos_cfg.get("modo_liquidez") == "Manual":
        liquidez_disp = int(saldos_cfg.get("liquidez_manual", max(0, int(presupuesto_usuario) - cap_inmovilizado)))
    else:
        liquidez_disp = max(0, int(presupuesto_usuario) - cap_inmovilizado)

    patrimonio_total_ref = max(1, liquidez_disp + cap_inmovilizado)
    ganancia_realizada = int(vendidos["Ganancia_Neta"].apply(num_seguro).sum()) if not vendidos.empty and "Ganancia_Neta" in vendidos.columns else 0
    aciertos = len(vendidos[vendidos["Ganancia_Neta"].apply(num_seguro) > 0]) if not vendidos.empty and "Ganancia_Neta" in vendidos.columns else 0
    efectividad = round((aciertos / len(vendidos) * 100), 1) if not vendidos.empty else 0.0

    filas_cartera_viva = []
    vnr_total = 0
    for idx_orig, r_a in activos.iterrows():
        nom_a = str(r_a.get("Jugador_Exacto", "Jugador"))
        costo_a = num_seguro(r_a.get("Precio_Pagado", 0), 0)
        est_a = str(r_a.get("Estrategia", "Flipping Meta"))
        if not est_a or est_a == "nan":
            est_a = "Flipping Meta"
        empate_a = redondear_escalon_ea(costo_a / 0.95, hacia_arriba=True)
        sl_a = redondear_escalon_ea((costo_a * factor_sl) / 0.95, hacia_arriba=False)

        horas_hold = 0.0
        try:
            dt_c = datetime.strptime(str(r_a.get("Fecha_Compra", "")).strip(), "%Y-%m-%d %H:%M")
            horas_hold = max(0.0, round((ahora_arg() - dt_c).total_seconds() / 3600.0, 1))
        except Exception:
            pass
        str_hold = f"{int(horas_hold)}h" if horas_hold < 48 else f"{round(horas_hold/24, 1)}d"

        p_mercado = costo_a
        enlace_a = ""
        if not df_ult_corte.empty and "Jugador" in df_ult_corte.columns:
            m = df_ult_corte[df_ult_corte["Jugador"].astype(str).str.lower().str.contains(nom_a.lower(), na=False)]
            if not m.empty:
                fila_m = m.iloc[0]
                p_mercado = max(
                    num_seguro(fila_m.get("Precio_Actual", 0)),
                    num_seguro(fila_m.get("Precio_Futbin", 0)),
                    num_seguro(fila_m.get("Precio_FUTGG", 0))
                )
                enlace_a = str(fila_m.get("Enlace_Optimo", fila_m.get("Enlace_Futbin", "")) or "")

        vnr_carta = int(round(p_mercado * 0.95))
        vnr_total += vnr_carta
        tenencia = vnr_carta - costo_a
        roi_lat = round((tenencia / costo_a) * 100, 1) if costo_a > 0 else 0.0
        lazy_a = redondear_escalon_ea(max(p_mercado * factor_lazy, (costo_a * factor_roi) / 0.95), hacia_arriba=True)

        if roi_lat <= float(stop_loss_pct):
            estado_op = "🛑 STOP-LOSS (Vender)"
        elif tenencia > 0:
            estado_op = "🚨 VENDER (En Ganancia)"
        elif horas_hold >= float(horas_rotacion):
            estado_op = f"🔄 ROTAR STOCK ({str_hold})"
        else:
            estado_op = f"⏳ HOLD / LAZY ({str_hold})"

        filas_cartera_viva.append({
            "_idx": idx_orig,
            "Eliminar": False,
            "Vender": False,
            "Precio_Venta_Real": p_mercado,
            "Jugador": nom_a,
            "Estrategia": est_a,
            "Costo_Pagado": costo_a,
            "Empate_5%": empate_a,
            "Mercado_Hoy": p_mercado,
            "Precio_Lazy_1h": lazy_a,
            "Stop_Loss": sl_a,
            "Resultado_Tenencia": tenencia,
            "ROI_Latente_%": roi_lat,
            "Acción_Sugerida": estado_op,
            "Antigüedad": str_hold,
            "Link": enlace_a
        })

    tenencia_total = vnr_total - cap_inmovilizado
    color_vnr = "kpi-purple" if tenencia_total >= 0 else "kpi-crimson"
    pct_libre = round(liquidez_disp / patrimonio_total_ref * 100)

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card kpi-green"><div class="kpi-title">💵 Liquidez Disponible</div><div class="kpi-value">{fmt_monedas(liquidez_disp, modo_streamer)}</div><div class="kpi-badge">🟢 {pct_libre}% disponible</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-blue"><div class="kpi-title">📦 Capital Inmovilizado</div><div class="kpi-value">{fmt_monedas(cap_inmovilizado, modo_streamer)}</div><div class="kpi-badge">🃏 {len(activos)} cartas en club</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card {color_vnr}"><div class="kpi-title">📈 Stock a VNR (5% Tax)</div><div class="kpi-value">{fmt_monedas(vnr_total, modo_streamer)}</div><div class="kpi-badge">⚖️ Tenencia: {fmt_monedas(tenencia_total, modo_streamer, True)}</div></div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card kpi-gold"><div class="kpi-title">💰 Ganancia Realizada</div><div class="kpi-value">{fmt_monedas(ganancia_realizada, modo_streamer, True)}</div><div class="kpi-badge">📈 {len(vendidos)} ventas cerradas</div></div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""<div class="kpi-card kpi-purple"><div class="kpi-title">🎯 Efectividad Trading</div><div class="kpi-value">{efectividad}%</div><div class="kpi-badge">🏆 {aciertos}/{len(vendidos)} positivos</div></div>""", unsafe_allow_html=True)

    # PANEL EN PÁGINA PRINCIPAL PARA MODIFICAR LIQUIDEZ DISPONIBLE Y CAPITAL INMOVILIZADO
    with st.expander("💳 Modificar Liquidez Disponible y Capital Inmovilizado (Ajuste Directo en Página)", expanded=False):
        st.caption("Modificá tus monedas disponibles en caja o elegí si el Capital Inmovilizado se calcula solo con tus cartas activas o con un valor manual fijo:")
        cs_bal1, cs_bal2, cs_bal3, cs_bal4 = st.columns([1.3, 1.3, 1.3, 1])
        with cs_bal1:
            nueva_liq_man = st.number_input("💵 Liquidez Disponible (Caja Hoy)", min_value=0, max_value=50000000, step=1000, value=int(liquidez_disp))
        with cs_bal2:
            idx_modo_cap = 1 if saldos_cfg.get("modo_capital") == "Manual (Fijo)" else 0
            nuevo_modo_cap = st.selectbox("📦 Cálculo Capital Inmovilizado", ["Auto (Suma de mis cartas)", "Manual (Fijo)"], index=idx_modo_cap)
        with cs_bal3:
            nuevo_cap_man = st.number_input("📦 Capital Inmovilizado Manual", min_value=0, max_value=50000000, step=1000, value=int(cap_inmovilizado), disabled=(nuevo_modo_cap != "Manual (Fijo)"))
        with cs_bal4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 Guardar Saldos", use_container_width=True):
                guardar_saldos_personalizados({
                    "modo_liquidez": "Manual",
                    "liquidez_manual": int(nueva_liq_man),
                    "modo_capital": nuevo_modo_cap,
                    "capital_manual": int(nuevo_cap_man)
                }, es_admin)
                st.success("✅ ¡Liquidez y Capital actualizados!")
                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # PESTAÑAS INTERACTIVAS COMPLETAS
    # ==========================================
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Cartera Editable, P&L & Curva",
        "⚡ Radar Sin Límite, Lupa & Compartir",
        "🎯 Alertas, Optimizador SBC & TOTW",
        "🔮 Oráculo & Reporte IA En Vivo"
    ])

    # ==========================================
    # PESTAÑA 1: GESTIÓN INTERACTIVA DE JUGADORES (AGREGAR, EDITAR, VENDER Y QUITAR)
    # ==========================================
    with tab1:
        st.subheader("🎯 Gestor Interactivo de Mis Jugadores Activos (Editá, Vendé o Quitá desde la tabla)")

        # Barra rápida para agregar jugadores directamente desde la pestaña principal
        with st.expander("➕ Agregar Nuevo Jugador a Mi Cartera Activa", expanded=False):
            ca1, ca2, ca3, ca4 = st.columns([2, 1.2, 1.2, 1])
            with ca1:
                nom_add_tab = st.text_input("Nombre del Jugador", placeholder="Ej: Bellingham, Theo Hernandez...", key="add_nom_t1")
            with ca2:
                costo_add_tab = st.number_input("Costo Pagado (Monedas)", min_value=200, max_value=15000000, step=250, value=15000, key="add_cost_t1")
            with ca3:
                est_add_tab = st.selectbox("Estrategia", ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"], key="add_est_t1")
            with ca4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Sumar al Club", use_container_width=True):
                    if nom_add_tab.strip():
                        nueva_f = pd.DataFrame([{
                            "Jugador_Exacto": nom_add_tab.strip().title(),
                            "Estrategia": est_add_tab,
                            "Precio_Pagado": int(costo_add_tab),
                            "Estado": "Activo",
                            "Precio_Venta": 0,
                            "Costo_EA_Tax": 0,
                            "Ganancia_Neta": 0,
                            "ROI_%": 0.0,
                            "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                            "Fecha_Venta": ""
                        }])
                        df_c_upd = pd.concat([df_cartera, nueva_f], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_upd, "Mi_Cartera", es_admin=es_admin)
                        if saldos_cfg.get("modo_liquidez") == "Manual":
                            saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - int(costo_add_tab))
                            guardar_saldos_personalizados(saldos_cfg, es_admin)
                        st.success(f"✅ ¡{nom_add_tab.title()} agregado a tu cartera!")
                        st.rerun()

        if filas_cartera_viva:
            st.caption("💡 **Cómo usar la tabla:** Podés cambiar el **Jugador**, **Estrategia** o **Costo Pagado** haciendo doble clic en la celda; tildar **💰 Vender** (ajustando el *Precio Venta Real*) para liquidarla, o tildar **🗑️ Quitar** para borrar cartas que ya no tenés. Luego tocá el botón **💾 Guardar Cambios en Mis Jugadores**.")
            df_viva_ed = pd.DataFrame(filas_cartera_viva)
            cols_orden_viva = [
                "_idx", "Eliminar", "Vender", "Precio_Venta_Real", "Jugador", "Estrategia",
                "Costo_Pagado", "Empate_5%", "Mercado_Hoy", "Precio_Lazy_1h",
                "Stop_Loss", "Resultado_Tenencia", "ROI_Latente_%", "Acción_Sugerida", "Antigüedad", "Link"
            ]
            df_editor_activos = st.data_editor(
                df_viva_ed[cols_orden_viva],
                use_container_width=True,
                hide_index=True,
                disabled=["_idx", "Empate_5%", "Mercado_Hoy", "Precio_Lazy_1h", "Stop_Loss", "Resultado_Tenencia", "ROI_Latente_%", "Acción_Sugerida", "Antigüedad", "Link"],
                column_config={
                    "_idx": None,
                    "Eliminar": st.column_config.CheckboxColumn("🗑️ Quitar"),
                    "Vender": st.column_config.CheckboxColumn("💰 Vender"),
                    "Precio_Venta_Real": st.column_config.NumberColumn("💵 Precio Venta", min_value=200, max_value=15000000, step=250, format="%d mnds"),
                    "Jugador": st.column_config.TextColumn("👤 Jugador (Editable)"),
                    "Estrategia": st.column_config.SelectboxColumn("🏷️ Estrategia", options=["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"]),
                    "Costo_Pagado": st.column_config.NumberColumn("💵 Costo Pagado (Editable)", min_value=200, max_value=15000000, step=250, format="%d mnds"),
                    "Empate_5%": st.column_config.NumberColumn("⚖️ Empate (5%)", format="%d mnds"),
                    "Mercado_Hoy": st.column_config.NumberColumn("📊 Mercado Hoy", format="%d mnds"),
                    "Precio_Lazy_1h": st.column_config.NumberColumn(f"🎯 Lazy (+{lazy_pct}%)", format="%d mnds"),
                    "Stop_Loss": st.column_config.NumberColumn(f"🛑 SL ({stop_loss_pct}%)", format="%d mnds"),
                    "Resultado_Tenencia": st.column_config.NumberColumn("✨ Neto Hoy", format="%+d mnds"),
                    "ROI_Latente_%": st.column_config.NumberColumn("📈 ROI %", format="%+.1f%%"),
                    "Link": st.column_config.LinkColumn("🔗 Web", display_text="Abrir Carta")
                },
                key="editor_cartera_activa"
            )

            if st.button("💾 Guardar Cambios en Mis Jugadores (Ediciones / Ventas / Eliminaciones)", use_container_width=True):
                df_cartera_mod = df_cartera.copy()
                indices_borrar = []
                delta_liquidez = 0

                for _, fila_ed in df_editor_activos.iterrows():
                    idx_r = int(fila_ed["_idx"])
                    if idx_r not in df_cartera_mod.index:
                        continue

                    if bool(fila_ed["Eliminar"]):
                        indices_borrar.append(idx_r)
                    elif bool(fila_ed["Vender"]):
                        p_comp_f = float(num_seguro(fila_ed["Costo_Pagado"], 0))
                        p_vent_f = int(num_seguro(fila_ed["Precio_Venta_Real"], 0))
                        tax_f = int(round(p_vent_f * 0.05))
                        neto_f = int(round((p_vent_f - tax_f) - p_comp_f))
                        roi_f = round((neto_f / p_comp_f) * 100, 2) if p_comp_f > 0 else 0.0

                        df_cartera_mod.loc[idx_r, "Jugador_Exacto"] = str(fila_ed["Jugador"]).strip().title()
                        df_cartera_mod.loc[idx_r, "Estrategia"] = str(fila_ed["Estrategia"])
                        df_cartera_mod.loc[idx_r, "Precio_Pagado"] = int(p_comp_f)
                        df_cartera_mod.loc[idx_r, "Estado"] = "Vendido"
                        df_cartera_mod.loc[idx_r, "Precio_Venta"] = p_vent_f
                        df_cartera_mod.loc[idx_r, "Costo_EA_Tax"] = tax_f
                        df_cartera_mod.loc[idx_r, "Ganancia_Neta"] = neto_f
                        df_cartera_mod.loc[idx_r, "ROI_%"] = roi_f
                        df_cartera_mod.loc[idx_r, "Fecha_Venta"] = ahora_arg().strftime("%Y-%m-%d %H:%M")
                        delta_liquidez += (p_vent_f - tax_f)
                    else:
                        df_cartera_mod.loc[idx_r, "Jugador_Exacto"] = str(fila_ed["Jugador"]).strip().title()
                        df_cartera_mod.loc[idx_r, "Estrategia"] = str(fila_ed["Estrategia"])
                        df_cartera_mod.loc[idx_r, "Precio_Pagado"] = int(num_seguro(fila_ed["Costo_Pagado"], 0))

                if indices_borrar:
                    df_cartera_mod = df_cartera_mod.drop(index=indices_borrar).reset_index(drop=True)

                guardar_hoja_sincronizada(df_cartera_mod, "Mi_Cartera", es_admin=es_admin)
                if delta_liquidez != 0 and saldos_cfg.get("modo_liquidez") == "Manual":
                    saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) + delta_liquidez)
                    guardar_saldos_personalizados(saldos_cfg, es_admin)

                st.success("✅ ¡Cartera y saldos actualizados con éxito!")
                st.rerun()
        else:
            st.info("No tenés cartas activas en cartera. Desplegá '➕ Agregar Nuevo Jugador' arriba o seleccioná cartas con 1 clic en la pestaña Radar.")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            if filas_cartera_viva:
                df_bar_viva = pd.DataFrame(filas_cartera_viva)
                fig_bar = px.bar(
                    df_bar_viva, x="Jugador", y=["Costo_Pagado", "Mercado_Hoy", "Precio_Lazy_1h"],
                    barmode="group", title="Valuación por Carta Activa: Costo vs Mercado Hoy vs Target Lazy",
                    color_discrete_sequence=["#2979ff", "#ff5252", "#00e676"]
                )
                fig_bar.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_bar, use_container_width=True)

        with col_g2:
            if not vendidos.empty:
                df_curva = vendidos.copy()
                df_curva["Ganancia_Neta_Num"] = df_curva["Ganancia_Neta"].apply(num_seguro)
                df_curva["Ganancia_Acumulada"] = df_curva["Ganancia_Neta_Num"].cumsum()
                df_curva["Operacion"] = [f"#{i+1} {r.get('Jugador_Exacto', 'Carta')}" for i, (_, r) in enumerate(df_curva.iterrows())]
                fig_curva = px.area(
                    df_curva, x="Operacion", y="Ganancia_Acumulada",
                    markers=True, title="📈 Curva de Contribución Marginal Acumulada (P&L Realizado)",
                    color_discrete_sequence=["#00e676"]
                )
                fig_curva.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_curva, use_container_width=True)

        col_pnl1, col_pnl2 = st.columns([1, 1])
        with col_pnl1:
            st.subheader("📑 Estado de Resultados (Modelo Costeo Variable)")
            ventas_brutas = int(vendidos["Precio_Venta"].apply(num_seguro).sum()) if not vendidos.empty and "Precio_Venta" in vendidos.columns else 0
            if not vendidos.empty and "Precio_Venta" in vendidos.columns and "Precio_Pagado" in vendidos.columns:
                vendidos_liquidados = vendidos[vendidos["Precio_Venta"].apply(num_seguro) > 0]
                costo_adq = int(vendidos_liquidados["Precio_Pagado"].apply(num_seguro).sum())
            else:
                costo_adq = 0
            ea_tax = int(round(ventas_brutas * 0.05))
            ingreso_neto = ventas_brutas - ea_tax
            contrib_marginal = ingreso_neto - costo_adq

            df_estado_resultados = pd.DataFrame([
                {"Rubro Contable": "(+) Ingresos Brutos por Ventas Cerradas", "Monto (Monedas)": ventas_brutas},
                {"Rubro Contable": "(-) Costo Variable Comercial (5% EA Tax)", "Monto (Monedas)": -ea_tax},
                {"Rubro Contable": "(=) Ingreso Neto de Realización", "Monto (Monedas)": ingreso_neto},
                {"Rubro Contable": "(-) Costo de Adquisición de Cartas Vendidas (CMV)", "Monto (Monedas)": -costo_adq},
                {"Rubro Contable": "(=) CONTRIBUCIÓN MARGINAL NETA (GANANCIA REALIZADA)", "Monto (Monedas)": contrib_marginal},
                {"Rubro Contable": "📦 Inventario Activo (Capital Inmovilizado)", "Monto (Monedas)": cap_inmovilizado},
                {"Rubro Contable": "📈 Inventario Activo a VNR Dual FB+GG (Neto 5% Tax)", "Monto (Monedas)": vnr_total},
                {"Rubro Contable": "⚖️ Resultado por Tenencia (No Realizado)", "Monto (Monedas)": tenencia_total},
                {"Rubro Contable": "💵 Liquidez Disponible (Caja Actual)", "Monto (Monedas)": liquidez_disp},
            ])
            styled_er = aplicar_estilo_seguro(df_estado_resultados, cols_neto=["Monto (Monedas)"], cols_accion=["Rubro Contable"])
            st.dataframe(styled_er, use_container_width=True, hide_index=True, column_config={"Monto (Monedas)": st.column_config.NumberColumn("💰 Monedas", format="%+d mnds")})

            excel_bytes = exportar_contabilidad_excel(df_estado_resultados, pd.DataFrame(filas_cartera_viva) if filas_cartera_viva else pd.DataFrame(), vendidos)
            st.download_button("📥 Descargar Contabilidad Completa en Excel (.xlsx)", data=excel_bytes, file_name="Contabilidad_FC27.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        with col_pnl2:
            st.subheader("📜 Libro Mayor de Ventas Cerradas (Editable)")
            if not vendidos.empty:
                df_vend_ed = vendidos.copy()
                df_vend_ed["_idx_v"] = df_vend_ed.index
                df_vend_ed.insert(0, "Quitar_Venta", False)
                edicion_ventas = st.data_editor(
                    df_vend_ed,
                    use_container_width=True,
                    hide_index=True,
                    disabled=["_idx_v", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%"],
                    column_config={
                        "_idx_v": None,
                        "Quitar_Venta": st.column_config.CheckboxColumn("🗑️ Quitar"),
                        "Jugador_Exacto": st.column_config.TextColumn("👤 Jugador"),
                        "Precio_Pagado": st.column_config.NumberColumn("💵 Costo", format="%d mnds"),
                        "Precio_Venta": st.column_config.NumberColumn("💰 Venta Bruta", format="%d mnds"),
                        "Ganancia_Neta": st.column_config.NumberColumn("✨ Neto", format="%+d mnds"),
                    },
                    key="editor_ventas_cerradas"
                )
                if st.button("💾 Actualizar Libro de Ventas Cerradas", use_container_width=True):
                    df_c_vmod = df_cartera.copy()
                    borrar_v = []
                    for _, rv in edicion_ventas.iterrows():
                        iv = int(rv["_idx_v"])
                        if iv not in df_c_vmod.index:
                            continue
                        if bool(rv["Quitar_Venta"]):
                            borrar_v.append(iv)
                        else:
                            pc_v = int(num_seguro(rv["Precio_Pagado"], 0))
                            pv_v = int(num_seguro(rv["Precio_Venta"], 0))
                            tx_v = int(round(pv_v * 0.05))
                            nt_v = (pv_v - tx_v) - pc_v if pv_v > 0 else 0
                            roi_v = round((nt_v / pc_v) * 100, 2) if pc_v > 0 and pv_v > 0 else 0.0
                            df_c_vmod.loc[iv, "Jugador_Exacto"] = str(rv["Jugador_Exacto"]).strip().title()
                            df_c_vmod.loc[iv, "Precio_Pagado"] = pc_v
                            df_c_vmod.loc[iv, "Precio_Venta"] = pv_v
                            df_c_vmod.loc[iv, "Costo_EA_Tax"] = tx_v
                            df_c_vmod.loc[iv, "Ganancia_Neta"] = nt_v
                            df_c_vmod.loc[iv, "ROI_%"] = roi_v
                    if borrar_v:
                        df_c_vmod = df_c_vmod.drop(index=borrar_v).reset_index(drop=True)
                    guardar_hoja_sincronizada(df_c_vmod, "Mi_Cartera", es_admin=es_admin)
                    st.success("✅ ¡Libro de ventas actualizado!")
                    st.rerun()
            else:
                st.info("Todavía no hay ventas cerradas registradas.")

    # ==========================================
    # PESTAÑA 2: RADAR CON LÍMITE DE PRECIO AMPLIADO HASTA 15,000,000 MNDS
    # ==========================================
    with tab2:
        st.subheader("🔍 Lupa Visual Histórica (Evolución de Precio, Piso, Promedio y Zona de Puja)")
        if not df_historial.empty and "Jugador" in df_historial.columns:
            lista_jugadores_unicos = sorted(df_historial["Jugador"].astype(str).unique().tolist())
            jug_seleccionado = st.selectbox("Elegí un jugador para ver su gráfico histórico:", lista_jugadores_unicos)

            if jug_seleccionado:
                df_jug_hist = df_historial[df_historial["Jugador"] == jug_seleccionado].sort_values(by="Fecha")
                precios_serie = df_jug_hist["Precio_Actual"].apply(num_seguro)
                p_actual_j = int(precios_serie.iloc[-1]) if not precios_serie.empty else 0
                p_min_j = int(precios_serie.min()) if not precios_serie.empty else 0
                p_prom_j = int(round(precios_serie.mean())) if not precios_serie.empty else 0
                puja_j = redondear_escalon_ea((p_actual_j * 0.95) / factor_roi, hacia_arriba=False)
                lazy_j = redondear_escalon_ea(p_actual_j * factor_lazy, hacia_arriba=True)

                m_l1, m_l2, m_l3, m_l4, m_l5 = st.columns(5)
                m_l1.metric("💰 Precio Actual", f"{p_actual_j:,} mnds")
                m_l2.metric("💎 Piso Mínimo Hist.", f"{p_min_j:,} mnds")
                m_l3.metric("⚖️ Promedio Hist.", f"{p_prom_j:,} mnds")
                m_l4.metric(f"🔨 Pujar (+{roi_puja_pct}%)", f"{puja_j:,} mnds")
                m_l5.metric(f"🎯 Target Lazy (+{lazy_pct}%)", f"{lazy_j:,} mnds")

                fig_hist = go.Figure()
                fig_hist.add_trace(go.Scatter(x=df_jug_hist["Fecha"], y=precios_serie, mode="lines+markers", name=etiqueta_plat, line=dict(color="#00e5ff", width=3)))
                fig_hist.add_hline(y=p_min_j, line_dash="dash", line_color="#00e676", annotation_text=f"Piso ({p_min_j:,})")
                fig_hist.add_hline(y=puja_j, line_dash="dashdot", line_color="#ffd600", annotation_text=f"Puja Máx ({puja_j:,})")
                fig_hist.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig_hist, use_container_width=True)

        if not df_ult_corte.empty:
            with st.expander("📲 Copiar Resumen de Gangas para WhatsApp / Discord / Telegram"):
                df_top_g = df_ult_corte.sort_values(by="Variacion_Bruta_%").head(5) if "Variacion_Bruta_%" in df_ult_corte.columns else df_ult_corte.head(5)
                lineas_share = [f"⚡ *TOP 5 GANGAS FC27 ({etiqueta_plat} | Corte {ultima_act})* ⚡"]
                for _, rg in df_top_g.iterrows():
                    nom_sh = str(rg.get("Jugador", "")).split(" (")[0]
                    p_sh = num_seguro(rg.get("Precio_Actual", 0))
                    puja_sh = num_seguro(rg.get("Puja_Max_10%_ROI", 0))
                    lazy_sh = num_seguro(rg.get("Precio_Lazy_Sugerido", 0))
                    lineas_share.append(f"• *{nom_sh}*: {p_sh:,} mnds | 🔨 Pujar <= {puja_sh:,} | 🎯 Lazy: {lazy_sh:,}")
                st.code("\n".join(lineas_share), language="markdown")

        st.markdown("---")
        st.subheader("⚡ Padrón Dual Interactivo (Límite de Precio Ampliado hasta 15,000,000 mnds)")
        c_f1, c_f2, c_f3, c_f4 = st.columns([1.8, 1.3, 1.2, 1.2])
        with c_f1:
            busqueda_jug = st.text_input("🔍 Filtrar por Nombre, Media o Tipo", placeholder="Ej: 86, TOTW, Mbappe...")
        with c_f2:
            filtro_accion = st.selectbox("Filtrar por Señal", [
                "Todas", "🟢 COMPRAR", "⚡ Flipping Rápido", "💎 Mínimo Histórico",
                "📦 Out of Packs (TOTW / Promo)", "🔴 VENDER", "⚠️ Evitar Cuchillo"
            ])
        with c_f3:
            piso_precio_filtro = st.number_input("Precio Mínimo (mnds)", min_value=0, max_value=15000000, value=200, step=500)
        with c_f4:
            tope_precio_filtro = st.number_input("Precio Máximo (Hasta 15M)", min_value=500, max_value=15000000, value=5000000, step=25000)

        if not df_ult_corte.empty:
            precios_corte = df_ult_corte["Precio_Actual"].apply(num_seguro)
            df_filtrado = df_ult_corte[(precios_corte >= piso_precio_filtro) & (precios_corte <= tope_precio_filtro)].copy()

            if busqueda_jug.strip():
                df_filtrado = df_filtrado[df_filtrado["Jugador"].astype(str).str.lower().str.contains(busqueda_jug.strip().lower(), na=False)]
            if filtro_accion == "🟢 COMPRAR" and "Accion" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Accion"].astype(str).str.contains("COMPRAR", na=False)]
            elif filtro_accion == "⚡ Flipping Rápido" and "Senal_Trading" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Senal_Trading"].astype(str).str.contains("⚡", na=False)]
            elif filtro_accion == "💎 Mínimo Histórico" and "Estado_Historico" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Estado_Historico"].astype(str).str.contains("💎", na=False)]
            elif filtro_accion == "📦 Out of Packs (TOTW / Promo)":
                df_filtrado = df_filtrado[~df_filtrado["Jugador"].astype(str).str.contains(r"- Normal\)", regex=True, na=False)]
            elif filtro_accion == "🔴 VENDER" and "Accion" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Accion"].astype(str).str.contains("VENDER", na=False)]
            elif filtro_accion == "⚠️ Evitar Cuchillo" and "Senal_Trading" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Senal_Trading"].astype(str).str.contains("⚠️", na=False)]

            cols_mostrar = [c for c in [
                "Jugador", "Precio_Actual", "Fuente_Mejor_Precio", "Precio_Futbin", "Precio_FUTGG",
                "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Break_Even_5%", "Precio_Lazy_Sugerido", "Enlace_Optimo"
            ] if c in df_filtrado.columns]

            df_editor_padron = df_filtrado[cols_mostrar].head(300).copy()
            df_editor_padron.insert(0, "Comprar_1Clic", False)

            edicion_padron = st.data_editor(
                df_editor_padron, use_container_width=True, hide_index=True,
                disabled=[c for c in df_editor_padron.columns if c != "Comprar_1Clic"],
                column_config={
                    "Comprar_1Clic": st.column_config.CheckboxColumn("🛒 Comprar"),
                    "Precio_Actual": st.column_config.NumberColumn("💰 Mejor Precio", format="%d mnds"),
                    "Variacion_Bruta_%": st.column_config.NumberColumn("📊 Var %", format="%+.1f%%"),
                    "Puja_Max_10%_ROI": st.column_config.NumberColumn(f"🔨 Puja (+{roi_puja_pct}%)", format="%d mnds"),
                    "Precio_Lazy_Sugerido": st.column_config.NumberColumn(f"🎯 Lazy (+{lazy_pct}%)", format="%d mnds"),
                    "Enlace_Optimo": st.column_config.LinkColumn("🔗 Link", display_text="Ver Carta")
                }
            )

            seleccionadas_1clic = edicion_padron[edicion_padron["Comprar_1Clic"] == True]
            if not seleccionadas_1clic.empty:
                if st.button(f"📥 Enviar {len(seleccionadas_1clic)} carta(s) a Mi Cartera"):
                    df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                    nuevas_filas = []
                    costo_total_1c = 0
                    for _, r_sel in seleccionadas_1clic.iterrows():
                        nom_limpio = str(r_sel["Jugador"]).split(" (")[0].strip()
                        p_compra_1c = num_seguro(r_sel["Precio_Actual"], 0)
                        costo_total_1c += p_compra_1c
                        est_auto = autodetectar_estrategia(r_sel["Jugador"], p_compra_1c)
                        nuevas_filas.append({
                            "Jugador_Exacto": nom_limpio,
                            "Estrategia": est_auto,
                            "Precio_Pagado": p_compra_1c,
                            "Estado": "Activo",
                            "Precio_Venta": 0,
                            "Costo_EA_Tax": 0,
                            "Ganancia_Neta": 0,
                            "ROI_%": 0.0,
                            "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                            "Fecha_Venta": ""
                        })
                    df_c_act = pd.concat([df_c_act, pd.DataFrame(nuevas_filas)], ignore_index=True)
                    guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin)
                    if saldos_cfg.get("modo_liquidez") == "Manual":
                        saldos_cfg["liquidez_manual"] = max(0, int(saldos_cfg.get("liquidez_manual", liquidez_disp)) - costo_total_1c)
                        guardar_saldos_personalizados(saldos_cfg, es_admin)
                    st.success("¡Cartas agregadas a tu Cartera!")
                    st.rerun()

    # ==========================================
    # PESTAÑA 3: OPTIMIZADOR SBC, SIMULADOR FODDER Y SNIPER
    # ==========================================
    with tab3:
        st.subheader("🧩 Optimizador Matemático de Plantillas SBC (Combo de 11 cartas más barato hoy)")
        med_sbc_obj = st.selectbox("Elegí la Valoración de Plantilla que pide el SBC:", [84, 85, 86, 87, 88], index=2)
        df_combos_sbc = optimizar_plantilla_sbc(med_sbc_obj, df_ult_corte)
        mejor_combo = df_combos_sbc.iloc[0]
        st.success(f"🏆 **Combinación más barata hoy para Plantilla {med_sbc_obj}:** `{mejor_combo['Fórmula SBC (11 Cartas)']}` ➡️ **Costo Total: `{mejor_combo['Costo_Total']:,} mnds`**")
        st.dataframe(df_combos_sbc, use_container_width=True, hide_index=True, column_config={"Costo_Total": st.column_config.NumberColumn("💰 Costo Total (11 Cartas)", format="%d mnds")})

        st.markdown("---")
        st.subheader("🧮 Simulador Personalizado de Inversión en Medias (Fodder SBC)")
        ref_fodder_sim = {84: (1200, 2200), 85: (2800, 4500), 86: (5500, 8500), 87: (9500, 14000), 88: (14500, 21000), 89: (22000, 32000)}
        cs1, cs2, cs3, cs4 = st.columns(4)
        with cs1:
            med_sim = st.selectbox("Media Objetivo", [84, 85, 86, 87, 88, 89], index=2)
        piso_def, techo_def = ref_fodder_sim[med_sim]
        with cs2:
            monto_inv_fod = st.number_input("Monedas a Invertir", min_value=5000, max_value=15000000, value=min(50000, max(5000, int(liquidez_disp))), step=5000)
        with cs3:
            p_compra_fod = st.number_input("Precio Compra por Carta", min_value=700, max_value=100000, value=piso_def, step=100)
        with cs4:
            p_venta_fod = st.number_input("Precio Venta Esperado (SBC)", min_value=1000, max_value=150000, value=techo_def, step=250)

        cant_cartas_fod = int(monto_inv_fod // max(1, p_compra_fod))
        inv_real_fod = cant_cartas_fod * p_compra_fod
        ingreso_neto_fod = int(round(cant_cartas_fod * (p_venta_fod * 0.95)))
        ganancia_neta_fod = ingreso_neto_fod - inv_real_fod
        roi_fod = round((ganancia_neta_fod / max(1, inv_real_fod)) * 100, 1)

        st.info(f"📦 **Resultado (Media {med_sim}):** Comprás **{cant_cartas_fod} cartas** (`{inv_real_fod:,} mnds`) ➡️ Al vender en `{p_venta_fod:,} mnds` recibís limpio Post-Tax `{ingreso_neto_fod:,} mnds` (**Ganancia Neta: `{ganancia_neta_fod:+,} mnds` | ROI: `{roi_fod:+.1f}%`**).")

        st.markdown("---")
        st.subheader("🎯 Gestor Visual de Alertas Sniper")
        col_al1, col_al2 = st.columns([1, 2])
        with col_al1:
            with st.form("form_alerta_web", clear_on_submit=True):
                jug_al = st.text_input("Jugador a vigilar", placeholder="Ej: Theo Hernandez")
                obj_al = st.number_input("Precio Objetivo (<=)", min_value=200, max_value=15000000, step=250, value=15000)
                if st.form_submit_button("🎯 Activar Sniper", use_container_width=True) and jug_al.strip():
                    df_al_act = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
                    nueva_al = pd.DataFrame([{"Jugador_Buscado": jug_al.strip().title(), "Precio_Objetivo": int(obj_al), "Estado": "Activa"}])
                    df_al_act = pd.concat([df_al_act, nueva_al], ignore_index=True)
                    guardar_hoja_sincronizada(df_al_act, "Mis_Alertas", es_admin=es_admin)
                    st.rerun()

        with col_al2:
            df_alertas_w = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
            activas_al = df_alertas_w[df_alertas_w["Estado"].astype(str).str.strip() == "Activa"].copy() if not df_alertas_w.empty and "Estado" in df_alertas_w.columns else pd.DataFrame()
            if not activas_al.empty:
                st.dataframe(activas_al, use_container_width=True, hide_index=True)
                if st.button("🗑️ Limpiar Alertas Activas"):
                    df_alertas_w.loc[df_alertas_w["Estado"].astype(str).str.strip() == "Activa", "Estado"] = "Cancelada"
                    guardar_hoja_sincronizada(df_alertas_w, "Mis_Alertas", es_admin=es_admin)
                    st.rerun()
            else:
                st.info("No tenés alertas Sniper activas en este momento.")

    # ==========================================
    # PESTAÑA 4: CENTRO DE INTELIGENCIA ARTIFICIAL EN VIVO
    # ==========================================
    with tab4:
        st.subheader("🧠 Centro de Inteligencia Artificial & Oráculo En Vivo (Bajo Demanda)")
        df_ia_base = datos.get("Recomendacion_IA", pd.DataFrame())
        rep_guardado = str(df_ia_base["Reporte"].dropna().iloc[-1]) if not df_ia_base.empty and "Reporte" in df_ia_base.columns and not df_ia_base["Reporte"].dropna().empty else "Tocá el botón '🤖 Generar Nuevo Reporte IA Ahora' para analizar el mercado en vivo."
        ora_guardado = str(df_ia_base["Oraculo"].dropna().iloc[-1]) if not df_ia_base.empty and "Oraculo" in df_ia_base.columns and not df_ia_base["Oraculo"].dropna().empty else "Tocá el botón '🔮 Tirar Nuevo Oráculo' para predecir el contenido del día."

        if "reporte_ia_vivo" not in st.session_state:
            st.session_state["reporte_ia_vivo"] = rep_guardado
        if "oraculo_ia_vivo" not in st.session_state:
            st.session_state["oraculo_ia_vivo"] = ora_guardado

        col_ctrl1, col_ctrl2, col_ctrl3 = st.columns([2, 1, 1])
        with col_ctrl1:
            enfoque_ia = st.selectbox("🎯 Elegí el enfoque del análisis:", [
                "🎯 Estrategia Global Balanceada (Flipping + SBC)",
                "⚡ Flipping Rápido Intradía (Rebotes y Liquidez)",
                "🧱 Acumulación de Medias / Fodder SBC y TOTW",
                "🏆 Compras Meta para Ciclo FUT Champions",
                "🛡️ Auditoría Profunda de Mi Cartera Activa"
            ])
        with col_ctrl2:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_nuevo_rep = st.button("🤖 Generar Nuevo Reporte IA Ahora", use_container_width=True)
        with col_ctrl3:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_nuevo_ora = st.button(f"🔮 Tirar Nuevo Oráculo ({hora_drop_txt})", use_container_width=True)

        if btn_nuevo_rep:
            with st.spinner("🧠 Analizando el padrón y tu cartera con Gemini AI en tiempo real..."):
                nuevo_rep_txt = generar_reporte_estrategico_vivo(df_ult_corte, filas_cartera_viva, int(liquidez_disp), roi_puja_pct, lazy_pct, etiqueta_plat, enfoque_ia, api_key_custom)
                st.session_state["reporte_ia_vivo"] = nuevo_rep_txt
                if es_admin:
                    guardar_hoja_sincronizada(pd.DataFrame([{"Fecha": ahora_arg().strftime("%Y-%m-%d %H:%M"), "Reporte": nuevo_rep_txt, "Oraculo": st.session_state["oraculo_ia_vivo"]}]), "Recomendacion_IA", es_admin=True)
                st.success("✅ ¡Nuevo Reporte Estratégico generado en vivo!")

        if btn_nuevo_ora:
            with st.spinner("🔮 Consultando al Oráculo de Contenido en vivo..."):
                nuevo_ora_txt = generar_oraculo_contenido_vivo(df_ult_corte, hora_drop_txt, etiqueta_plat, api_key_custom)
                st.session_state["oraculo_ia_vivo"] = nuevo_ora_txt
                if es_admin:
                    guardar_hoja_sincronizada(pd.DataFrame([{"Fecha": ahora_arg().strftime("%Y-%m-%d %H:%M"), "Reporte": st.session_state["reporte_ia_vivo"], "Oraculo": nuevo_ora_txt}]), "Recomendacion_IA", es_admin=True)
                st.success("✅ ¡Nuevo Oráculo generado en vivo!")

        col_ia1, col_ia2 = st.columns(2)
        with col_ia1:
            st.subheader("🤖 Reporte Estratégico Dual (En Vivo)")
            st.markdown(f'<div class="ia-card-report">{st.session_state["reporte_ia_vivo"]}</div>', unsafe_allow_html=True)
        with col_ia2:
            st.subheader(f"🔮 Oráculo de Contenido ({hora_drop_txt})")
            st.markdown(f'<div class="ia-card-oracle">{st.session_state["oraculo_ia_vivo"]}</div>', unsafe_allow_html=True)

    if auto_ref:
        time.sleep(60)
        st.cache_data.clear()
        st.rerun()

except Exception as err_global:
    st.error(f"🚨 Error detectado: {err_global}")
    st.code(traceback.format_exc())
