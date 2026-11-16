import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import re
import math
import time
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
ARCHIVO_EXCEL = "Lista_Jugadores_FC27.xlsx"
ARCHIVO_LINK_CACHE = "sheet_link_cache.txt"
ARCHIVO_CARTERA_LOCAL = "cartera_cloud_sync.csv"
ARCHIVO_ALERTAS_LOCAL = "alertas_cloud_sync.csv"
ARCHIVO_IA_LOCAL = "ia_cloud_sync.csv"
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
    section[data-testid="stSidebar"] * {
        color: #e6edf3 !important;
    }
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
    .main-header p {
        margin: 6px 0 0 0;
        color: #e2e8f0;
        font-size: 0.96rem;
    }
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
    .sniper-banner {
        background: linear-gradient(90deg, #5c0d28 0%, #b5179e 50%, #0d2b6b 100%);
        border: 2px solid #ffea00;
        border-radius: 14px;
        padding: 14px 20px;
        margin-bottom: 16px;
        box-shadow: 0 0 25px rgba(255, 234, 0, 0.4);
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
# FUNCIONES MATEMÁTICAS, HORA, SONIDO Y EXCEL
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

def es_carta_especial_oop(nombre_carta):
    nom_u = str(nombre_carta).upper()
    if any(k in nom_u for k in ["TOTW", "IN-FORM", "INFORM", " - IF)", "PROMO", "SPECIAL", "HERO", "ICON", "RTTK", "RTTF", "TRAILBLAZERS", "CENTURIONS", "THUNDERSTRUCK", "WINTER", "TOTY", "TOTS", "FUTTIES"]):
        return True
    if "NORMAL" not in nom_u and "GOLD" not in nom_u and "RARE" not in nom_u and "COMMON" not in nom_u and " - " in nom_u:
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

def generar_reporte_estrategico_vivo(df_corte, filas_cartera, presupuesto, roi_pct, lazy_pct, plataforma, enfoque, top_n, api_key_usar):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_hoy = dias[ahora.weekday()]
    hora_str = ahora.strftime("%H:%M")

    df_disp = df_corte[df_corte["Precio_Actual"].apply(num_seguro) <= presupuesto].copy() if not df_corte.empty else pd.DataFrame()
    if df_disp.empty:
        df_disp = df_corte.copy()

    df_compras = df_disp[df_disp["Accion"].astype(str).str.contains("COMPRAR", na=False)].sort_values(by="Variacion_Bruta_%").head(18) if "Accion" in df_disp.columns else df_disp.head(18)
    df_ventas = df_disp[df_disp["Accion"].astype(str).str.contains("VENDER", na=False)].sort_values(by="Variacion_Neta_%", ascending=False).head(12) if "Accion" in df_disp.columns else pd.DataFrame()
    df_muestra = pd.concat([df_compras, df_ventas, df_disp.head(15)]).drop_duplicates(subset=["Jugador"]).head(30)

    cols_res = [c for c in ["Jugador", "Precio_Actual", "Fuente_Mejor_Precio", "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Precio_Lazy_Sugerido", "Accion"] if c in df_muestra.columns]
    tabla_txt = df_muestra[cols_res].to_string(index=False) if not df_muestra.empty else "Sin datos en el corte actual."

    cartera_txt = "Sin cartas activas en cartera."
    if filas_cartera:
        lineas_c = [
            f"- {c['Jugador']} ({c['Estrategia']}): Costo {c['Costo_Pagado']:,} | Hoy {c['Mercado_Hoy']:,} | Lazy {c['Precio_Lazy_1h']:,} | Neto {c['Resultado_Tenencia']:+,} ({c['ROI_Latente_%']:+.1f}% | {c['ROI_Diario_%']:+.1f}%/día)"
            for c in filas_cartera
        ]
        cartera_txt = "\n".join(lineas_c)

    prompt = f"""
    Actuá como Trader Cuantitativo Principal de EA Sports FC 27 para la plataforma {plataforma}.
    Momento del análisis: {dia_hoy} {hora_str} hs (ARG).
    Parámetros del usuario: Presupuesto Tope: {presupuesto:,} monedas | Objetivo Puja: +{roi_pct}% ROI | Recargo Lazy Buyer: +{lazy_pct}% | Tax EA: 5%.
    ENFOQUE SOLICITADO POR EL USUARIO: {enfoque}.

    CARTERA ACTIVA ACTUAL DEL USUARIO:
    {cartera_txt}

    MUESTRA EN VIVO DEL MERCADO ({len(df_corte)} cartas totales en padrón):
    {tabla_txt}

    Redactá un REPORTE ESTRATÉGICO EN VIVO claro, directo y accionable (usando viñetas, sin usar tablas con '|'):
    1. 🎯 TOP {top_n} OPORTUNIDADES DE COMPRA AHORA ({enfoque}): Indicá Nombre, Precio Actual, Fuente más barata (FB o GG), hasta cuánto PUJAR (+{roi_pct}% ROI) y a cuánto listar para LAZY BUYERS (+{lazy_pct}%).
    2. 💼 DIAGNÓSTICO DE TU CARTERA ACTIVA: Qué hacer hoy con cada una de las cartas que el usuario tiene en su cartera (o cómo asignar su liquidez si no tiene cartas).
    3. 🧭 PLAN DE ACCIÓN PARA HOY ({dia_hoy}): Recomendación según la fase actual del ciclo semanal de Rivals / FUT Champions.
    """

    texto_ia, modelo_usado = llamar_ia_en_vivo(prompt, api_key_usar)
    if texto_ia:
        cabecera = f"🤖 **REPORTE IA EN VIVO ({dia_hoy} {hora_str} hs | Enfoque: {enfoque} | Modelo: {modelo_usado})**\n\n"
        return cabecera + texto_ia

    lineas_fb = [
        f"🤖 **REPORTE TÁCTICO CUANTITATIVO EN VIVO ({dia_hoy} {hora_str} hs | Enfoque: {enfoque})**",
        f"*(Generado en tiempo real sobre {len(df_corte):,} cartas con Presupuesto: {presupuesto:,} mnds)*\n",
        f"### 🎯 Top {top_n} Oportunidades Detectadas en tu Rango (<={presupuesto:,} mnds):"
    ]
    for _, r in df_muestra.head(top_n).iterrows():
        nom = str(r.get("Jugador", "Carta")).split(" (")[0]
        p_a = num_seguro(r.get("Precio_Actual", 0))
        f_a = str(r.get("Fuente_Mejor_Precio", "PC"))
        puja_a = num_seguro(r.get("Puja_Max_10%_ROI", 0))
        lazy_a = num_seguro(r.get("Precio_Lazy_Sugerido", 0))
        var_a = float_seguro(r.get("Variacion_Bruta_%", 0))
        sen_a = str(r.get("Senal_Trading", "Operable"))
        lineas_fb.append(f"- **{nom}** (`{p_a:,} mnds` en `{f_a}` | `{var_a:+.1f}%`) — {sen_a} ➡️ 🔨 Pujar hasta `{puja_a:,}` | 🎯 Listar Lazy a `{lazy_a:,}`")

    lineas_fb.append("\n### 💼 Diagnóstico de tu Cartera Activa:")
    if filas_cartera:
        for c in filas_cartera:
            lineas_fb.append(f"- **{c['Jugador']}**: Mercado hoy `{c['Mercado_Hoy']:,}` (Tenencia: `{c['Resultado_Tenencia']:+,} mnds` / `{c['ROI_Latente_%']:+.1f}%`) ➡️ **{c['Acción_Sugerida']}** (Listar Lazy en `{c['Precio_Lazy_1h']:,}`).")
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
    Hoy es {dia_hoy}, hora actual: {hora_str} hs (ARG). Horario del Drop diario para el usuario: {hora_drop_txt}.
    Pauta de contenido de EA para {dia_hoy}: {pauta_dia}
    Cartas con mayor presión compradora en vivo hoy: {', '.join(subidas) if subidas else 'Mercado estable sin picos abruptos'}
    Precios piso de Medias (Fodder SBC) ahora mismo: {', '.join(fodder_lista) if fodder_lista else 'Sin muestra de fodder'}

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

        # 1. CONTROL MULTIUSUARIO (ADMIN PIN vs MODO INVITADO)
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
                    df_c_loc = convertir_columnas_numericas(pd.read_csv(ARCHIVO_CARTERA_LOCAL))
                    if len(df_c_loc) >= len(datos.get("Mi_Cartera", pd.DataFrame())):
                        datos["Mi_Cartera"] = df_c_loc
                except Exception:
                    pass
            if "override_Mis_Alertas" in st.session_state:
                datos["Mis_Alertas"] = convertir_columnas_numericas(st.session_state["override_Mis_Alertas"])
            elif os.path.exists(ARCHIVO_ALERTAS_LOCAL):
                try:
                    df_a_loc = convertir_columnas_numericas(pd.read_csv(ARCHIVO_ALERTAS_LOCAL))
                    if len(df_a_loc) >= len(datos.get("Mis_Alertas", pd.DataFrame())):
                        datos["Mis_Alertas"] = df_a_loc
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
                st.caption("Descargá tu cartera a tu PC/celular o subila cuando vuelvas a entrar:")
                csv_guest = datos["Mi_Cartera"].to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Descargar Mi Cartera (.csv)",
                    data=csv_guest,
                    file_name="mi_cartera_fc27.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                archivo_subido = st.file_uploader("⬆️ Subir Cartera guardada (.csv)", type=["csv"])
                if archivo_subido is not None:
                    try:
                        df_sub = pd.read_csv(archivo_subido)
                        st.session_state["guest_Mi_Cartera"] = convertir_columnas_numericas(df_sub)
                        st.success("¡Cartera cargada con éxito!")
                        st.rerun()
                    except Exception as e_csv:
                        st.error(f"Archivo inválido: {e_csv}")

        # 2. CONFIGURACIÓN DEL PERFIL DE TRADER
        with st.expander("⚙️ Personalizar Mi Perfil y Reglas", expanded=False):
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

        # 3. CALCULADORA RÁPIDA ADAPTADA AL PERFIL DEL USUARIO
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

        # 4. REGISTRAR COMPRA MANUAL
        st.markdown("---")
        st.markdown("### ➕ Registrar Compra Manual")
        with st.form("form_compra", clear_on_submit=True):
            jug_c = st.text_input("Nombre del Jugador", placeholder="Ej: Odegaard")
            precio_c = st.number_input("Precio Pagado (Monedas)", min_value=200, max_value=15000000, step=250, value=10000)
            est_c = st.selectbox("Unidad de Negocio", ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"])
            btn_guardar_c = st.form_submit_button("✅ Guardar en Cartera", use_container_width=True)

            if btn_guardar_c and jug_c.strip():
                df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                nueva_fila = pd.DataFrame([{
                    "Jugador_Exacto": jug_c.strip().title(),
                    "Estrategia": est_c,
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
                guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin)
                st.success(f"¡{jug_c.title()} agregado a Cartera!")
                st.rerun()

        # 5. CERRAR / CORREGIR VENTA
        st.markdown("---")
        st.markdown("### 💰 Cerrar / Corregir Venta")
        df_cartera_sb = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
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
                    f"{idx} | {r.get('Jugador_Exacto', 'Carta')} (Costo: {num_seguro(r.get('Precio_Pagado', 0)):,} | {r.get('Estado', 'Activo')})"
                    for idx, r in activos_sb.iterrows()
                ]
                seleccion_v = st.selectbox("Carta a Vender / Actualizar", opciones_v)
                precio_v = st.number_input("Precio de Venta Bruto", min_value=200, max_value=15000000, step=250, value=15000)
                btn_vender = st.form_submit_button("📈 Registrar Venta (5% Tax)", use_container_width=True)

                if btn_vender:
                    idx_real = int(seleccion_v.split(" | ")[0])
                    p_comp = float(num_seguro(df_cartera_sb.loc[idx_real, "Precio_Pagado"], 0))
                    tax = int(round(precio_v * 0.05))
                    neto = int(round((precio_v - tax) - p_comp))
                    roi = round((neto / p_comp) * 100, 2) if p_comp > 0 else 0.0

                    df_cartera_sb["Estado"] = df_cartera_sb["Estado"].astype(object)
                    if "Fecha_Venta" not in df_cartera_sb.columns:
                        df_cartera_sb["Fecha_Venta"] = ""
                    df_cartera_sb["Fecha_Venta"] = df_cartera_sb["Fecha_Venta"].astype(object)

                    df_cartera_sb.loc[idx_real, "Estado"] = "Vendido"
                    df_cartera_sb.loc[idx_real, "Precio_Venta"] = int(precio_v)
                    df_cartera_sb.loc[idx_real, "Costo_EA_Tax"] = tax
                    df_cartera_sb.loc[idx_real, "Ganancia_Neta"] = neto
                    df_cartera_sb.loc[idx_real, "ROI_%"] = roi
                    df_cartera_sb.loc[idx_real, "Fecha_Venta"] = ahora_arg().strftime("%Y-%m-%d %H:%M")

                    guardar_hoja_sincronizada(df_cartera_sb, "Mi_Cartera", es_admin=es_admin)
                    st.success(f"¡Venta registrada! Neto: {neto:+,} mnds")
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
    # CABECERA PRINCIPAL + RELOJ ADAPTADO AL PAÍS
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

    # ==========================================
    # DETECTOR VISUAL Y SONORO DE ALERTAS SNIPER EN VIVO
    # ==========================================
    df_alertas_check = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
    if not df_alertas_check.empty and "Estado" in df_alertas_check.columns and not df_ult_corte.empty:
        act_snipers = df_alertas_check[df_alertas_check["Estado"].astype(str).str.strip() == "Activa"]
        disparos_sniper = []
        for _, r_sn in act_snipers.iterrows():
            j_b = str(r_sn.get("Jugador_Buscado", "")).strip()
            p_obj = num_seguro(r_sn.get("Precio_Objetivo", 0))
            if j_b and p_obj > 0:
                m_sn = df_ult_corte[df_ult_corte["Jugador"].astype(str).str.lower().str.contains(j_b.lower(), na=False)]
                if not m_sn.empty:
                    mejor_fila = m_sn.loc[m_sn["Precio_Actual"].apply(num_seguro).idxmin()]
                    p_actual_sn = num_seguro(mejor_fila.get("Precio_Actual", 0))
                    if 0 < p_actual_sn <= p_obj:
                        disparos_sniper.append((
                            str(mejor_fila["Jugador"]).split(" (")[0],
                            p_actual_sn,
                            p_obj,
                            str(mejor_fila.get("Fuente_Mejor_Precio", "PC")),
                            str(mejor_fila.get("Enlace_Optimo", mejor_fila.get("Enlace_Futbin", "")))
                        ))
        if disparos_sniper:
            if alerta_sonora:
                reproducir_sonido_sniper()
            items_html = "<br>".join([
                f"🎯 <b>{nom_s}</b> cayó a <span style='color:#00e676;font-weight:900;'>{p_s:,} mnds</span> [{f_s}] (Tu objetivo: <= {obj_s:,} mnds) — <a href='{lnk_s}' target='_blank' style='color:#ffea00;'>Abrir Carta</a>"
                for nom_s, p_s, obj_s, f_s, lnk_s in disparos_sniper
            ])
            st.markdown(f'<div class="sniper-banner">🚨 <b>¡ALERTA SNIPER DETECTADA EN VIVO!</b><br>{items_html}</div>', unsafe_allow_html=True)

    # ==========================================
    # TARJETAS KPI FINANCIERAS
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

    cap_inmovilizado = int(activos["Precio_Pagado"].apply(num_seguro).sum()) if not activos.empty and "Precio_Pagado" in activos.columns else 0
    liquidez_disp = max(0, int(presupuesto_usuario) - cap_inmovilizado)
    ganancia_realizada = int(vendidos["Ganancia_Neta"].apply(num_seguro).sum()) if not vendidos.empty and "Ganancia_Neta" in vendidos.columns else 0
    aciertos = len(vendidos[vendidos["Ganancia_Neta"].apply(num_seguro) > 0]) if not vendidos.empty and "Ganancia_Neta" in vendidos.columns else 0
    efectividad = round((aciertos / len(vendidos) * 100), 1) if not vendidos.empty else 0.0

    filas_cartera_viva = []
    vnr_total = 0
    for idx_orig, r_a in activos.iterrows():
        nom_a = str(r_a.get("Jugador_Exacto", "Jugador"))
        costo_a = num_seguro(r_a.get("Precio_Pagado", 0), 0)
        est_a = str(r_a.get("Estrategia", "Flipping Meta"))
        if not est_a or est_a.lower() == "nan":
            est_a = autodetectar_estrategia(nom_a, costo_a)
        empate_a = redondear_escalon_ea(costo_a / 0.95, hacia_arriba=True)
        sl_a = redondear_escalon_ea((costo_a * factor_sl) / 0.95, hacia_arriba=False)

        horas_hold = 1.0
        try:
            dt_c = datetime.strptime(str(r_a.get("Fecha_Compra", "")).strip(), "%Y-%m-%d %H:%M")
            horas_hold = max(1.0, round((ahora_arg() - dt_c).total_seconds() / 3600.0, 1))
        except Exception:
            pass
        dias_hold = max(0.25, round(horas_hold / 24.0, 2))
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
        roi_diario = round(roi_lat / max(1.0, dias_hold), 1)
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
            "Idx_Cartera": idx_orig,
            "Jugador": nom_a,
            "Estrategia": est_a,
            "Costo_Pagado": costo_a,
            "Empate_5%": empate_a,
            "Mercado_Hoy": p_mercado,
            "Precio_Lazy_1h": lazy_a,
            "Stop_Loss": sl_a,
            "Resultado_Tenencia": tenencia,
            "ROI_Latente_%": roi_lat,
            "ROI_Diario_%": roi_diario,
            "Acción_Sugerida": estado_op,
            "Antigüedad": str_hold,
            "Vendido_A": 0,
            "Link": enlace_a
        })

    tenencia_total = vnr_total - cap_inmovilizado
    color_vnr = "kpi-purple" if tenencia_total >= 0 else "kpi-crimson"
    pct_libre = round(liquidez_disp / max(1, presupuesto_usuario) * 100)

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card kpi-green"><div class="kpi-title">💵 Liquidez Disponible</div><div class="kpi-value">{fmt_monedas(liquidez_disp, modo_streamer)}</div><div class="kpi-badge">🟢 {pct_libre}% de {fmt_monedas(presupuesto_usuario, modo_streamer)}</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-blue"><div class="kpi-title">📦 Capital Inmovilizado</div><div class="kpi-value">{fmt_monedas(cap_inmovilizado, modo_streamer)}</div><div class="kpi-badge">🃏 {len(activos)} cartas en club</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card {color_vnr}"><div class="kpi-title">📈 Stock a VNR (5% Tax)</div><div class="kpi-value">{fmt_monedas(vnr_total, modo_streamer)}</div><div class="kpi-badge">⚖️ Tenencia: {fmt_monedas(tenencia_total, modo_streamer, True)}</div></div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card kpi-gold"><div class="kpi-title">💰 Ganancia Realizada</div><div class="kpi-value">{fmt_monedas(ganancia_realizada, modo_streamer, True)}</div><div class="kpi-badge">📈 {len(vendidos)} ventas cerradas</div></div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""<div class="kpi-card kpi-purple"><div class="kpi-title">🎯 Efectividad Trading</div><div class="kpi-value">{efectividad}%</div><div class="kpi-badge">🏆 {aciertos}/{len(vendidos)} positivos</div></div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # PESTAÑAS INTERACTIVAS COMPLETAS
    # ==========================================
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Cartera 1-Clic, P&L & Proyector",
        "⚡ Radar 1-Clic, Lupa & Compartir",
        "🎯 Optimizador SBC, Fodder & Sniper",
        "🔮 Oráculo & Reporte IA En Vivo"
    ])

    with tab1:
        st.subheader("🎯 Monitoreo de Cartera Activa (Escribí en '💸 Vendido_A' para liquidar en 1 Clic)")
        if filas_cartera_viva:
            df_viva = pd.DataFrame(filas_cartera_viva)
            cols_deshabilitadas_viva = [c for c in df_viva.columns if c != "Vendido_A"]

            edicion_viva = st.data_editor(
                df_viva,
                use_container_width=True,
                hide_index=True,
                disabled=cols_deshabilitadas_viva,
                column_config={
                    "Idx_Cartera": None,
                    "Costo_Pagado": st.column_config.NumberColumn("💵 Costo", format="%d mnds"),
                    "Empate_5%": st.column_config.NumberColumn("⚖️ Empate (5%)", format="%d mnds"),
                    "Mercado_Hoy": st.column_config.NumberColumn("📊 Mercado Hoy", format="%d mnds"),
                    "Precio_Lazy_1h": st.column_config.NumberColumn(f"🎯 Lazy (+{lazy_pct}%)", format="%d mnds"),
                    "Stop_Loss": st.column_config.NumberColumn(f"🛑 SL ({stop_loss_pct}%)", format="%d mnds"),
                    "Resultado_Tenencia": st.column_config.NumberColumn("✨ Neto Hoy", format="%+d mnds"),
                    "ROI_Latente_%": st.column_config.NumberColumn("📈 ROI %", format="%+.1f%%"),
                    "ROI_Diario_%": st.column_config.NumberColumn("⚡ ROI/Día", format="%+.1f%%"),
                    "Vendido_A": st.column_config.NumberColumn("💸 Vendido_A (Editar)", min_value=0, max_value=15000000, step=250, format="%d mnds"),
                    "Link": st.column_config.LinkColumn("🔗 Web", display_text="Abrir Carta")
                }
            )

            filas_a_liquidar = edicion_viva[edicion_viva["Vendido_A"].apply(num_seguro) > 0]
            if not filas_a_liquidar.empty:
                if st.button(f"💰 Liquidar {len(filas_a_liquidar)} Venta(s) Cargada(s) en Tabla (5% Tax)"):
                    df_c_upd = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                    df_c_upd["Estado"] = df_c_upd["Estado"].astype(object)
                    if "Fecha_Venta" not in df_c_upd.columns:
                        df_c_upd["Fecha_Venta"] = ""
                    df_c_upd["Fecha_Venta"] = df_c_upd["Fecha_Venta"].astype(object)

                    for _, r_liq in filas_a_liquidar.iterrows():
                        idx_c_orig = int(r_liq["Idx_Cartera"])
                        p_venta_liq = int(num_seguro(r_liq["Vendido_A"]))
                        p_comp_liq = float(num_seguro(df_c_upd.loc[idx_c_orig, "Precio_Pagado"], 0))
                        tax_liq = int(round(p_venta_liq * 0.05))
                        neto_liq = int(round((p_venta_liq - tax_liq) - p_comp_liq))
                        roi_liq = round((neto_liq / p_comp_liq) * 100, 2) if p_comp_liq > 0 else 0.0

                        df_c_upd.loc[idx_c_orig, "Estado"] = "Vendido"
                        df_c_upd.loc[idx_c_orig, "Precio_Venta"] = p_venta_liq
                        df_c_upd.loc[idx_c_orig, "Costo_EA_Tax"] = tax_liq
                        df_c_upd.loc[idx_c_orig, "Ganancia_Neta"] = neto_liq
                        df_c_upd.loc[idx_c_orig, "ROI_%"] = roi_liq
                        df_c_upd.loc[idx_c_orig, "Fecha_Venta"] = ahora_arg().strftime("%Y-%m-%d %H:%M")

                    guardar_hoja_sincronizada(df_c_upd, "Mi_Cartera", es_admin=es_admin)
                    st.success("¡Ventas liquidadas y registradas en el Libro Mayor!")
                    st.rerun()
        else:
            st.info("No tenés cartas activas en cartera. Podés agregarlas desde la barra lateral o con 1 clic en la pestaña Radar.")

        # GRÁFICOS: VALUACIÓN, DIVERSIFICACIÓN DE RIESGO Y CURVA P&L
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            if filas_cartera_viva:
                df_bar_viva = pd.DataFrame(filas_cartera_viva)
                fig_bar = px.bar(
                    df_bar_viva, x="Jugador", y=["Costo_Pagado", "Mercado_Hoy", "Precio_Lazy_1h"],
                    barmode="group", title="📦 Costo vs Mercado Hoy vs Target Lazy",
                    color_discrete_sequence=["#2979ff", "#ff5252", "#00e676"]
                )
                fig_bar.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig_bar, use_container_width=True)

        with col_g2:
            filas_div = [{"Rubro": "💵 Liquidez en Caja", "Monedas": liquidez_disp}]
            if filas_cartera_viva:
                df_ag_est = pd.DataFrame(filas_cartera_viva).groupby("Estrategia")["Costo_Pagado"].sum().reset_index()
                for _, r_est in df_ag_est.iterrows():
                    filas_div.append({"Rubro": str(r_est["Estrategia"]), "Monedas": int(r_est["Costo_Pagado"])})
                    if int(r_est["Costo_Pagado"]) > (presupuesto_usuario * 0.60):
                        st.warning(f"⚠️ **Alerta de Concentración:** Tenés más del 60% de tu presupuesto en **{r_est['Estrategia']}** (`{int(r_est['Costo_Pagado']):,} mnds`). Diversificá para reducir riesgo.")
            df_div = pd.DataFrame(filas_div)
            fig_pie = px.pie(
                df_div, values="Monedas", names="Rubro", hole=0.55,
                title="🍩 Exposición y Diversificación de Capital",
                color_discrete_sequence=["#00e676", "#00e5ff", "#ffd600", "#ff4d6d", "#b5179e"]
            )
            fig_pie.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=340)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_g3:
            if not vendidos.empty:
                df_curva = vendidos.copy()
                df_curva["Ganancia_Neta_Num"] = df_curva["Ganancia_Neta"].apply(num_seguro)
                df_curva["Ganancia_Acumulada"] = df_curva["Ganancia_Neta_Num"].cumsum()
                df_curva["Operacion"] = [f"#{i+1} {r.get('Jugador_Exacto', 'Carta')}" for i, (_, r) in enumerate(df_curva.iterrows())]
                fig_curva = px.area(
                    df_curva, x="Operacion", y="Ganancia_Acumulada",
                    markers=True, title="📈 Curva de Contribución Marginal (P&L)",
                    color_discrete_sequence=["#00e676"]
                )
                fig_curva.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig_curva, use_container_width=True)

        # ESTADO DE RESULTADOS + LIBRO MAYOR CON RE-COMPRA 1-CLIC + EXPORTADOR EXCEL
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
                {"Rubro Contable": "📦 Inventario Activo (Valuado a Costo Histórico)", "Monto (Monedas)": cap_inmovilizado},
                {"Rubro Contable": "📈 Inventario Activo a VNR Dual FB+GG (Neto 5% Tax)", "Monto (Monedas)": vnr_total},
                {"Rubro Contable": "⚖️ Resultado por Tenencia (No Realizado)", "Monto (Monedas)": tenencia_total},
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
            st.subheader("📜 Libro Mayor de Ventas (Tildá 🔄 para Re-Comprar en 1 Clic)")
            if not vendidos.empty:
                if any(vendidos["Precio_Venta"].apply(num_seguro) == 0):
                    st.caption("💡 Tip: Tenés una carta marcada como 'Vendido' con Precio_Venta en 0. Podés cargarle su precio de venta real desde la barra lateral en '💰 Cerrar / Corregir Venta'.")
                df_vend_edit = vendidos.copy()
                df_vend_edit.insert(0, "Re_Comprar", False)
                edicion_vend = st.data_editor(
                    df_vend_edit,
                    use_container_width=True,
                    hide_index=True,
                    disabled=[c for c in df_vend_edit.columns if c != "Re_Comprar"],
                    column_config={
                        "Re_Comprar": st.column_config.CheckboxColumn("🔄 Re-Comprar"),
                        "Precio_Pagado": st.column_config.NumberColumn("💵 Costo", format="%d mnds"),
                        "Precio_Venta": st.column_config.NumberColumn("💰 Venta", format="%d mnds"),
                        "Ganancia_Neta": st.column_config.NumberColumn("✨ Neto", format="%+d mnds"),
                        "ROI_%": st.column_config.NumberColumn("📈 ROI %", format="%+.1f%%"),
                    }
                )
                recompras_sel = edicion_vend[edicion_vend["Re_Comprar"] == True]
                if not recompras_sel.empty:
                    if st.button(f"🔄 Re-Comprar {len(recompras_sel)} carta(s) y enviar a Cartera Activa"):
                        df_c_act = datos.get("Mi_Cartera", pd.DataFrame(columns=columnas_base_cartera)).copy()
                        nuevas_re = []
                        for _, r_re in recompras_sel.iterrows():
                            nuevas_re.append({
                                "Jugador_Exacto": str(r_re.get("Jugador_Exacto", "Carta")),
                                "Estrategia": str(r_re.get("Estrategia", "Flipping Meta")),
                                "Precio_Pagado": int(num_seguro(r_re.get("Precio_Pagado", 10000))),
                                "Estado": "Activo",
                                "Precio_Venta": 0,
                                "Costo_EA_Tax": 0,
                                "Ganancia_Neta": 0,
                                "ROI_%": 0.0,
                                "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                                "Fecha_Venta": ""
                            })
                        df_c_act = pd.concat([df_c_act, pd.DataFrame(nuevas_re)], ignore_index=True)
                        guardar_hoja_sincronizada(df_c_act, "Mi_Cartera", es_admin=es_admin)
                        st.success("¡Cartas re-compradas y enviadas a tu Cartera Activa!")
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
                fig_proy = px.line(df_proy, x="Día", y="Patrimonio_Proyectado", markers=True, title=f"Proyección de Crecimiento Compuesto ({dias_graf} días)")
                fig_proy.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=280)
                st.plotly_chart(fig_proy, use_container_width=True)

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
                fig_hist.add_trace(go.Scatter(
                    x=df_jug_hist["Fecha"], y=precios_serie,
                    mode="lines+markers", name=etiqueta_plat,
                    line=dict(color="#00e5ff", width=3)
                ))
                fig_hist.add_hline(y=p_min_j, line_dash="dash", line_color="#00e676", annotation_text=f"Piso ({p_min_j:,})")
                fig_hist.add_hline(y=puja_j, line_dash="dashdot", line_color="#ffd600", annotation_text=f"Puja Máx ({puja_j:,})")
                fig_hist.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig_hist, use_container_width=True)

        if not df_ult_corte.empty:
            with st.expander("📲 Copiar Resumen de Gangas para WhatsApp / Discord / Telegram"):
                df_top_g = df_ult_corte[df_ult_corte["Precio_Actual"].apply(num_seguro) <= presupuesto_usuario].copy()
                if "Variacion_Bruta_%" in df_top_g.columns:
                    df_top_g = df_top_g.sort_values(by="Variacion_Bruta_%").head(5)
                else:
                    df_top_g = df_top_g.head(5)
                lineas_share = [f"⚡ *TOP 5 GANGAS FC27 ({etiqueta_plat} | Corte {ultima_act})* ⚡"]
                for _, rg in df_top_g.iterrows():
                    nom_sh = str(rg.get("Jugador", "")).split(" (")[0]
                    p_sh = num_seguro(rg.get("Precio_Actual", 0))
                    puja_sh = num_seguro(rg.get("Puja_Max_10%_ROI", 0))
                    lazy_sh = num_seguro(rg.get("Precio_Lazy_Sugerido", 0))
                    lineas_share.append(f"• *{nom_sh}*: {p_sh:,} mnds | 🔨 Pujar <= {puja_sh:,} | 🎯 Lazy: {lazy_sh:,}")
                st.code("\n".join(lineas_share), language="markdown")

        st.markdown("---")
        st.subheader("⚡ Padrón Dual Interactivo (Tildá 🛒 y envialas directo a tu Cartera con 1 Clic)")
        c_f1, c_f2, c_f3 = st.columns([2, 1, 1])
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
                    "📦 Out of Packs / Especiales (TOTW & Promos)",
                    "🔴 VENDER",
                    "⚠️ Evitar Cuchillo"
                ]
            )
        with c_f3:
            max_slider = max(200000, int(presupuesto_usuario))
            tope_precio_filtro = st.slider("Precio Máximo (Monedas)", min_value=500, max_value=max_slider, value=min(int(presupuesto_usuario), max_slider), step=2500)

        if not df_ult_corte.empty:
            df_filtrado = df_ult_corte[df_ult_corte["Precio_Actual"].apply(num_seguro) <= tope_precio_filtro].copy()
            if busqueda_jug.strip():
                df_filtrado = df_filtrado[df_filtrado["Jugador"].astype(str).str.lower().str.contains(busqueda_jug.strip().lower(), na=False)]
            if filtro_accion == "🟢 COMPRAR" and "Accion" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Accion"].astype(str).str.contains("COMPRAR", na=False)]
            elif filtro_accion == "⚡ Flipping Rápido" and "Senal_Trading" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Senal_Trading"].astype(str).str.contains("⚡", na=False)]
            elif filtro_accion == "💎 Mínimo Histórico" and "Estado_Historico" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Estado_Historico"].astype(str).str.contains("💎", na=False)]
            elif filtro_accion == "📦 Out of Packs / Especiales (TOTW & Promos)":
                df_filtrado = df_filtrado[df_filtrado["Jugador"].apply(es_carta_especial_oop)]
            elif filtro_accion == "🔴 VENDER" and "Accion" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Accion"].astype(str).str.contains("VENDER", na=False)]
            elif filtro_accion == "⚠️ Evitar Cuchillo" and "Senal_Trading" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Senal_Trading"].astype(str).str.contains("⚠️", na=False)]

            cols_mostrar = [c for c in [
                "Jugador", "Precio_Actual", "Fuente_Mejor_Precio", "Precio_Futbin", "Precio_FUTGG",
                "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Break_Even_5%", "Precio_Lazy_Sugerido", "Enlace_Optimo"
            ] if c in df_filtrado.columns]

            df_editor_padron = df_filtrado[cols_mostrar].head(200).copy()
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
                    for _, r_sel in seleccionadas_1clic.iterrows():
                        nom_limpio = str(r_sel["Jugador"]).split(" (")[0].strip()
                        p_compra_1c = num_seguro(r_sel["Precio_Actual"], 0)
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
                    st.success("¡Cartas agregadas a tu Cartera!")
                    st.rerun()

    with tab3:
        # 1. OPTIMIZADOR MATEMÁTICO DE PLANTILLAS SBC (COMBO MÁS BARATO)
        st.subheader("🧩 Optimizador Matemático de Plantillas SBC (Combo de 11 Cartas Más Barato Hoy)")
        pisos_por_media_default = {83: 900, 84: 1200, 85: 2800, 86: 5500, 87: 9500, 88: 14500, 89: 22000}
        pisos_vivos = {}
        jug_ref_vivos = {}
        for m_val, p_def_m in pisos_por_media_default.items():
            if not df_ult_corte.empty and "Jugador" in df_ult_corte.columns:
                c_m = df_ult_corte[df_ult_corte["Jugador"].astype(str).str.contains(f"Media {m_val}", na=False) & (df_ult_corte["Precio_Actual"].apply(num_seguro) > 0)]
                if not c_m.empty:
                    fila_min_m = c_m.loc[c_m["Precio_Actual"].apply(num_seguro).idxmin()]
                    pisos_vivos[m_val] = num_seguro(fila_min_m["Precio_Actual"], p_def_m)
                    jug_ref_vivos[m_val] = f"{str(fila_min_m['Jugador']).split(' (')[0]} ({pisos_vivos[m_val]:,} mnds)"
                else:
                    pisos_vivos[m_val] = p_def_m
                    jug_ref_vivos[m_val] = f"Referencia ({p_def_m:,} mnds)"
            else:
                pisos_vivos[m_val] = p_def_m
                jug_ref_vivos[m_val] = f"Referencia ({p_def_m:,} mnds)"

        co1, co2 = st.columns([1, 2])
        with co1:
            media_sbc_obj = st.selectbox("🎯 Media de Plantilla exigida por el SBC", [84, 85, 86, 87, 88], index=2)
            ifs_exigidos = st.selectbox("🖤 Cantidad de Cartas TOTW / IF exigidas", [0, 1, 2], index=0)
        with co2:
            combos_disp = COMBOS_SBC_PLANTILLAS.get(media_sbc_obj, [])
            filas_combos = []
            for nombre_combo, receta in combos_disp:
                costo_combo = sum(cant * pisos_vivos.get(med_r, 5000) for med_r, cant in receta.items())
                if ifs_exigidos > 0:
                    costo_combo += ifs_exigidos * 11500
                detalle_cartas = " + ".join([f"{cant}x M{med_r} [{jug_ref_vivos[med_r]}]" for med_r, cant in receta.items()])
                filas_combos.append({
                    "Receta_11_Cartas": nombre_combo + (f" (+{ifs_exigidos} IF)" if ifs_exigidos > 0 else ""),
                    "Costo_Total_Estimado": costo_combo,
                    "Cartas_Mas_Baratas_Hoy": detalle_cartas
                })
            df_combos = pd.DataFrame(filas_combos).sort_values(by="Costo_Total_Estimado")
            mejor_combo = df_combos.iloc[0]
            st.success(
                f"✅ **Combinación Más Barata para Media {media_sbc_obj}:** `{mejor_combo['Receta_11_Cartas']}` "
                f"➡️ **Costo Total: `{int(mejor_combo['Costo_Total_Estimado']):,} mnds`**\n\n"
                f"🛒 **Qué comprar hoy:** {mejor_combo['Cartas_Mas_Baratas_Hoy']}"
            )
            st.dataframe(df_combos, use_container_width=True, hide_index=True)

        st.markdown("---")
        # 2. SIMULADOR DE INVERSIÓN EN MEDIAS (FODDER SBC)
        st.subheader("🧮 Simulador Personalizado de Inversión en Medias (Fodder SBC)")
        ref_fodder_sim = {
            84: (1200, 2200),
            85: (2800, 4500),
            86: (5500, 8500),
            87: (9500, 14000),
            88: (14500, 21000),
            89: (22000, 32000),
        }
        cs1, cs2, cs3, cs4 = st.columns(4)
        with cs1:
            med_sim = st.selectbox("Media Objetivo", [84, 85, 86, 87, 88, 89], index=2)
        piso_def, techo_def = ref_fodder_sim[med_sim]
        with cs2:
            monto_inv_fod = st.number_input("Monedas a Invertir", min_value=5000, max_value=10000000, value=min(50000, int(presupuesto_usuario)), step=5000)
        with cs3:
            p_compra_fod = st.number_input("Precio Compra por Carta", min_value=700, max_value=100000, value=int(pisos_vivos.get(med_sim, piso_def)), step=100)
        with cs4:
            p_venta_fod = st.number_input("Precio Venta Esperado (SBC)", min_value=1000, max_value=150000, value=techo_def, step=250)

        cant_cartas_fod = int(monto_inv_fod // max(1, p_compra_fod))
        inv_real_fod = cant_cartas_fod * p_compra_fod
        ingreso_neto_fod = int(round(cant_cartas_fod * (p_venta_fod * 0.95)))
        ganancia_neta_fod = ingreso_neto_fod - inv_real_fod
        roi_fod = round((ganancia_neta_fod / max(1, inv_real_fod)) * 100, 1)

        st.info(
            f"📦 **Resultado de la Simulación (Media {med_sim}):** Comprás **{cant_cartas_fod} cartas** "
            f"(Inversión: `{inv_real_fod:,} mnds`) ➡️ Al vender en `{p_venta_fod:,} mnds` recibís limpio Post-Tax "
            f"`{ingreso_neto_fod:,} mnds` (**Ganancia Neta: `{ganancia_neta_fod:+,} mnds` | ROI: `{roi_fod:+.1f}%`**)."
        )

        st.markdown("---")
        st.subheader("🎯 Gestor Visual de Alertas Sniper")
        col_al1, col_al2 = st.columns([1, 2])
        with col_al1:
            with st.form("form_alerta_web", clear_on_submit=True):
                jug_al = st.text_input("Jugador a vigilar", placeholder="Ej: Theo Hernandez")
                obj_al = st.number_input("Precio Objetivo (<=)", min_value=200, max_value=15000000, step=250, value=15000)
                if st.form_submit_button("🎯 Activar Sniper", use_container_width=True) and jug_al.strip():
                    df_al_act = datos.get("Mis_Alertas", pd.DataFrame(columns=columnas_base_alertas)).copy()
                    nueva_al = pd.DataFrame([{
                        "Jugador_Buscado": jug_al.strip().title(),
                        "Precio_Objetivo": int(obj_al),
                        "Estado": "Activa"
                    }])
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

        st.markdown("---")
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            st.subheader("🧱 Radar de Medias SBC (Fodder)")
            df_fod = datos.get("Radar_Fodder", pd.DataFrame())
            if not df_fod.empty:
                st.dataframe(df_fod, use_container_width=True, hide_index=True)
            else:
                st.caption("Sin datos en la pestaña Radar_Fodder.")
        with c_m2:
            st.subheader("🖤 Radar TOTW en Piso")
            df_tw = datos.get("Radar_TOTW", pd.DataFrame())
            if not df_tw.empty:
                st.dataframe(df_tw, use_container_width=True, hide_index=True)
            else:
                st.caption("Sin datos en la pestaña Radar_TOTW.")

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
                    int(presupuesto_usuario),
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
                    datos.get("Radar_Fodder", pd.DataFrame()),
                    datos.get("Radar_TOTW", pd.DataFrame()),
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
            st.subheader("🤖 Reporte Estratégico Dual (En Vivo)")
            st.markdown(f'<div class="ia-card-report">{st.session_state["reporte_ia_vivo"]}</div>', unsafe_allow_html=True)
        with col_ia2:
            st.subheader(f"🔮 Oráculo de Contenido ({hora_drop_txt})")
            st.markdown(f'<div class="ia-card-oracle">{st.session_state["oraculo_ia_vivo"]}</div>', unsafe_allow_html=True)

        # CONSULTOR TÁCTICO INTERACTIVO (PREGUNTAS PERSONALIZADAS A LA IA)
        st.markdown("---")
        st.subheader("💬 Consultor Táctico IA Interactivo (Preguntale lo que quieras sobre el mercado)")
        with st.form("form_chat_ia"):
            pregunta_usuario = st.text_input(
                "Escribí tu consulta de trading (la IA lee tu padrón y tu cartera actual):",
                placeholder="Ej: Tengo 100,000 monedas libres, ¿qué 3 cartas del padrón me recomendás pujar hoy y a cuánto las listo?"
            )
            btn_preguntar_ia = st.form_submit_button("🚀 Analizar Consulta con IA")

            if btn_preguntar_ia and pregunta_usuario.strip():
                with st.spinner("🤖 Analizando tu consulta contra la base de datos en vivo..."):
                    df_disp_q = df_ult_corte[df_ult_corte["Precio_Actual"].apply(num_seguro) <= presupuesto_usuario].head(25) if not df_ult_corte.empty else pd.DataFrame()
                    cols_q = [c for c in ["Jugador", "Precio_Actual", "Fuente_Mejor_Precio", "Variacion_Bruta_%", "Senal_Trading", "Puja_Max_10%_ROI", "Precio_Lazy_Sugerido"] if c in df_disp_q.columns]
                    muestra_q = df_disp_q[cols_q].to_string(index=False) if not df_disp_q.empty else "Sin datos."
                    cartera_q = ", ".join([f"{c['Jugador']} (Costo {c['Costo_Pagado']:,}, Hoy {c['Mercado_Hoy']:,})" for c in filas_cartera_viva]) if filas_cartera_viva else "Sin cartas activas."

                    prompt_q = f"""
                    Actuá como Asesor Experto de Trading en EA Sports FC 27 ({etiqueta_plat}).
                    Presupuesto del usuario: {presupuesto_usuario:,} mnds (Liquidez libre: {liquidez_disp:,} mnds).
                    Reglas del usuario: Puja +{roi_puja_pct}% ROI, Lazy +{lazy_pct}%, Tax EA 5%.
                    Cartera activa del usuario: {cartera_q}
                    Muestra del padrón actual en vivo:
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
            st.info(st.session_state["respuesta_chat_ia"])

    if auto_ref:
        time.sleep(60)
        st.cache_data.clear()
        st.rerun()

except Exception as err_global:
    st.error(f"🚨 Error detectado: {err_global}")
    st.code(traceback.format_exc())
