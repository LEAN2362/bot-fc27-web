import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import math
import time
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

# ⬇️ SI QUERÉS DEJAR TU LINK FIJO, PEGALO ACÁ ENTRE LAS COMILLAS ⬇️
LINK_GOOGLE_SHEET = ""

PRESUPUESTO_TOPE = 200000
ARCHIVO_EXCEL = "Lista_Jugadores_FC27.xlsx"
ARCHIVO_LINK_CACHE = "sheet_link_cache.txt"
ARCHIVO_CARTERA_LOCAL = "cartera_cloud_sync.csv"
ARCHIVO_ALERTAS_LOCAL = "alertas_cloud_sync.csv"
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
    div.stButton > button, div[data-testid="stFormSubmitButton"] > button {
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
    div[data-baseweb="tab-list"] { gap: 8px; background-color: #111625; padding: 8px; border-radius: 12px; border: 1px solid #26334d; }
    button[data-baseweb="tab"] { font-size: 0.95rem !important; font-weight: 700 !important; color: #94a3b8 !important; border-radius: 8px !important; padding: 10px 16px !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background: linear-gradient(90deg, #6200ea, #00b0ff) !important; color: #ffffff !important; }
    h2, h3 { color: #38bdf8 !important; font-weight: 800 !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIONES MATEMÁTICAS, HORA ARG Y ESTILOS
# ==========================================
def ahora_arg():
    tz_arg = timezone(timedelta(hours=-3))
    return datetime.now(tz_arg).replace(tzinfo=None)

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

def calcular_countdown_15hs():
    ahora = ahora_arg()
    drop_hoy = ahora.replace(hour=15, minute=0, second=0, microsecond=0)
    if ahora >= drop_hoy:
        proximo_drop = drop_hoy + timedelta(days=1)
        estado_drop = "🔥 Contenido de hoy ya activo"
    else:
        proximo_drop = drop_hoy
        if ahora.hour == 14 and ahora.minute >= 30:
            estado_drop = "🚨 ¡VENTANA PRE-DROP! (Comprar en pánico)"
        else:
            estado_drop = "⏳ Esperando Drop Diario EA"
    resta = proximo_drop - ahora
    horas, rem = divmod(int(resta.total_seconds()), 3600)
    minutos, _ = divmod(rem, 60)
    return f"{horas:02d}h {minutos:02d}m", estado_drop

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
        "Precio_Anterior", "Precio_Actual", "Precio_Futbin", "Precio_FUTGG",
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
                hojas[h] = convertir_columnas_numericas(df_h)
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

def guardar_hoja_sincronizada(df_guardar, nombre_hoja):
    st.session_state[f"override_{nombre_hoja}"] = df_guardar.copy()
    try:
        if nombre_hoja == "Mi_Cartera":
            df_guardar.to_csv(ARCHIVO_CARTERA_LOCAL, index=False)
        elif nombre_hoja == "Mis_Alertas":
            df_guardar.to_csv(ARCHIVO_ALERTAS_LOCAL, index=False)
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

    # --- INICIO DEL BLOQUE ACTUALIZADO (COMBINA SHEETS + GITHUB AUTO-SYNC) ---
    if sheet_id:
        datos, msg_error = cargar_por_id_publico(sheet_id)
        if msg_error == "OK":
            estado_conexion = "☁️ Conectado a Google Sheets (Nube 24/7)"

    if os.path.exists(ARCHIVO_EXCEL):
        try:
            df_hist_excel = convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name="Historial_Tendencias"))
            fecha_excel = str(df_hist_excel["Fecha"].max()) if not df_hist_excel.empty and "Fecha" in df_hist_excel.columns else ""
            fecha_sheet = str(datos.get("Historial_Tendencias", pd.DataFrame()).get("Fecha", pd.Series([""])).max()) if datos else ""

            if fecha_excel and (not datos or fecha_excel >= fecha_sheet):
                cartera_respaldo = datos.get("Mi_Cartera", pd.DataFrame()) if datos else pd.DataFrame()
                alertas_respaldo = datos.get("Mis_Alertas", pd.DataFrame()) if datos else pd.DataFrame()
                for h in [
                    "Historial_Tendencias", "Recomendacion_IA", "Dashboard_Financiero",
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
    # --- FIN DEL BLOQUE ACTUALIZADO ---

    # Integrar cambios guardados en la sesión o en el contenedor Cloud
    if datos:
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

    # ==========================================
    # BARRA LATERAL COMPLETA (SIDEBAR)
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
        st.markdown("### 🧮 Calculadora Rápida (5% Tax)")
        precio_sim = st.number_input("Precio visto en Mercado (PC)", min_value=200, max_value=2000000, step=250, value=25000)
        puja_sim = redondear_escalon_ea((precio_sim * 0.95) / 1.10, hacia_arriba=False)
        empate_sim = redondear_escalon_ea(precio_sim / 0.95, hacia_arriba=True)
        lazy_sim = redondear_escalon_ea(precio_sim * 1.12, hacia_arriba=True)
        sl_sim = redondear_escalon_ea((precio_sim * 0.90) / 0.95, hacia_arriba=False)
        neto_lazy_sim = int(round((lazy_sim * 0.95) - precio_sim))

        st.markdown(f"""
        <div class="calc-box">
            🔨 <b>Pujar hasta (+10% ROI):</b> <span style="color:#00e676;font-weight:bold;">{puja_sim:,} mnds</span><br>
            ⚖️ <b>Empate si comprás a {precio_sim:,}:</b> <span style="color:#84ffff;">{empate_sim:,} mnds</span><br>
            🎯 <b>Listar Lazy 1h (+12%):</b> <span style="color:#ffea00;font-weight:bold;">{lazy_sim:,} mnds</span><br>
            ✨ <b>Ganancia Neta en Lazy:</b> <span style="color:#00e676;">+{neto_lazy_sim:,} mnds</span><br>
            🛑 <b>Stop-Loss (-10%):</b> <span style="color:#ff5252;">{sl_sim:,} mnds</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### ➕ Registrar Compra Manual")
        with st.form("form_compra", clear_on_submit=True):
            jug_c = st.text_input("Nombre del Jugador", placeholder="Ej: Odegaard")
            precio_c = st.number_input("Precio Pagado (Monedas)", min_value=200, max_value=2000000, step=250, value=10000)
            est_c = st.selectbox("Unidad de Negocio", ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"])
            btn_guardar_c = st.form_submit_button("✅ Guardar en Cartera", use_container_width=True)

            if btn_guardar_c and jug_c.strip():
                df_c_act = datos.get("Mi_Cartera", pd.DataFrame()).copy()
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
                guardar_hoja_sincronizada(df_c_act, "Mi_Cartera")
                st.success(f"¡{jug_c.title()} agregado a Mi Cartera!")
                st.rerun()

        st.markdown("---")
        st.markdown("### 💰 Cerrar / Corregir Venta")
        df_cartera_sb = datos.get("Mi_Cartera", pd.DataFrame()).copy()
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
                precio_v = st.number_input("Precio de Venta Bruto", min_value=200, max_value=2000000, step=250, value=15000)
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

                    guardar_hoja_sincronizada(df_cartera_sb, "Mi_Cartera")
                    st.success(f"¡Venta registrada! Neto: {neto:+,} mnds")
                    st.rerun()
        else:
            st.caption("No tenés cartas pendientes de venta.")

        with st.expander("🔗 Vínculo de Google Sheets"):
            link_sb = st.text_input("URL de tu planilla:", value=link_activo)
            if st.button("Guardar Vínculo", use_container_width=True):
                st.session_state["link_manual"] = link_sb.strip()
                guardar_link_cache(link_sb.strip())
                st.cache_data.clear()
                st.rerun()

    # ==========================================
    # CABECERA PRINCIPAL + RELOJ 15:00 HS ARG
    # ==========================================
    df_historial = datos.get("Historial_Tendencias", pd.DataFrame())
    ultima_act = str(df_historial["Fecha"].max()) if not df_historial.empty and "Fecha" in df_historial.columns else "Sin escaneo"
    df_ult_corte = df_historial[df_historial["Fecha"] == ultima_act].copy() if not df_historial.empty and "Fecha" in df_historial.columns else pd.DataFrame()
    tiempo_drop, leyenda_drop = calcular_countdown_15hs()

    st.markdown(f"""
    <div class="main-header">
        <div>
            <h1>⚡ FC 27 TRADING TERMINAL PRO — CLOUD 24/7</h1>
            <p>{estado_conexion} &nbsp;|&nbsp; 📦 Padrón: <b>{len(df_ult_corte):,} cartas</b> &nbsp;|&nbsp; 🕒 Último corte: <b>{ultima_act}</b></p>
        </div>
        <div class="drop-timer">
            <div class="drop-timer-title">🎯 Próximo Drop EA (15:00 ARG)</div>
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
    # TARJETAS KPI FINANCIERAS MULTICOLOR
    # ==========================================
    df_cartera = datos.get("Mi_Cartera", pd.DataFrame()).copy()
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
    liquidez_disp = max(0, PRESUPUESTO_TOPE - cap_inmovilizado)
    ganancia_realizada = int(vendidos["Ganancia_Neta"].apply(num_seguro).sum()) if not vendidos.empty and "Ganancia_Neta" in vendidos.columns else 0
    aciertos = len(vendidos[vendidos["Ganancia_Neta"].apply(num_seguro) > 0]) if not vendidos.empty and "Ganancia_Neta" in vendidos.columns else 0
    efectividad = round((aciertos / len(vendidos) * 100), 1) if not vendidos.empty else 0.0

    filas_cartera_viva = []
    vnr_total = 0
    for _, r_a in activos.iterrows():
        nom_a = str(r_a.get("Jugador_Exacto", "Jugador"))
        costo_a = num_seguro(r_a.get("Precio_Pagado", 0), 0)
        est_a = str(r_a.get("Estrategia", "Flipping Meta"))
        empate_a = redondear_escalon_ea(costo_a / 0.95, hacia_arriba=True)
        sl_a = redondear_escalon_ea((costo_a * 0.90) / 0.95, hacia_arriba=False)

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
        lazy_a = redondear_escalon_ea(max(p_mercado * 1.12, (costo_a * 1.10) / 0.95), hacia_arriba=True)

        if roi_lat <= -10.0:
            estado_op = "🛑 STOP-LOSS (Vender)"
        elif tenencia > 0:
            estado_op = "🚨 VENDER (En Ganancia)"
        elif horas_hold >= 72.0:
            estado_op = f"🔄 ROTAR STOCK ({str_hold})"
        else:
            estado_op = f"⏳ HOLD / LAZY ({str_hold})"

        filas_cartera_viva.append({
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

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card kpi-green"><div class="kpi-title">💵 Liquidez Disponible</div><div class="kpi-value">{liquidez_disp:,} mnds</div><div class="kpi-badge">🟢 {round(liquidez_disp/PRESUPUESTO_TOPE*100)}% del tope</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-blue"><div class="kpi-title">📦 Capital Inmovilizado</div><div class="kpi-value">{cap_inmovilizado:,} mnds</div><div class="kpi-badge">🃏 {len(activos)} cartas en club</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card {color_vnr}"><div class="kpi-title">📈 Stock a VNR (5% Tax)</div><div class="kpi-value">{vnr_total:,} mnds</div><div class="kpi-badge">⚖️ Tenencia: {tenencia_total:+,}</div></div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card kpi-gold"><div class="kpi-title">💰 Ganancia Realizada</div><div class="kpi-value">{ganancia_realizada:+,} mnds</div><div class="kpi-badge">📈 {len(vendidos)} ventas cerradas</div></div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""<div class="kpi-card kpi-purple"><div class="kpi-title">🎯 Efectividad Trading</div><div class="kpi-value">{efectividad}%</div><div class="kpi-badge">🏆 {aciertos}/{len(vendidos)} positivos</div></div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # PESTAÑAS INTERACTIVAS COMPLETAS
    # ==========================================
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Cartera, P&L & Curva Capital",
        "⚡ Radar 1-Clic & Lupa Histórica",
        "🎯 Alertas Sniper, Fodder & Evos",
        "🔮 Oráculo & Reporte IA"
    ])

    with tab1:
        st.subheader("🎯 Monitoreo de Cartera Activa (Semáforo en Vivo con Rotación, Lazy y Stop-Loss)")
        if filas_cartera_viva:
            df_viva = pd.DataFrame(filas_cartera_viva)
            styled_viva = aplicar_estilo_seguro(
                df_viva,
                cols_neto=["Resultado_Tenencia", "ROI_Latente_%"],
                cols_accion=["Acción_Sugerida"]
            )
            st.dataframe(
                styled_viva, use_container_width=True, hide_index=True,
                column_config={
                    "Costo_Pagado": st.column_config.NumberColumn("💵 Costo", format="%d mnds"),
                    "Empate_5%": st.column_config.NumberColumn("⚖️ Empate (5%)", format="%d mnds"),
                    "Mercado_Hoy": st.column_config.NumberColumn("📊 Mercado Hoy", format="%d mnds"),
                    "Precio_Lazy_1h": st.column_config.NumberColumn("🎯 Lazy 1h (+12%)", format="%d mnds"),
                    "Stop_Loss": st.column_config.NumberColumn("🛑 Stop-Loss", format="%d mnds"),
                    "Resultado_Tenencia": st.column_config.NumberColumn("✨ Neto Hoy", format="%+d mnds"),
                    "ROI_Latente_%": st.column_config.NumberColumn("📈 ROI %", format="%+.1f%%"),
                    "Link": st.column_config.LinkColumn("🔗 Web", display_text="Abrir Carta")
                }
            )
        else:
            st.info("No tenés cartas activas en cartera.")

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

        with col_pnl2:
            st.subheader("📜 Libro Mayor de Ventas Cerradas")
            if not vendidos.empty:
                if any(vendidos["Precio_Venta"].apply(num_seguro) == 0):
                    st.caption("💡 Tip: Tenés una carta marcada como 'Vendido' con Precio_Venta en 0. Podés cargarle su precio de venta real desde la barra lateral en '💰 Cerrar / Corregir Venta'.")
                styled_vend = aplicar_estilo_seguro(vendidos, cols_neto=["Ganancia_Neta", "ROI_%"])
                st.dataframe(styled_vend, use_container_width=True, hide_index=True)
            else:
                st.info("Todavía no hay ventas cerradas registradas.")

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
                puja_j = redondear_escalon_ea((p_actual_j * 0.95) / 1.10, hacia_arriba=False)
                lazy_j = redondear_escalon_ea(p_actual_j * 1.12, hacia_arriba=True)

                m_l1, m_l2, m_l3, m_l4, m_l5 = st.columns(5)
                m_l1.metric("💰 Precio Actual", f"{p_actual_j:,} mnds")
                m_l2.metric("💎 Piso Mínimo Hist.", f"{p_min_j:,} mnds")
                m_l3.metric("⚖️ Promedio Hist.", f"{p_prom_j:,} mnds")
                m_l4.metric("🔨 Pujar hasta (+10%)", f"{puja_j:,} mnds")
                m_l5.metric("🎯 Target Lazy 1h", f"{lazy_j:,} mnds")

                fig_hist = go.Figure()
                fig_hist.add_trace(go.Scatter(
                    x=df_jug_hist["Fecha"], y=precios_serie,
                    mode="lines+markers", name="Precio PC",
                    line=dict(color="#00e5ff", width=3)
                ))
                fig_hist.add_hline(y=p_min_j, line_dash="dash", line_color="#00e676", annotation_text=f"Piso ({p_min_j:,})")
                fig_hist.add_hline(y=puja_j, line_dash="dashdot", line_color="#ffd600", annotation_text=f"Puja Máx ({puja_j:,})")
                fig_hist.update_layout(template="plotly_dark", paper_bgcolor="rgba(17,22,37,0.7)", plot_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig_hist, use_container_width=True)

        st.markdown("---")
        st.subheader("⚡ Padrón Dual Interactivo (Tildá 🛒 y envialas directo a Mi Cartera con 1 Clic)")
        c_f1, c_f2, c_f3 = st.columns([2, 1, 1])
        with c_f1:
            busqueda_jug = st.text_input("🔍 Filtrar por Nombre, Media o Tipo de Carta", placeholder="Ej: 86, TOTW, Mbappe...")
        with c_f2:
            filtro_accion = st.selectbox("Filtrar por Señal", ["Todas", "🟢 COMPRAR", "⚡ Flipping Rápido", "💎 Mínimo Histórico", "🔴 VENDER", "⚠️ Evitar Cuchillo"])
        with c_f3:
            tope_precio_filtro = st.slider("Precio Máximo (Monedas)", min_value=500, max_value=200000, value=200000, step=2500)

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
                    "Puja_Max_10%_ROI": st.column_config.NumberColumn("🔨 Puja Máx", format="%d mnds"),
                    "Precio_Lazy_Sugerido": st.column_config.NumberColumn("🎯 Lazy 1h", format="%d mnds"),
                    "Enlace_Optimo": st.column_config.LinkColumn("🔗 Link", display_text="Ver Carta")
                }
            )

            seleccionadas_1clic = edicion_padron[edicion_padron["Comprar_1Clic"] == True]
            if not seleccionadas_1clic.empty:
                if st.button(f"📥 Enviar {len(seleccionadas_1clic)} carta(s) a Mi Cartera"):
                    df_c_act = datos.get("Mi_Cartera", pd.DataFrame()).copy()
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
                    guardar_hoja_sincronizada(df_c_act, "Mi_Cartera")
                    st.success("¡Cartas agregadas a Mi Cartera!")
                    st.rerun()

    with tab3:
        st.subheader("🎯 Gestor Visual de Alertas Sniper")
        col_al1, col_al2 = st.columns([1, 2])
        with col_al1:
            with st.form("form_alerta_web", clear_on_submit=True):
                jug_al = st.text_input("Jugador a vigilar", placeholder="Ej: Theo Hernandez")
                obj_al = st.number_input("Precio Objetivo (<=)", min_value=200, max_value=2000000, step=250, value=15000)
                if st.form_submit_button("🎯 Activar Sniper", use_container_width=True) and jug_al.strip():
                    df_al_act = datos.get("Mis_Alertas", pd.DataFrame()).copy()
                    nueva_al = pd.DataFrame([{
                        "Jugador_Buscado": jug_al.strip().title(),
                        "Precio_Objetivo": int(obj_al),
                        "Estado": "Activa"
                    }])
                    df_al_act = pd.concat([df_al_act, nueva_al], ignore_index=True)
                    guardar_hoja_sincronizada(df_al_act, "Mis_Alertas")
                    st.rerun()

        with col_al2:
            df_alertas_w = datos.get("Mis_Alertas", pd.DataFrame()).copy()
            activas_al = df_alertas_w[df_alertas_w["Estado"].astype(str).str.strip() == "Activa"].copy() if not df_alertas_w.empty and "Estado" in df_alertas_w.columns else pd.DataFrame()
            if not activas_al.empty:
                st.dataframe(activas_al, use_container_width=True, hide_index=True)
                if st.button("🗑️ Limpiar Alertas Activas"):
                    df_alertas_w.loc[df_alertas_w["Estado"].astype(str).str.strip() == "Activa", "Estado"] = "Cancelada"
                    guardar_hoja_sincronizada(df_alertas_w, "Mis_Alertas")
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

    with tab4:
        df_ia = datos.get("Recomendacion_IA", pd.DataFrame())
        if not df_ia.empty:
            col_ia1, col_ia2 = st.columns(2)
            with col_ia1:
                st.subheader("🤖 Reporte Estratégico Dual (Gemini AI)")
                if "Reporte" in df_ia.columns and not df_ia["Reporte"].dropna().empty:
                    st.success(str(df_ia["Reporte"].dropna().iloc[-1]))
                else:
                    st.dataframe(df_ia, use_container_width=True, hide_index=True)
            with col_ia2:
                st.subheader("🔮 Oráculo de Contenido (15:00 hs ARG)")
                if "Oraculo" in df_ia.columns and not df_ia["Oraculo"].dropna().empty:
                    st.info(str(df_ia["Oraculo"].dropna().iloc[-1]))
                else:
                    st.info("Consultá las columnas de tu pestaña Recomendacion_IA.")
        else:
            st.info("No se encontraron registros en la pestaña Recomendacion_IA.")

    if auto_ref:
        time.sleep(60)
        st.cache_data.clear()
        st.rerun()

except Exception as err_global:
    st.error(f"🚨 Error detectado: {err_global}")
    st.code(traceback.format_exc())
