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
from datetime import datetime, timedelta

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="FC27 Trading Terminal Pro (Cloud)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ⬇️ PEGÁ ACÁ ENTRE LAS COMILLAS EL VÍNCULO DE TU GOOGLE SHEETS ("Copiar vínculo") ⬇️
LINK_GOOGLE_SHEET = "https://docs.google.com/spreadsheets/d/1JDSpJGB6SaPe53cTMMKCdr-OMzzDZPnPsXc4YiS1DjE/edit?usp=sharing"

PRESUPUESTO_TOPE = 200000
ARCHIVO_EXCEL = "Lista_Jugadores_FC27.xlsx"

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
    div.stButton > button, div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(90deg, #00c853 0%, #00e676 50%, #00b0ff 100%) !important;
        color: #070a12 !important;
        font-weight: 800 !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 10px 16px !important;
    }
    .main-header {
        background: linear-gradient(120deg, #0f2027 0%, #1c1042 50%, #004d40 100%);
        padding: 22px 28px;
        border-radius: 16px;
        border: 1px solid #00e676;
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
        color: #00e676;
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
        padding: 18px 20px;
        border-radius: 14px;
        color: white;
        margin-bottom: 10px;
    }
    .kpi-title { font-size: 0.85rem; font-weight: 700; text-transform: uppercase; margin-bottom: 6px; }
    .kpi-value { font-size: 1.7rem; font-weight: 900; margin-bottom: 8px; }
    .kpi-badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 700; background: rgba(0, 0, 0, 0.35); }
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
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNCIONES MATEMÁTICAS Y AUXILIARES
# ==========================================
def num_seguro(val, defecto=0):
    try:
        if pd.isna(val) or str(val).strip() == "":
            return defecto
        limpio = str(val).replace("$", "").replace(",", "").strip()
        return int(float(limpio))
    except Exception:
        return defecto

def redondear_escalon_ea(precio, hacia_arriba=False):
    precio = float(num_seguro(precio, 0))
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
    ahora = datetime.now()
    drop_hoy = ahora.replace(hour=15, minute=0, second=0, microsecond=0)
    if ahora >= drop_hoy:
        proximo_drop = drop_hoy + timedelta(days=1)
        estado_drop = "🔥 Contenido de hoy ya activo"
    else:
        proximo_drop = drop_hoy
        estado_drop = "⏳ Esperando Drop Diario EA"
    resta = proximo_drop - ahora
    horas, rem = divmod(int(resta.total_seconds()), 3600)
    minutos, _ = divmod(rem, 60)
    return f"{horas:02d}h {minutos:02d}m", estado_drop

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
        "Precio_Objetivo"
    ]
    for c in cols_num:
        if c in df.columns:
            s = df[c].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.strip()
            df[c] = pd.to_numeric(s, errors="coerce").fillna(0)
    return df

# ==========================================
# LECTURA DIRECTA SIN CREDENCIALES NI SECRETS
# ==========================================
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
                # Si Google devuelve la primera hoja por defecto porque la pestaña no existe, la descartamos
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

# ==========================================
# EJECUCIÓN PRINCIPAL
# ==========================================
try:
    if "link_manual" not in st.session_state:
        st.session_state["link_manual"] = LINK_GOOGLE_SHEET

    link_activo = st.session_state["link_manual"] or LINK_GOOGLE_SHEET
    sheet_id = extraer_id_sheet(link_activo)

    datos = {}
    msg_error = ""
    estado_conexion = "⚠️ Sin conexión a datos"

    if sheet_id:
        datos, msg_error = cargar_por_id_publico(sheet_id)
        if msg_error == "OK":
            estado_conexion = "☁️ Conectado a Google Sheets (Nube 24/7)"

    if not datos and os.path.exists(ARCHIVO_EXCEL):
        for h in ["Historial_Tendencias", "Mi_Cartera", "Recomendacion_IA", "Radar_Fodder", "Radar_TOTW", "Mis_Alertas"]:
            try:
                datos[h] = convertir_columnas_numericas(pd.read_excel(ARCHIVO_EXCEL, sheet_name=h))
            except Exception:
                datos[h] = pd.DataFrame()
        if not datos.get("Historial_Tendencias", pd.DataFrame()).empty:
            estado_conexion = "💻 Conectado a Excel Local"

    # BARRA LATERAL
    with st.sidebar:
        st.markdown("## 🛠️ Centro de Mando")
        st.caption(estado_conexion)
        if st.button("🔄 Sincronizar"):
            st.cache_data.clear()
            st.rerun()

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

    # CABECERA PRINCIPAL
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

    # SI TODAVÍA NO PUSO EL LINK, MOSTRAR CAJA DIRECTA EN PANTALLA
    if not datos:
        if msg_error and msg_error != "OK":
            st.error(f"🔍 {msg_error}")
        st.info("👇 **Conexión instantánea sin llaves:** En tu Google Sheets `Base_Datos_FC27`, tocá **Compartir** ➡️ cambiá 'Restringido' a **'Cualquier persona con el vínculo'** ➡️ tocá **Copiar vínculo** y pegalo acá abajo:")
        nuevo_link = st.text_input("🔗 Pegá acá el vínculo de tu Google Sheets:", value=link_activo, placeholder="https://docs.google.com/spreadsheets/d/...")
        if st.button("🚀 Conectar ahora"):
            st.session_state["link_manual"] = nuevo_link.strip()
            st.cache_data.clear()
            st.rerun()
        st.stop()

    # TARJETAS KPI
    df_cartera = datos.get("Mi_Cartera", pd.DataFrame())
    if not df_cartera.empty and "Estado" in df_cartera.columns:
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

        p_mercado = costo_a
        if not df_ult_corte.empty and "Jugador" in df_ult_corte.columns:
            m = df_ult_corte[df_ult_corte["Jugador"].astype(str).str.lower().str.contains(nom_a.lower(), na=False)]
            if not m.empty:
                fila_m = m.iloc[0]
                p_mercado = max(
                    num_seguro(fila_m.get("Precio_Actual", 0)),
                    num_seguro(fila_m.get("Precio_Futbin", 0)),
                    num_seguro(fila_m.get("Precio_FUTGG", 0))
                )

        vnr_carta = int(round(p_mercado * 0.95))
        vnr_total += vnr_carta
        tenencia = vnr_carta - costo_a
        roi_lat = round((tenencia / costo_a) * 100, 1) if costo_a > 0 else 0.0
        lazy_a = redondear_escalon_ea(max(p_mercado * 1.12, (costo_a * 1.10) / 0.95), hacia_arriba=True)

        filas_cartera_viva.append({
            "Jugador": nom_a,
            "Estrategia": est_a,
            "Costo_Pagado": costo_a,
            "Empate_5%": empate_a,
            "Mercado_Hoy": p_mercado,
            "Precio_Lazy_1h": lazy_a,
            "Stop_Loss": sl_a,
            "Resultado_Tenencia": tenencia,
            "ROI_Latente_%": roi_lat
        })

    tenencia_total = vnr_total - cap_inmovilizado
    color_vnr = "kpi-purple" if tenencia_total >= 0 else "kpi-crimson"

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card kpi-green"><div class="kpi-title">💵 Liquidez Disponible</div><div class="kpi-value">{liquidez_disp:,} mnds</div><div class="kpi-badge">🟢 {round(liquidez_disp/PRESUPUESTO_TOPE*100)}% libre</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-blue"><div class="kpi-title">📦 Capital Inmovilizado</div><div class="kpi-value">{cap_inmovilizado:,} mnds</div><div class="kpi-badge">🃏 {len(activos)} cartas</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card {color_vnr}"><div class="kpi-title">📈 Stock a VNR (Neto)</div><div class="kpi-value">{vnr_total:,} mnds</div><div class="kpi-badge">⚖️ Tenencia: {tenencia_total:+,}</div></div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card kpi-gold"><div class="kpi-title">💰 Ganancia Realizada</div><div class="kpi-value">{ganancia_realizada:+,} mnds</div><div class="kpi-badge">📈 {len(vendidos)} ventas</div></div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""<div class="kpi-card kpi-purple"><div class="kpi-title">🎯 Efectividad</div><div class="kpi-value">{efectividad}%</div><div class="kpi-badge">🏆 {aciertos}/{len(vendidos)} positivos</div></div>""", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Cartera & P&L",
        "⚡ Radar & Lupa Histórica",
        "🎯 Fodder & TOTW",
        "🔮 Oráculo IA"
    ])

    with tab1:
        st.subheader("🎯 Monitoreo de Cartera Activa")
        if filas_cartera_viva:
            st.dataframe(pd.DataFrame(filas_cartera_viva), hide_index=True)
        else:
            st.info("No tenés cartas activas en cartera.")

        st.subheader("📜 Libro Mayor de Ventas Cerradas")
        if not vendidos.empty:
            st.dataframe(vendidos, hide_index=True)

    with tab2:
        st.subheader("🔍 Lupa Visual Histórica")
        if not df_historial.empty and "Jugador" in df_historial.columns:
            lista_jugadores_unicos = sorted(df_historial["Jugador"].astype(str).unique().tolist())
            jug_seleccionado = st.selectbox("Elegí un jugador:", lista_jugadores_unicos)
            if jug_seleccionado:
                df_jug_hist = df_historial[df_historial["Jugador"] == jug_seleccionado].sort_values(by="Fecha")
                fig_hist = px.line(df_jug_hist, x="Fecha", y="Precio_Actual", markers=True, title=f"Evolución: {jug_seleccionado}")
                fig_hist.update_layout(template="plotly_dark")
                st.plotly_chart(fig_hist)

        st.subheader("⚡ Padrón Completo en Nube")
        if not df_ult_corte.empty:
            c_f1, c_f2 = st.columns(2)
            with c_f1:
                busq = st.text_input("🔍 Buscar jugador:", placeholder="Ej: Mendy, Mbappe...")
            with c_f2:
                tope_p = st.slider("Precio Máximo:", min_value=500, max_value=200000, value=200000, step=2500)
            df_mostrar = df_ult_corte[df_ult_corte["Precio_Actual"] <= tope_p].copy()
            if busq.strip():
                df_mostrar = df_mostrar[df_mostrar["Jugador"].astype(str).str.lower().str.contains(busq.strip().lower(), na=False)]
            st.dataframe(df_mostrar.head(250), hide_index=True)

    with tab3:
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            st.subheader("🧱 Radar Fodder")
            df_fod = datos.get("Radar_Fodder", pd.DataFrame())
            if not df_fod.empty:
                st.dataframe(df_fod, hide_index=True)
        with c_m2:
            st.subheader("🖤 Radar TOTW")
            df_tw = datos.get("Radar_TOTW", pd.DataFrame())
            if not df_tw.empty:
                st.dataframe(df_tw, hide_index=True)

    with tab4:
        df_ia = datos.get("Recomendacion_IA", pd.DataFrame())
        if not df_ia.empty:
            st.dataframe(df_ia, hide_index=True)

except Exception as err_global:
    st.error(f"🚨 Error detectado: {err_global}")
    st.code(traceback.format_exc())
