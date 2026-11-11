import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import math
import time
import traceback
import gspread
from datetime import datetime, timedelta

st.set_page_config(
    page_title="FC27 Trading Terminal Pro (Cloud)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

ARCHIVO_CRED = "credenciales_google.json"
NOMBRE_SHEET = "Base_Datos_FC27"
PRESUPUESTO_TOPE = 200000

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

def obtener_cliente_gsheets():
    if os.path.exists(ARCHIVO_CRED):
        try:
            return gspread.service_account(filename=ARCHIVO_CRED), None
        except Exception as e:
            return None, f"Error leyendo '{ARCHIVO_CRED}': {e}"
    return None, f"No se encontró el archivo '{ARCHIVO_CRED}' en tu repositorio de GitHub. Subilo con el botón 'Add file' -> 'Upload files'."

def leer_pestana_segura(ws):
    valores = ws.get_all_values()
    if not valores or len(valores) < 2:
        return pd.DataFrame()
    encabezados = [str(c).strip() for c in valores[0]]
    indices_validos = [i for i, h in enumerate(encabezados) if h != ""]
    if not indices_validos:
        return pd.DataFrame()
    cols_limpias = []
    usados = {}
    for i in indices_validos:
        nombre_col = encabezados[i]
        if nombre_col in usados:
            usados[nombre_col] += 1
            nombre_col = f"{nombre_col}_{usados[nombre_col]}"
        else:
            usados[nombre_col] = 0
        cols_limpias.append(nombre_col)

    filas_limpias = []
    for fila in valores[1:]:
        if any(str(celda).strip() != "" for celda in fila):
            filas_limpias.append([fila[i] if i < len(fila) else "" for i in indices_validos])
    return pd.DataFrame(filas_limpias, columns=cols_limpias)

def convertir_columnas_numericas(df):
    cols_num = [
        "Precio_Pagado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%",
        "Precio_Anterior", "Precio_Actual", "Precio_Futbin", "Precio_FUTGG",
        "Precio_Min_Hist", "Precio_Prom_Hist", "Puja_Max_10%_ROI", "Break_Even_5%",
        "Precio_Lazy_Sugerido", "Variacion_Bruta_%", "Beneficio_Neto_Mnds", "Variacion_Neta_%"
    ]
    for c in cols_num:
        if c in df.columns:
            s = df[c].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.strip()
            df[c] = pd.to_numeric(s, errors="coerce").fillna(0)
    return df

def guardar_hoja_sincronizada(df_guardar, nombre_hoja):
    gc, _ = obtener_cliente_gsheets()
    if gc:
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
        except Exception as e:
            st.warning(f"Aviso sincronizando nube: {e}")

@st.cache_data(ttl=30)
def descargar_hojas_nube():
    hojas = {}
    nombres_hojas = [
        "Mi_Cartera", "Historial_Tendencias", "Recomendacion_IA",
        "Dashboard_Financiero", "Radar_Fodder", "Arbitraje_FUTGG",
        "Radar_Evos", "Radar_TOTW", "Radar_Links_Domino", "Mis_Alertas"
    ]
    gc, err_auth = obtener_cliente_gsheets()
    if not gc:
        return {}, err_auth

    try:
        sh = gc.open(NOMBRE_SHEET)
        pestanas_reales = [w.title for w in sh.worksheets()]
        for h in nombres_hojas:
            if h in pestanas_reales:
                try:
                    ws = sh.worksheet(h)
                    df_h = leer_pestana_segura(ws)
                    hojas[h] = convertir_columnas_numericas(df_h)
                except Exception:
                    hojas[h] = pd.DataFrame()
            else:
                hojas[h] = pd.DataFrame()
        return hojas, "OK"
    except Exception as e:
        return {}, f"Error abriendo '{NOMBRE_SHEET}': {e}"

try:
    hojas_nube, msg_nube = descargar_hojas_nube()
    if msg_nube == "OK" and any(not df.empty for df in hojas_nube.values()):
        datos = hojas_nube
        estado_conexion = "☁️ Conectado a Google Sheets (Nube 24/7)"
    else:
        datos = {}
        estado_conexion = "⚠️ Sin conexión a datos"
        if msg_nube != "OK":
            st.error(f"🔍 Diagnóstico: {msg_nube}")

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

        st.markdown("---")
        st.markdown("### ➕ Registrar Compra Manual")
        with st.form("form_compra", clear_on_submit=True):
            jug_c = st.text_input("Nombre del Jugador", placeholder="Ej: Odegaard")
            precio_c = st.number_input("Precio Pagado (Monedas)", min_value=200, max_value=2000000, step=250, value=10000)
            est_c = st.selectbox("Unidad de Negocio", ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"])
            btn_guardar_c = st.form_submit_button("✅ Guardar en Nube")

            if btn_guardar_c and jug_c.strip():
                df_c_act = datos.get("Mi_Cartera", pd.DataFrame())
                nueva_fila = pd.DataFrame([{
                    "Jugador_Exacto": jug_c.strip().title(),
                    "Estrategia": est_c,
                    "Precio_Pagado": int(precio_c),
                    "Estado": "Activo",
                    "Precio_Venta": 0,
                    "Costo_EA_Tax": 0,
                    "Ganancia_Neta": 0,
                    "ROI_%": 0.0,
                    "Fecha_Compra": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Fecha_Venta": ""
                }])
                df_c_act = pd.concat([df_c_act, nueva_fila], ignore_index=True)
                guardar_hoja_sincronizada(df_c_act, "Mi_Cartera")
                st.cache_data.clear()
                st.success(f"¡{jug_c.title()} sincronizado en la nube!")
                st.rerun()

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
        st.warning("⚠️ Seguí los pasos del cartel rojo de arriba para conectar tus datos.")
        st.stop()

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
            st.dataframe(df_ult_corte.head(250), hide_index=True)

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
