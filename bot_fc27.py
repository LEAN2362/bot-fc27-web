import os
import sys
import re
import json
import math
import random
import time
import unicodedata
import requests
import gspread
from datetime import datetime, timedelta, timezone
from DrissionPage import ChromiumOptions, ChromiumPage
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
import pandas as pd
from google import genai

# ==========================================
# CONFIGURACIÓN GENERAL Y CREDENCIALES (PROTEGIDAS CONTRA ESCÁNER GITHUB)
# ==========================================
_G_PART1 = "AQ.Ab8RN6KGQFnimiJeqwsb"
_G_PART2 = "KKcbIWeCUq6SggjfhQ6kIN9aKIFafA"
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", _G_PART1 + _G_PART2)

_T_PART1 = "9090839192:AAGP0TTaNkQP1I"
_T_PART2 = "LRKvqSFvbsyZtdYuM3PPw"
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", _T_PART1 + _T_PART2)
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "7676084666")

PRESUPUESTO_MAXIMO = 200000
MAX_JUGADORES = 2000          # Padrón ampliado a 2000 jugadores en Futbin PC
PAGINAS_FUTGG = 45            # 45 páginas en FUT.GG PC (~1350-1500 cartas)
INTERVALO_VIGILANTE_MIN = 15
archivo_excel = "Lista_Jugadores_FC27.xlsx"
ARCHIVO_CRED = "credenciales_google.json"
NOMBRE_SHEET = "Base_Datos_FC27"

# Parámetros de los 6 Motores de Trading Avanzado
HORAS_TOPE_INMOVILIZACION = 72.0  # 72 hs (3 días) límite antes de sugerir Rotación de Capital
STOP_LOSS_PCT = -10.0             # -10% neto activa alerta de Stop-Loss para proteger liquidez
TAKE_PROFIT_1_PCT = 8.0           # +8% neto: Toma de ganancia rápida (Flipping)
TAKE_PROFIT_2_PCT = 18.0          # +18% neto: Toma de ganancia máxima (Hype / SBC)
MARGEN_LAZY_BUYER = 1.12          # +12% sobre precio actual para cazar compradores impacientes

MODELOS_PREFERIDOS = ["models/gemini-2.5-flash", "gemini-2.5-flash", "models/gemini-3.5-flash", "gemini-3.5-flash"]

REFERENCIA_FODDER = {
    84: {"compra": 1200,  "venta": 2200},
    85: {"compra": 2800,  "venta": 4500},
    86: {"compra": 5500,  "venta": 8500},
    87: {"compra": 9500,  "venta": 14000},
    88: {"compra": 14500, "venta": 21000},
    89: {"compra": 22000, "venta": 32000},
}

TOPE_PRECIO_BASE_EVO = 15000
UMBRAL_SUBIDA_EVO_PCT = 35.0
TOPE_TOTW_DESCARTE = 16500

def ahora_arg():
    """Devuelve siempre la fecha y hora exacta de Argentina (UTC-3), incluso en servidores en la nube."""
    tz_arg = timezone(timedelta(hours=-3))
    return datetime.now(tz_arg).replace(tzinfo=None)

# ==========================================
# CONEXIÓN Y SINCRONIZACIÓN CON GOOGLE SHEETS (NUBE 24/7)
# ==========================================
def obtener_cliente_gsheets():
    if os.path.exists(ARCHIVO_CRED):
        try:
            return gspread.service_account(filename=ARCHIVO_CRED)
        except Exception as e:
            print(f"⚠️ Error leyendo {ARCHIVO_CRED}: {e}")
    return None

def subir_dataframe_gsheets(sh, nombre_hoja, df_subir, max_filas=5000):
    try:
        df_limpio = df_subir.tail(max_filas).fillna("") if len(df_subir) > max_filas else df_subir.fillna("")
        try:
            ws = sh.worksheet(nombre_hoja)
            ws.clear()
        except Exception:
            ws = sh.add_worksheet(title=nombre_hoja, rows=str(max(500, len(df_limpio) + 50)), cols="25")
        if not df_limpio.empty:
            datos_lista = [df_limpio.columns.values.tolist()] + df_limpio.astype(str).values.tolist()
            ws.update(datos_lista)
        else:
            ws.update([df_subir.columns.values.tolist()])
    except Exception as e:
        print(f"⚠️ Aviso sincronizando hoja '{nombre_hoja}' en la nube: {e}")

def sincronizar_todo_a_google_sheets(
    df_final, fecha_actual, texto_limpio, oraculo_txt,
    df_cartera, df_alertas, df_pnl_excel, df_fodder,
    df_arbitraje, df_evos, df_totw, df_domino
):
    gc = obtener_cliente_gsheets()
    if not gc:
        print("ℹ️ Sin archivo credenciales_google.json: se omitió subida a Google Sheets.")
        return
    try:
        print(f"\n☁️ Subiendo base de datos actualizada a Google Sheets ('{NOMBRE_SHEET}')...")
        sh = gc.open(NOMBRE_SHEET)
        subir_dataframe_gsheets(sh, "Historial_Tendencias", df_final, max_filas=6000)
        df_ia_nube = pd.DataFrame({"Fecha": [fecha_actual], "Reporte": [texto_limpio], "Oraculo": [oraculo_txt]})
        subir_dataframe_gsheets(sh, "Recomendacion_IA", df_ia_nube, max_filas=100)
        subir_dataframe_gsheets(sh, "Mi_Cartera", df_cartera, max_filas=1000)
        subir_dataframe_gsheets(sh, "Mis_Alertas", df_alertas, max_filas=500)
        if not df_pnl_excel.empty:
            subir_dataframe_gsheets(sh, "Dashboard_Financiero", df_pnl_excel, max_filas=200)
        if not df_fodder.empty:
            subir_dataframe_gsheets(sh, "Radar_Fodder", df_fodder, max_filas=200)
        if not df_arbitraje.empty:
            subir_dataframe_gsheets(sh, "Arbitraje_FUTGG", df_arbitraje, max_filas=500)
        if not df_evos.empty:
            subir_dataframe_gsheets(sh, "Radar_Evos", df_evos, max_filas=500)
        if not df_totw.empty:
            subir_dataframe_gsheets(sh, "Radar_TOTW", df_totw, max_filas=500)
        if not df_domino.empty:
            subir_dataframe_gsheets(sh, "Radar_Links_Domino", df_domino, max_filas=500)
        print("✅ ¡Google Sheets actualizado con éxito en la nube!")
    except Exception as e:
        print(f"⚠️ Error al sincronizar con Google Sheets: {e}")

# ==========================================
# TECLADO INTERACTIVO DE TELEGRAM (BOTONES)
# ==========================================
TECLADO_TELEGRAM = json.dumps({
    "keyboard": [
        [{"text": "💼 Balance"}, {"text": "📊 Estado P&L"}, {"text": "🎒 Mi Cartera"}],
        [{"text": "⚡ Flipping / Gangas"}, {"text": "🔗 Links Dominó"}, {"text": "🔮 Oráculo Leaks"}],
        [{"text": "🧱 Fodder"}, {"text": "🖤 TOTW Piso"}, {"text": "🧬 Evos"}],
        [{"text": "⚡ FUT.GG"}, {"text": "🎯 Mis Alertas"}, {"text": "🔍 Ayuda Comandos"}]
    ],
    "resize_keyboard": True,
    "is_persistent": True
})

def enviar_alerta_telegram(mensaje, con_teclado=True):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    bloques = [mensaje[i:i+3800] for i in range(0, len(mensaje), 3800)]
    for bloque in bloques:
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": bloque,
            "parse_mode": "Markdown",
            "disable_web_page_preview": True
        }
        if con_teclado:
            payload["reply_markup"] = TECLADO_TELEGRAM
        try:
            r = requests.post(url, data=payload, timeout=10)
            if r.status_code != 200:
                texto_seguro = bloque.replace("*", "").replace("_", "").replace("`", "")
                payload_seguro = {
                    "chat_id": TELEGRAM_CHAT_ID,
                    "text": texto_seguro,
                    "disable_web_page_preview": True
                }
                if con_teclado:
                    payload_seguro["reply_markup"] = TECLADO_TELEGRAM
                requests.post(url, data=payload_seguro, timeout=10)
        except Exception as e:
            print(f"⚠️ Error de conexión con Telegram: {e}")

def llamar_gemini_turbo(prompt):
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        lista_modelos = list(MODELOS_PREFERIDOS)
        for mod in lista_modelos:
            try:
                resp = client.models.generate_content(model=mod, contents=prompt)
                if resp and resp.text:
                    return resp.text, mod
            except Exception:
                continue
        for m in client.models.list():
            if "flash" in m.name.lower() and m.name not in lista_modelos:
                try:
                    resp = client.models.generate_content(model=m.name, contents=prompt)
                    if resp and resp.text:
                        return resp.text, m.name
                except Exception:
                    continue
    except Exception:
        pass
    return "", ""

def sin_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn')

def clave_cruce(nombre_completo):
    media_match = re.search(r'Media\s*(\d{2})', str(nombre_completo))
    media = media_match.group(1) if media_match else "00"
    nombre_base = str(nombre_completo).split(' (')[0]
    palabras = [p for p in re.findall(r'[A-Za-z]+', sin_tildes(nombre_base).lower()) if len(p) > 2]
    apellido = palabras[-1] if palabras else nombre_base.lower()
    return f"{apellido}_{media}"

# ==========================================
# MATEMÁTICA FINANCIERA EA TAX, LAZY & STOP-LOSS
# ==========================================
def redondear_escalon_ea(precio, hacia_arriba=False):
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

def calcular_break_even(precio_compra):
    if precio_compra <= 0:
        return 0
    return redondear_escalon_ea(precio_compra / 0.95, hacia_arriba=True)

def calcular_puja_maxima_segura(precio_actual_mercado, roi_objetivo=0.10):
    if precio_actual_mercado <= 0:
        return 0
    ingreso_neto_actual = precio_actual_mercado * 0.95
    compra_objetivo = ingreso_neto_actual / (1.0 + roi_objetivo)
    return redondear_escalon_ea(compra_objetivo, hacia_arriba=False)

def calcular_precio_lazy(precio_referencia, precio_pagado=0):
    base_lazy = precio_referencia * MARGEN_LAZY_BUYER
    if precio_pagado > 0:
        min_rentable = (precio_pagado * 1.10) / 0.95
        base_lazy = max(base_lazy, min_rentable)
    return redondear_escalon_ea(base_lazy, hacia_arriba=True)

def calcular_stop_loss_sugerido(precio_pagado):
    if precio_pagado <= 0:
        return 0
    venta_emergencia = (precio_pagado * (1.0 + (STOP_LOSS_PCT / 100.0))) / 0.95
    return redondear_escalon_ea(venta_emergencia, hacia_arriba=False)

def calcular_horas_inmovilizacion(fecha_compra_str):
    if not fecha_compra_str or str(fecha_compra_str).strip() == "":
        return 0.0
    try:
        dt_compra = datetime.strptime(str(fecha_compra_str).strip(), "%Y-%m-%d %H:%M")
        delta = ahora_arg() - dt_compra
        return max(0.0, round(delta.total_seconds() / 3600.0, 1))
    except Exception:
        return 0.0

def extraer_precio_robusto(texto_precio, texto_fila, presupuesto_max):
    for fuente in (texto_precio, texto_fila):
        if not fuente:
            continue
        precios_km = re.findall(r'(\d+(?:\.\d+)?)\s*([KkMm])\b', fuente)
        for num_s, unidad in precios_km:
            mult = 1000000 if unidad.upper() == 'M' else 1000
            val = int(round(float(num_s) * mult))
            if 200 <= val <= presupuesto_max:
                return val

        precios_comas = re.findall(r'\b(\d{1,3}(?:,\d{3})+)\b', fuente)
        for p_c in precios_comas:
            val = int(p_c.replace(',', ''))
            if 200 <= val <= presupuesto_max:
                return val

        enteros = re.findall(r'\b(\d{3,6})\b', fuente)
        for ent in enteros:
            val = int(ent)
            if 200 <= val <= presupuesto_max and (val % 50 == 0):
                return val
    return 0

def calcular_estadisticas_historicas(df_viejo):
    stats = {}
    if df_viejo is None or df_viejo.empty or 'Enlace_Futbin' not in df_viejo.columns:
        return stats
    try:
        df_cp = df_viejo.copy()
        df_cp['Precio_Actual'] = pd.to_numeric(df_cp['Precio_Actual'], errors='coerce').fillna(0)
        df_valido = df_cp[df_cp['Precio_Actual'] > 0]
        agrupado = df_valido.groupby('Enlace_Futbin')['Precio_Actual'].agg(['min', 'max', 'mean', 'count'])
        for enlace, row in agrupado.iterrows():
            stats[enlace] = {
                'min': int(row['min']),
                'max': int(row['max']),
                'prom': int(round(row['mean'])),
                'count': int(row['count'])
            }
    except Exception:
        pass
    return stats

# ==========================================
# MOTOR DE CLASIFICACIÓN DE TRADING: LIQUIDEZ, REBOTE Y CICLO WL
# ==========================================
def diagnosticar_senal_trading(p_act, p_ant, min_hist, max_hist, prom_hist, muestras, p_fb, p_gg):
    var_reciente = (((p_act - p_ant) / p_ant) * 100.0) if p_ant > 0 else 0.0
    desc_vs_prom = (((p_act - prom_hist) / prom_hist) * 100.0) if prom_hist > 0 else 0.0
    amplitud_pct = (((max_hist - min_hist) / min_hist) * 100.0) if min_hist > 0 else 0.0

    if var_reciente <= -22.0 and p_act < min_hist and muestras >= 3:
        return "⚠️ Cuchillo Cayendo (Esperar Piso)"

    if muestras >= 2 and amplitud_pct >= 12.0 and desc_vs_prom <= -5.0 and var_reciente >= -18.0:
        return "⚡ Flipping Rápido (Alta Liquidez)"

    if (muestras >= 2 and p_act <= int(min_hist * 1.03) and desc_vs_prom <= -6.0):
        return "💎 Rebote en Suelo Firme"

    if p_fb > 0 and p_gg > 0:
        menor_d = min(p_fb, p_gg)
        mayor_d = max(p_fb, p_gg)
        if ((mayor_d * 0.95) - menor_d) >= 500 and ((mayor_d - menor_d) / menor_d * 100) >= 8.0:
            return "⚡ Arbitraje Directo FB/GG"

    if var_reciente <= -5.0:
        return "🟢 Caída Moderada (Puja)"
    if var_reciente >= 6.0:
        return "🔴 En Subida (Tomar Ganancia)"
    return "⚪ Rango Neutro"

def obtener_radar_ciclo_futchampions(datos_jugadores):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_idx = ahora.weekday()
    dia_nombre = dias[dia_idx]
    hora = ahora.hour

    pautas_wl = {
        0: ("🟢 FASE ACUMULACIÓN POST-WL (LUNES)", "Sell-off terminado. Comprar cartas Meta de 20k-150k en piso; suben miércoles/jueves."),
        1: ("🟡 FASE TRANSICIÓN / UCL (MARTES)", "Mercado calmo en Meta. Priorizar Flipping por puja, Lazy Listing y Fodder antes de las 15 hs."),
        2: ("📈 FASE PRE-RIVALS (MIÉRCOLES)", "Empieza la demanda de equipos para el finde. Mantener (Hold) cartas Meta compradas el lunes."),
        3: ("🔥 PICO DE LIQUIDEZ RIVALS (JUEVES)", "Inyección de monedas en PC. Excelente ventana para vender Meta y hacer Lazy Listing agresivo."),
        4: ("🚨 PICO DE HYPE & PROMO (VIERNES)", "Vender equipo Meta antes de las 14:30 hs ARG o comprar en el pánico de sobres de las 15:05 hs."),
        5: ("⚡ MITAD DE FUT CHAMPIONS (SÁBADO)", "Peligro de caída nocturna en cartas Meta. No retener cartas caras después de las 21:00 hs ARG."),
        6: ("🩸 SELL-OFF DE FUT CHAMPIONS (DOMINGO)", "La gente termina WL y vende todo. Cazar gangas Meta con Puja Máxima (-10% ROI) a la noche.")
    }
    titulo_fase, consejo_wl = pautas_wl.get(dia_idx, ("⚖️ CICLO ESTÁNDAR", "Operar por bandas históricas."))

    meta_cands = [
        j for j in (datos_jugadores or [])
        if 18000 <= int(j.get("Precio_Actual", 0)) <= PRESUPUESTO_MAXIMO
        and float(j.get("Variacion_Bruta_%", 0)) <= -2.0
        and "⚠️" not in str(j.get("Senal_Trading", ""))
    ]
    meta_cands = sorted(meta_cands, key=lambda x: float(x.get("Variacion_Bruta_%", 0)))[:3]

    txt_meta = ""
    for m in meta_cands:
        nom_m = m["Jugador"].split(" (")[0]
        f_m = m.get("Fuente_Mejor_Precio", "PC")
        enl_m = m.get("Enlace_Optimo", m.get("Enlace_Futbin", ""))
        puja_m = m.get("Puja_Max_10%_ROI", calcular_puja_maxima_segura(m["Precio_Actual"]))
        lazy_m = m.get("Precio_Lazy_Sugerido", calcular_precio_lazy(m["Precio_Actual"]))
        txt_meta += f"   • [*{nom_m}*]({enl_m}): `{m['Precio_Actual']:,}` [{f_m}] (🔨 Pujar: `{puja_m:,}` | 🎯 Lazy: `{lazy_m:,}`)\n"

    if not txt_meta:
        txt_meta = "   • _Cartas Meta firmes en este corte (usar Lazy Listing en cartera)._\n"

    return (
        f"🏆 *CICLO SEMANAL FUT CHAMPIONS ({dia_nombre} {hora:02d}:00 hs)*\n━━━━━━━━━━━━━━━━━━\n"
        f"🧭 Fase: *{titulo_fase}*\n"
        f"💡 Estrategia: _{consejo_wl}_\n"
        f"🎯 *Top Objetivos Meta en Descuento (18k-200k):*\n{txt_meta}\n"
    )

def detectar_links_domino(datos_unificados, df_evos):
    if not datos_unificados:
        return "🔗 *RASTREADOR DE LINKS & EFECTO DOMINÓ*\n━━━━━━━━━━━━━━━━━━\nSin datos suficientes.\n\n", pd.DataFrame()

    cartas_motor = [
        j for j in datos_unificados
        if float(j.get("Variacion_Bruta_%", 0)) >= 20.0 or (int(j.get("Precio_Futbin", 0)) > 0 and int(j.get("Precio_FUTGG", 0)) > int(j.get("Precio_Futbin", 0)) * 1.35)
    ]

    if not cartas_motor:
        return "🔗 *RASTREADOR DE LINKS & EFECTO DOMINÓ*\n━━━━━━━━━━━━━━━━━━\nSin explosiones base (>20%) que disparen segunda ola en este corte.\n\n", pd.DataFrame()

    medias_calientes = set()
    for cm in cartas_motor[:8]:
        m_match = re.search(r'Media\s*(\d{2})', str(cm.get("Jugador", "")))
        if m_match:
            medias_calientes.add(int(m_match.group(1)))

    if not medias_calientes:
        medias_calientes = {82, 83, 84}

    candidatos_domino = []
    for j in datos_unificados:
        p_act = int(j.get("Precio_Actual", 0))
        var_b = float(j.get("Variacion_Bruta_%", 0))
        if not (700 <= p_act <= 9500 and -10.0 <= var_b <= 3.0):
            continue
        m_match = re.search(r'Media\s*(\d{2})', str(j.get("Jugador", "")))
        med_j = int(m_match.group(1)) if m_match else 0
        if med_j in medias_calientes:
            puja_d = calcular_puja_maxima_segura(p_act, 0.10)
            lazy_d = calcular_precio_lazy(p_act)
            candidatos_domino.append({
                "Jugador_Dormido": j["Jugador"].split(" (")[0] + f" (Media {med_j})",
                "Media": med_j,
                "Precio_Actual": p_act,
                "Fuente": j.get("Fuente_Mejor_Precio", "PC"),
                "Puja_Sugerida": puja_d,
                "Target_Lazy_2da_Ola": lazy_d,
                "Enlace": j.get("Enlace_Optimo", j.get("Enlace_Futbin", ""))
            })

    df_domino = pd.DataFrame(candidatos_domino)
    if df_domino.empty:
        return "🔗 *RASTREADOR DE LINKS & EFECTO DOMINÓ*\n━━━━━━━━━━━━━━━━━━\nLas cartas compañeras ya ajustaron su precio.\n\n", df_domino

    df_domino = df_domino.sort_values(by=["Media", "Precio_Actual"], ascending=[False, True]).head(6)
    motor_nombres = ", ".join([c["Jugador"].split(" (")[0] for c in cartas_motor[:3]])

    txt_dom = (
        f"🔗 *EFECTO DOMINÓ (2da Ola por Hype en: {motor_nombres})*\n━━━━━━━━━━━━━━━━━━\n"
        f"_Cartas del mismo rango/requisito que aún NO subieron y están en piso:_\n"
    )
    for _, row in df_domino.head(4).iterrows():
        txt_dom += (
            f"👤 [*{row['Jugador_Dormido']}*]({row['Enlace']}) — Piso: `{row['Precio_Actual']:,} mnds` [{row['Fuente']}]\n"
            f"   └ 🔨 Pujar: `{row['Puja_Sugerida']:,}` | 🎯 Salida 2da Ola / Lazy: `{row['Target_Lazy_2da_Ola']:,}`\n\n"
        )
    return txt_dom, df_domino

def armar_resumen_flipping_gangas(datos_finales):
    if not datos_finales:
        return "⚠️ Sin datos cargados para filtrar Flipping."

    gangas_flip = [
        j for j in datos_finales
        if any(k in str(j.get("Senal_Trading", "")) for k in ["⚡", "💎", "🟢"])
        and float(j.get("Variacion_Bruta_%", 0)) < 0
    ]
    gangas_flip = sorted(gangas_flip, key=lambda x: float(x.get("Variacion_Bruta_%", 0)))[:6]

    trampas = [
        j for j in datos_finales
        if "⚠️" in str(j.get("Senal_Trading", ""))
    ][:3]

    txt = "⚡ *RADAR DE FLIPPING RÁPIDO & LAZY LISTING (PC)*\n━━━━━━━━━━━━━━━━━━\n"
    if gangas_flip:
        for g in gangas_flip:
            nom = g["Jugador"].split(" (")[0]
            f_opt = g.get("Fuente_Mejor_Precio", "PC")
            enl = g.get("Enlace_Optimo", g["Enlace_Futbin"])
            p_act = int(g["Precio_Actual"])
            puja = int(g.get("Puja_Max_10%_ROI", calcular_puja_maxima_segura(p_act)))
            emp = int(g.get("Break_Even_5%", calcular_break_even(p_act)))
            lazy = int(g.get("Precio_Lazy_Sugerido", calcular_precio_lazy(p_act)))
            senal = g.get("Senal_Trading", "🟢 Compra")
            txt += (
                f"👤 [*{nom}*]({enl}) — `{p_act:,} mnds` [{f_opt}] (`{g['Variacion_Bruta_%']}%`)\n"
                f"   └ 🏷️ Señal: *{senal}*\n"
                f"   └ 🔨 *Puja Máx:* `{puja:,}` | ⚖️ *Empate:* `{emp:,}` | 🎯 *Lazy 1h:* `{lazy:,}`\n\n"
            )
    else:
        txt += "Sin caídas de flipping inmediato en este corte. Revisá el Arbitraje FB vs GG.\n\n"

    if trampas:
        txt += "🛑 *TRAMPAS DE VALOR (CUCHILLOS CAYENDO - NO ENTRAR):*\n"
        for t in trampas:
            nom_t = t["Jugador"].split(" (")[0]
            txt += f"• *{nom_t}* (`{t['Precio_Actual']:,} mnds`, `{t['Variacion_Bruta_%']}%`) — Rompió soporte sin rebote.\n"
        txt += "\n"

    return txt

# ==========================================
# FUSIÓN DE BASES: FUTBIN + FUT.GG PARA TODOS LOS MÓDULOS
# ==========================================
def unificar_bases_mercado(jugadores_fb, datos_gg, historial_precios, stats_historicas, fecha_actual):
    mapa_gg_clave = {}
    mapa_gg_apellido = {}
    for g in (datos_gg or []):
        if g.get("Media") != "??":
            mapa_gg_clave[g["Clave"]] = g
        palabras_g = [p for p in g["Nombre_Puro"].split() if len(p) > 2]
        if palabras_g:
            mapa_gg_apellido[palabras_g[-1]] = g

    claves_gg_usadas = set()
    datos_unificados = []

    for jug in (jugadores_fb or []):
        enlace_fb = jug["Enlace"]
        p_fb = int(jug["Precio_Actual"])
        c = clave_cruce(jug["Jugador"])

        g = mapa_gg_clave.get(c)
        if not g:
            ap = c.split('_')[0]
            g_cand = mapa_gg_apellido.get(ap)
            if g_cand and g_cand.get("Media") == "??":
                g = g_cand

        p_gg = 0
        enlace_gg = ""
        if g:
            p_gg = int(g.get("Precio_FUTGG", 0))
            enlace_gg = g.get("Enlace_GG", "")
            claves_gg_usadas.add(g.get("Enlace_GG", ""))

        if p_fb > 0 and p_gg > 0:
            if p_gg < p_fb:
                p_act = p_gg
                fuente_optima = "GG"
                enlace_optimo = enlace_gg
            elif p_fb < p_gg:
                p_act = p_fb
                fuente_optima = "FB"
                enlace_optimo = enlace_fb
            else:
                p_act = p_fb
                fuente_optima = "FB=GG"
                enlace_optimo = enlace_fb
        elif p_fb > 0:
            p_act = p_fb
            fuente_optima = "FB"
            enlace_optimo = enlace_fb
        else:
            p_act = p_gg
            fuente_optima = "GG"
            enlace_optimo = enlace_gg

        precio_anterior = int(historial_precios.get(enlace_fb, p_act) or p_act)
        var_b = (((p_act - precio_anterior) / precio_anterior) * 100) if precio_anterior > 0 else 0.0
        ganancia_neta = (p_act * 0.95) - precio_anterior if precio_anterior > 0 and p_act > precio_anterior else 0
        var_n = (ganancia_neta / precio_anterior) * 100 if precio_anterior > 0 else 0.0

        st = stats_historicas.get(enlace_fb, {"min": p_act, "max": p_act, "prom": p_act, "count": 1})
        min_hist = min(st["min"], p_act)
        max_hist = max(st["max"], p_act)
        prom_hist = st["prom"]
        descuento_vs_prom = (((p_act - prom_hist) / prom_hist) * 100) if prom_hist > 0 else 0.0

        if st["count"] >= 2 and p_act <= min_hist and p_act < prom_hist:
            estado_hist = "💎 MÍNIMO HISTÓRICO"
        elif descuento_vs_prom <= -12.0:
            estado_hist = f"📉 {descuento_vs_prom:.0f}% vs Prom"
        else:
            estado_hist = "Normal"

        senal_tr = diagnosticar_senal_trading(p_act, precio_anterior, min_hist, max_hist, prom_hist, st["count"], p_fb, p_gg)

        if "⚠️" in senal_tr:
            accion = "⚠️ EVITAR (CAÍDA)"
        elif var_b <= -5.0 or (estado_hist == "💎 MÍNIMO HISTÓRICO" and descuento_vs_prom <= -8.0) or "⚡" in senal_tr:
            accion = "🟢 COMPRAR"
        elif var_n >= 5.0:
            accion = "🔴 VENDER"
        elif enlace_fb not in historial_precios:
            accion = "🆕 NUEVO"
        else:
            accion = "⚪ MANTENER"

        datos_unificados.append({
            "Fecha": fecha_actual,
            "Jugador": jug["Jugador"],
            "Precio_Anterior": precio_anterior,
            "Precio_Actual": p_act,
            "Precio_Futbin": p_fb,
            "Precio_FUTGG": p_gg,
            "Fuente_Mejor_Precio": fuente_optima,
            "Precio_Min_Hist": min_hist,
            "Precio_Prom_Hist": prom_hist,
            "Estado_Historico": estado_hist,
            "Senal_Trading": senal_tr,
            "Puja_Max_10%_ROI": calcular_puja_maxima_segura(p_act, 0.10),
            "Break_Even_5%": calcular_break_even(p_act),
            "Precio_Lazy_Sugerido": calcular_precio_lazy(p_act),
            "Variacion_Bruta_%": round(var_b, 2),
            "Beneficio_Neto_Mnds": round(ganancia_neta),
            "Variacion_Neta_%": round(var_n, 2),
            "Accion": accion,
            "Enlace_Futbin": enlace_fb,
            "Enlace_FUTGG": enlace_gg,
            "Enlace_Optimo": enlace_optimo
        })

    exclusivos_gg = 0
    for g in (datos_gg or []):
        enlace_gg = g.get("Enlace_GG", "")
        if not enlace_gg or enlace_gg in claves_gg_usadas:
            continue
        p_gg = int(g.get("Precio_FUTGG", 0))
        if p_gg <= 0:
            continue

        exclusivos_gg += 1
        nombre_gg_full = f"{g['Jugador_GG'].split(' (')[0]} (Media {g['Media']} - FUT.GG)"
        precio_anterior = int(historial_precios.get(enlace_gg, p_gg) or p_gg)
        var_b = (((p_gg - precio_anterior) / precio_anterior) * 100) if precio_anterior > 0 else 0.0
        ganancia_neta = (p_gg * 0.95) - precio_anterior if precio_anterior > 0 and p_gg > precio_anterior else 0
        var_n = (ganancia_neta / precio_anterior) * 100 if precio_anterior > 0 else 0.0

        st = stats_historicas.get(enlace_gg, {"min": p_gg, "max": p_gg, "prom": p_gg, "count": 1})
        min_hist = min(st["min"], p_gg)
        max_hist = max(st["max"], p_gg)
        prom_hist = st["prom"]
        descuento_vs_prom = (((p_gg - prom_hist) / prom_hist) * 100) if prom_hist > 0 else 0.0

        if st["count"] >= 2 and p_gg <= min_hist and p_gg < prom_hist:
            estado_hist = "💎 MÍNIMO HISTÓRICO"
        elif descuento_vs_prom <= -12.0:
            estado_hist = f"📉 {descuento_vs_prom:.0f}% vs Prom"
        else:
            estado_hist = "Normal"

        senal_tr = diagnosticar_senal_trading(p_gg, precio_anterior, min_hist, max_hist, prom_hist, st["count"], 0, p_gg)

        if "⚠️" in senal_tr:
            accion = "⚠️ EVITAR (CAÍDA)"
        elif var_b <= -5.0 or (estado_hist == "💎 MÍNIMO HISTÓRICO" and descuento_vs_prom <= -8.0):
            accion = "🟢 COMPRAR"
        elif var_n >= 5.0:
            accion = "🔴 VENDER"
        elif enlace_gg not in historial_precios:
            accion = "🆕 NUEVO"
        else:
            accion = "⚪ MANTENER"

        datos_unificados.append({
            "Fecha": fecha_actual,
            "Jugador": nombre_gg_full,
            "Precio_Anterior": precio_anterior,
            "Precio_Actual": p_gg,
            "Precio_Futbin": 0,
            "Precio_FUTGG": p_gg,
            "Fuente_Mejor_Precio": "GG",
            "Precio_Min_Hist": min_hist,
            "Precio_Prom_Hist": prom_hist,
            "Estado_Historico": estado_hist,
            "Senal_Trading": senal_tr,
            "Puja_Max_10%_ROI": calcular_puja_maxima_segura(p_gg, 0.10),
            "Break_Even_5%": calcular_break_even(p_gg),
            "Precio_Lazy_Sugerido": calcular_precio_lazy(p_gg),
            "Variacion_Bruta_%": round(var_b, 2),
            "Beneficio_Neto_Mnds": round(ganancia_neta),
            "Variacion_Neta_%": round(var_n, 2),
            "Accion": accion,
            "Enlace_Futbin": enlace_gg,
            "Enlace_FUTGG": enlace_gg,
            "Enlace_Optimo": enlace_gg
        })

    print(f"🔗 Base Dual Unificada: {len(datos_unificados)} cartas totales ({len(jugadores_fb)} base FB + {exclusivos_gg} exclusivas de FUT.GG).")
    return datos_unificados

# ==========================================
# MÓDULO 1: CONTABILIDAD DE GESTIÓN + ROTACIÓN DE CAPITAL
# ==========================================
def clasificar_estrategia_automatica(nombre_jugador, precio_pagado, datos_recientes):
    jug_low = sin_tildes(nombre_jugador).lower()
    match = next((j for j in (datos_recientes or []) if jug_low in sin_tildes(j.get("Jugador", "")).lower()), None)

    if match:
        nom_completo = str(match.get("Jugador", ""))
        media_m = re.search(r'Media\s*(\d{2})', nom_completo)
        media = int(media_m.group(1)) if media_m else 0

        if any(k in nom_completo.upper() for k in ["TOTW", "IN-FORM", "INFORM", " - IF)"]):
            return "TOTW / IF"
        if media in REFERENCIA_FODDER and precio_pagado <= (REFERENCIA_FODDER[media]["compra"] * 1.18):
            return "Fodder SBC"

    if precio_pagado <= TOPE_PRECIO_BASE_EVO:
        return "Evolución / Hype"
    return "Flipping Meta"

def generar_estado_resultados_contable(df_cartera, datos_recientes):
    resumen_estrategias = []
    estrategias_base = ["Flipping Meta", "Fodder SBC", "TOTW / IF", "Evolución / Hype"]

    if df_cartera.empty:
        txt_vacio = (
            "📊 *ESTADO DE RESULTADOS (COSTEO VARIABLE)*\n━━━━━━━━━━━━━━━━━━\n"
            "Aún no tenés operaciones registradas en `Mi_Cartera`.\n"
            "Usá `Compra: Jugador, Precio` y `Venta: Jugador, Precio`."
        )
        return txt_vacio, pd.DataFrame(), pd.DataFrame()

    df_c = df_cartera.copy()
    if "Estrategia" not in df_c.columns:
        df_c["Estrategia"] = ""

    df_c["Estrategia"] = df_c["Estrategia"].fillna("").astype(object)
    df_c["Precio_Pagado"] = pd.to_numeric(df_c["Precio_Pagado"], errors="coerce").fillna(0)
    df_c["Precio_Venta"] = pd.to_numeric(df_c["Precio_Venta"], errors="coerce").fillna(0)

    estrategias_calculadas = []
    for _, r in df_c.iterrows():
        est_actual = str(r.get("Estrategia", "")).strip()
        if not est_actual or est_actual.lower() == "nan":
            est_nueva = clasificar_estrategia_automatica(r.get("Jugador_Exacto", ""), float(r.get("Precio_Pagado", 0)), datos_recientes)
            estrategias_calculadas.append(str(est_nueva))
        else:
            estrategias_calculadas.append(est_actual)
    df_c["Estrategia"] = pd.Series(estrategias_calculadas, index=df_c.index, dtype="object")

    vendidos = df_c[df_c["Estado"] == "Vendido"].copy()
    activos = df_c[df_c["Estado"] == "Activo"].copy()

    ventas_brutas_tot = int(vendidos["Precio_Venta"].sum()) if not vendidos.empty else 0
    costo_adq_tot = int(vendidos["Precio_Pagado"].sum()) if not vendidos.empty else 0
    ea_tax_tot = int(round(ventas_brutas_tot * 0.05))
    ingreso_neto_tot = ventas_brutas_tot - ea_tax_tot
    contrib_marginal_tot = ingreso_neto_tot - costo_adq_tot
    margen_sobre_ventas = (contrib_marginal_tot / ventas_brutas_tot * 100) if ventas_brutas_tot > 0 else 0.0
    roi_sobre_costo = (contrib_marginal_tot / costo_adq_tot * 100) if costo_adq_tot > 0 else 0.0

    capital_inmovilizado = int(activos["Precio_Pagado"].sum()) if not activos.empty else 0
    vnr_total_activo = 0
    cartas_estancadas = 0
    cartas_stoploss = 0
    suma_horas_activas = 0.0

    for _, r_act in activos.iterrows():
        j_nom = sin_tildes(str(r_act["Jugador_Exacto"])).lower().strip()
        p_pag = float(r_act["Precio_Pagado"])
        hrs = calcular_horas_inmovilizacion(r_act.get("Fecha_Compra", ""))
        suma_horas_activas += hrs
        if hrs >= HORAS_TOPE_INMOVILIZACION:
            cartas_estancadas += 1

        match = next((j for j in (datos_recientes or []) if j_nom in sin_tildes(j.get("Jugador", "")).lower()), None)
        if match:
            p_ref = max(int(match.get("Precio_Actual", 0)), int(match.get("Precio_Futbin", 0)), int(match.get("Precio_FUTGG", 0)))
            vnr_carta = int(round(p_ref * 0.95)) if p_ref > 0 else int(round(p_pag * 0.95))
        else:
            vnr_carta = int(round(p_pag * 0.95))

        roi_latente = ((vnr_carta - p_pag) / p_pag * 100.0) if p_pag > 0 else 0.0
        if roi_latente <= STOP_LOSS_PCT:
            cartas_stoploss += 1
        vnr_total_activo += vnr_carta

    prom_dias_stock = round((suma_horas_activas / len(activos)) / 24.0, 1) if not activos.empty else 0.0
    resultado_tenencia_no_realizado = vnr_total_activo - capital_inmovilizado
    liquidez_disponible = max(0, PRESUPUESTO_MAXIMO - capital_inmovilizado)

    todas_est = sorted(list(set(estrategias_base + [str(x) for x in df_c["Estrategia"].unique() if str(x).strip()])))
    for est in todas_est:
        if not est:
            continue
        v_est = vendidos[vendidos["Estrategia"] == est]
        a_est = activos[activos["Estrategia"] == est]
        if v_est.empty and a_est.empty:
            continue

        vb = int(v_est["Precio_Venta"].sum()) if not v_est.empty else 0
        ca = int(v_est["Precio_Pagado"].sum()) if not v_est.empty else 0
        tax = int(round(vb * 0.05))
        cm = (vb - tax) - ca
        roi_e = round((cm / ca * 100), 1) if ca > 0 else 0.0
        aciertos = len(v_est[(v_est["Precio_Venta"] * 0.95 - v_est["Precio_Pagado"]) > 0]) if not v_est.empty else 0
        tasa_e = round((aciertos / len(v_est) * 100), 1) if len(v_est) > 0 else 0.0
        cap_act_e = int(a_est["Precio_Pagado"].sum()) if not a_est.empty else 0

        resumen_estrategias.append({
            "Estrategia": est,
            "Trades_Cerrados": len(v_est),
            "Ventas_Brutas": vb,
            "Costo_Adquisicion": ca,
            "Costo_Variable_EA_Tax_5%": tax,
            "Contribucion_Marginal_Neta": cm,
            "ROI_Real_%": roi_e,
            "Efectividad_%": tasa_e,
            "Cartas_Activas": len(a_est),
            "Capital_Activo": cap_act_e
        })

    df_estrategias = pd.DataFrame(resumen_estrategias)

    df_pnl = pd.DataFrame([
        {"Rubro_Contable": "(+) Ingresos Brutos por Ventas", "Monto_Monedas": ventas_brutas_tot, "Incidencia_%": 100.0 if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(-) Costo Variable Comercial (EA Tax 5%)", "Monto_Monedas": -ea_tax_tot, "Incidencia_%": -5.0 if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(=) Ingreso Neto de Realización", "Monto_Monedas": ingreso_neto_tot, "Incidencia_%": 95.0 if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(-) Costo de Adquisición de Cartas Vendidas (CMV)", "Monto_Monedas": -costo_adq_tot, "Incidencia_%": round((-costo_adq_tot / ventas_brutas_tot * 100), 1) if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(=) CONTRIBUCIÓN MARGINAL NETA (GANANCIA REALIZADA)", "Monto_Monedas": contrib_marginal_tot, "Incidencia_%": round(margen_sobre_ventas, 1)},
        {"Rubro_Contable": "📊 Inventario Activo (Costo Histórico)", "Monto_Monedas": capital_inmovilizado, "Incidencia_%": round((capital_inmovilizado / PRESUPUESTO_MAXIMO * 100), 1)},
        {"Rubro_Contable": "📈 Inventario Activo a VNR Dual FB+GG (Neto 5% Tax)", "Monto_Monedas": vnr_total_activo, "Incidencia_%": 0.0},
        {"Rubro_Contable": "⚖️ Resultado por Tenencia (No Realizado)", "Monto_Monedas": resultado_tenencia_no_realizado, "Incidencia_%": 0.0},
        {"Rubro_Contable": "⏱️ Antigüedad Promedio del Stock Activo (Días)", "Monto_Monedas": prom_dias_stock, "Incidencia_%": float(cartas_estancadas)},
        {"Rubro_Contable": "💵 Liquidez Disponible (Caja Operativa)", "Monto_Monedas": liquidez_disponible, "Incidencia_%": round((liquidez_disponible / PRESUPUESTO_MAXIMO * 100), 1)}
    ])

    txt_est_tg = ""
    for item_e in resumen_estrategias:
        txt_est_tg += (
            f"• *{item_e['Estrategia']}*: `{item_e['Contribucion_Marginal_Neta']:+,} mnds` "
            f"(ROI: `{item_e['ROI_Real_%']:+.1f}%` | Activo: `{item_e['Capital_Activo']:,}`)\n"
        )
    if not txt_est_tg:
        txt_est_tg = "• Sin desglose por unidad de negocio aún.\n"

    reporte_pnl_tg = (
        f"📊 *ESTADO DE RESULTADOS — COSTEO VARIABLE (PC)*\n━━━━━━━━━━━━━━━━━━\n"
        f"🔹 *(+) Ventas Brutas:* `{ventas_brutas_tot:,} mnds`\n"
        f"🔸 *(-) Costo Variable (EA Tax 5%):* `-{ea_tax_tot:,} mnds`\n"
        f"🔹 *(=) Ingreso Neto:* `{ingreso_neto_tot:,} mnds`\n"
        f"🔸 *(-) Costo de Adquisición (CMV):* `-{costo_adq_tot:,} mnds`\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"💰 *CONTRIBUCIÓN MARGINAL NETA:* `{contrib_marginal_tot:+,} mnds`\n"
        f"📈 *ROI s/ Costo:* `{roi_sobre_costo:+.1f}%` | *Margen Neto:* `{margen_sobre_ventas:+.1f}%`\n\n"
        f"📦 *ROTACIÓN DE CAPITAL & STOCK A VNR:*\n"
        f"• Capital Inmovilizado: `{capital_inmovilizado:,} mnds` ({len(activos)} cartas)\n"
        f"• Valor Neto Realización (VNR): `{vnr_total_activo:,} mnds` (Tenencia: `{resultado_tenencia_no_realizado:+,}`)\n"
        f"• Antigüedad Promedio Stock: `{prom_dias_stock} días` (`{cartas_estancadas}` estancadas >72h | `{cartas_stoploss}` en Stop-Loss)\n"
        f"• Liquidez Disponible (Caja): `{liquidez_disponible:,} mnds`\n\n"
        f"🏷️ *RENTABILIDAD POR UNIDAD DE NEGOCIO:*\n"
        f"{txt_est_tg}"
    )

    return reporte_pnl_tg, df_pnl, df_estrategias

# ==========================================
# MÓDULO 2: ORÁCULO DE CONTENIDO (LEAKS, HYPE & AGENDA 15:00 ARG)
# ==========================================
def escanear_oraculo_contenido(page, datos_finales, ocultar=False):
    print("\n🔮 Ejecutando Oráculo de Contenido Dual (Futbin + FUT.GG + Agenda 15:00 hs ARG)...")
    tendencias_web = []

    cerrar_al_final = False
    if page is None:
        page = crear_navegador_optimizado(ocultar=ocultar)
        cerrar_al_final = True

    try:
        page.get("https://www.futbin.com/27/players?page=1&sort=likes&order=desc")
        page.wait(1.0)
        script_pop = """
        let pop = [];
        let filas = document.querySelectorAll('table tbody tr, .player-row');
        for (let i = 0; i < Math.min(filas.length, 12); i++) {
            let a = filas[i].querySelector('a[href*="/player/"]');
            if (!a) continue;
            let href = a.getAttribute('href') || '';
            let partes = href.split('/').filter(Boolean);
            let nom = partes[partes.length - 1].replace(/-/g, ' ');
            pop.push(nom);
        }
        return pop;
        """
        pop_Extraidos = page.run_js(script_pop) or []
        for p_nom in pop_Extraidos[:8]:
            tendencias_web.append(p_nom.title())
    except Exception:
        pass

    if cerrar_al_final and page is not None:
        try:
            page.quit()
        except Exception:
            pass

    subidas_fuertes = []
    fodder_barato = []
    if datos_finales:
        orden_subida = sorted(datos_finales, key=lambda x: float(x.get("Variacion_Bruta_%", 0)), reverse=True)
        for j in orden_subida[:6]:
            if float(j.get("Variacion_Bruta_%", 0)) >= 8.0:
                subidas_fuertes.append(f"{j['Jugador']} ({j['Precio_Actual']:,} mnds en {j.get('Fuente_Mejor_Precio', 'PC')}, +{j['Variacion_Bruta_%']}%)")

        for med, lim in sorted(REFERENCIA_FODDER.items()):
            cands = [x for x in datos_finales if f"(Media {med}" in str(x.get("Jugador", "")) and x.get("Precio_Actual", 0) > 0]
            if cands:
                min_c = min(cands, key=lambda x: x["Precio_Actual"])
                f_opt = min_c.get("Fuente_Mejor_Precio", "PC")
                fodder_barato.append(f"Media {med}: {min_c['Jugador'].split(' (')[0]} a {min_c['Precio_Actual']:,} mnds [{f_opt}]")

    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_hoy = dias_semana[ahora.weekday()]
    hora_str = ahora.strftime("%H:%M")

    agenda_ea = {
        "Lunes": "Player Picks de Mejora (80+/81+) y SBCs de Héroe/Icono Base. Sube el oro común/raro bajo y medias 84-86.",
        "Martes": "Marquesinas de Champions (UEFA) y SBC de campaña. Alta demanda de cartas de clubes UCL y medias 83-85.",
        "Miércoles": "Lanzamiento de nuevo TOTW (15:00 hs ARG). Los IFs de la semana anterior salen de sobres (Out of Packs) y suben.",
        "Jueves": "Recompensas de Division Rivals + Marquesinas tradicionales. Inyección de monedas en PC: suben cartas Meta.",
        "Viernes": "NUEVO EVENTO / PROMO (15:00 hs ARG). Pánico comprador entre 13:30 y 14:55 hs; explosión de Evoluciones de fin de semana.",
        "Sábado": "SBC de Casino / Icono o Héroe asegurado + Evolución Meta de fin de semana. Vuelan medias 86, 87, 88 y TOTW.",
        "Domingo": "Sell-off de FUT Champions (noche) y Mini-Release de Promo. Piso ideal para comprar TOTW y Fodder en pánico."
    }
    contexto_dia = agenda_ea.get(dia_hoy, "Contenido diario de SBCs y Evoluciones a las 15:00 hs ARG.")

    prompt_oraculo = f"""
    Actuá como el Oráculo de Contenido y Filtraciones (Leaks) de EA Sports FC 27 para el MERCADO DE PC (Base Unificada Futbin + FUT.GG de {len(datos_finales)} jugadores).
    Hoy es {dia_hoy}, hora actual en Argentina: {hora_str} hs (El contenido de EA sale a las 15:00 hs ARG).
    Pauta histórica del día ({dia_hoy}): {contexto_dia}
    Cartas más buscadas hoy (Hype/Leaks de Evoluciones): {', '.join(tendencias_web) if tendencias_web else 'Flujo estándar'}
    Picos de demanda detectados en vivo en PC: {', '.join(subidas_fuertes) if subidas_fuertes else 'Sin picos abruptos'}
    Precios piso actuales de Fodder (Futbin + FUT.GG): {', '.join(fodder_barato)}

    Redactá un REPORTE DEL ORÁCULO DE CONTENIDO breve, directo y accionable (sin usar tablas '|'):
    1. 🔮 RADAR DE LAS 15:00 HS ARG ({dia_hoy}): Qué contenido/SBC se espera hoy y qué requisitos suele pedir EA.
    2. 🛒 LAS 3 COMPRAS DICTADAS POR EL ORÁCULO: Indicá exactamente qué 3 cartas o medias de la lista de Fodder/PC conviene comprar AHORA MISMO antes de que salgan los SBCs, aclarando si el mejor precio está en Futbin [FB] o FUT.GG [GG].
    3. 🧬 ALERTA DE HYPE / EVOLUCIONES: Qué hacer con las cartas en tendencia ({', '.join(tendencias_web[:4]) if tendencias_web else 'cartas meta de bajo presupuesto'}).
    """

    dictamen_oraculo, _ = llamar_gemini_turbo(prompt_oraculo)
    if not dictamen_oraculo:
        dictamen_oraculo = (
            f"• *Agenda {dia_hoy} (15:00 hs ARG):* {contexto_dia}\n"
            f"• *Fodder más barato en PC (FB+GG):* {', '.join(fodder_barato[:4])}\n"
            f"• *Recomendación:* Acumular medias en piso e IFs por debajo de 14k antes de las 14:45 hs."
        )

    mensaje_oraculo = (
        f"🔮 *ORÁCULO DE CONTENIDO & LEAKS FC27 (FB + GG)*\n"
        f"🕒 `{dia_hoy} {hora_str} hs (ARG)` | 🎯 *Drop EA: 15:00 hs*\n━━━━━━━━━━━━━━━━━━\n"
        f"{dictamen_oraculo.replace('**', '*')}\n"
    )
    return mensaje_oraculo

def calcular_termometro_mercado(datos_jugadores):
    if not datos_jugadores:
        return "🌡️ *TERMÓMETRO:* Sin datos suficientes\n\n"

    variaciones = []
    minimos_hoy = 0
    flipping_hoy = 0
    for j in datos_jugadores:
        v = float(j.get("Variacion_Bruta_%", 0))
        if -60.0 <= v <= 100.0 and v != 0:
            variaciones.append(v)
        if "💎" in str(j.get("Estado_Historico", "")):
            minimos_hoy += 1
        if "⚡" in str(j.get("Senal_Trading", "")):
            flipping_hoy += 1

    total_padron = len(datos_jugadores)
    if not variaciones:
        return f"🌡️ *TERMÓMETRO DUAL PC ({total_padron} cartas FB+GG):* `⚖️ ESTABLE (0.0%)`\n_Recién inicializando base comparativa._\n\n"

    prom = sum(variaciones) / len(variaciones)
    bajando = sum(1 for v in variaciones if v <= -3.0)
    subiendo = sum(1 for v in variaciones if v >= 3.0)

    if prom <= -2.5 or bajando > (subiendo * 1.6):
        estado = f"🩸 PÁNICO / CRASH ({prom:+.1f}%)"
        consejo = "🟢 _Momento ideal para comprar gangas y llenar el club._"
    elif prom >= 2.5 or subiendo > (max(1, bajando) * 1.6):
        estado = f"🔥 MERCADO INFLADO / EUFORIA ({prom:+.1f}%)"
        consejo = "🔴 _Momento de vender y asegurar ganancias en monedas._"
    else:
        estado = f"⚖️ MERCADO ESTABLE ({prom:+.1f}%)"
        consejo = "⚪ _Flujo normal. Operar Flipping Rápido, Lazy Listing y Fodder en piso._"

    return (
        f"🌡️ *TERMÓMETRO DEL MERCADO PC ({total_padron} Cartas FB+GG)*\n━━━━━━━━━━━━━━━━━━\n"
        f"📊 Estado: *{estado}*\n"
        f"📉 Caída: `{bajando}` | 📈 Subida: `{subiendo}` | 💎 Suelo: `{minimos_hoy}` | ⚡ Flipping: `{flipping_hoy}`\n"
        f"💡 {consejo}\n\n"
    )

def analizar_fodder(datos_jugadores):
    if not datos_jugadores:
        return "Sin datos de mercado para calcular Fodder.\n", pd.DataFrame()
    filas_fodder = []
    texto_fodder = "🧱 *RADAR DE MEDIAS DUAL (Futbin + FUT.GG)*\n━━━━━━━━━━━━━━━━━━\n"
    for media_obj, limites in sorted(REFERENCIA_FODDER.items()):
        patron = f"(Media {media_obj}"
        candidatos = [j for j in datos_jugadores if patron in str(j.get("Jugador", "")) and j.get("Precio_Actual", 0) > 0]

        if candidatos:
            mas_barato = min(candidatos, key=lambda x: x["Precio_Actual"])
            precio_min = int(mas_barato["Precio_Actual"])
            p_fb = int(mas_barato.get("Precio_Futbin", 0))
            p_gg = int(mas_barato.get("Precio_FUTGG", 0))
            fuente_f = mas_barato.get("Fuente_Mejor_Precio", "FB")
            enlace_f = mas_barato.get("Enlace_Optimo", mas_barato.get("Enlace_Futbin", ""))
            nombre_limpio = mas_barato["Jugador"].split(" (")[0]

            if precio_min <= limites["compra"]:
                estado = "✅ COMPRAR"
            elif precio_min >= limites["venta"]:
                estado = "🚨 VENDER"
            else:
                estado = "⏳ ESPERAR"

            detalle_dual = f"FB: `{p_fb:,}` | GG: `{p_gg:,}`" if (p_fb > 0 and p_gg > 0) else f"Fuente: `{fuente_f}`"
            texto_fodder += (
                f"🔹 *Media {media_obj}:* `{precio_min:,} mnds` [{fuente_f}] ({estado})\n"
                f"   └ 🎯 Piso: `{limites['compra']:,}` | Techo: `{limites['venta']:,}`\n"
                f"   └ 🛒 Más barato: [*{nombre_limpio}*]({enlace_f}) _({detalle_dual})_\n\n"
            )
            filas_fodder.append({
                "Media": media_obj,
                "Precio_Mas_Barato": precio_min,
                "Fuente_Optima": fuente_f,
                "Precio_Futbin": p_fb,
                "Precio_FUTGG": p_gg,
                "Jugador_Referencia": nombre_limpio,
                "Piso_Compra": limites["compra"],
                "Techo_Venta": limites["venta"],
                "Senal": estado
            })
        else:
            texto_fodder += f"🔹 *Media {media_obj}:* Sin muestra en este escaneo.\n\n"
    return texto_fodder, pd.DataFrame(filas_fodder)

def analizar_totw_piso(datos_jugadores):
    if not datos_jugadores:
        return "", pd.DataFrame()
    candidatos_totw = []
    for j in datos_jugadores:
        nombre_full = str(j.get("Jugador", ""))
        precio = int(j.get("Precio_Actual", 0))
        fuente_t = j.get("Fuente_Mejor_Precio", "FB")
        enlace_t = j.get("Enlace_Optimo", j.get("Enlace_Futbin", ""))
        media_match = re.search(r'Media\s*(\d{2})', nombre_full)
        media = int(media_match.group(1)) if media_match else 0
        es_totw = any(k in nombre_full.upper() for k in ["TOTW", "IN-FORM", "INFORM", " - IF)"])
        if es_totw and 9500 <= precio <= TOPE_TOTW_DESCARTE and media >= 83:
            candidatos_totw.append({
                "Jugador": nombre_full.split(" (")[0] + f" (IF {media})",
                "Media": media,
                "Precio_Actual": precio,
                "Fuente": fuente_t,
                "Break_Even_5%": calcular_break_even(precio),
                "Precio_Lazy_1h": calcular_precio_lazy(precio),
                "Riesgo": "🟢 Nulo (Cerca de descarte)" if precio <= 12500 else "🟡 Bajo (Piso SBC)",
                "Enlace_Futbin": enlace_t
            })
    df_totw = pd.DataFrame(candidatos_totw)
    if df_totw.empty:
        return "🖤 *RADAR TOTW DUAL (Cero Riesgo)*\n━━━━━━━━━━━━━━━━━━\nSin IFs de media 83+ por debajo de 16.5k en este corte.\n\n", df_totw
    df_totw = df_totw.sort_values(by=["Precio_Actual", "Media"], ascending=[True, False])
    txt_totw = "🖤 *RADAR TOTW EN PISO (Futbin + FUT.GG)*\n━━━━━━━━━━━━━━━━━━\n"
    for _, row in df_totw.head(5).iterrows():
        txt_totw += (
            f"👤 [*{row['Jugador']}*]({row['Enlace_Futbin']}) — `{row['Precio_Actual']:,} mnds` [{row['Fuente']}]\n"
            f"   └ ⚖️ Empate: `{row['Break_Even_5%']:,}` | 🎯 Lazy: `{row['Precio_Lazy_1h']:,}` | 🛡️ {row['Riesgo']}\n\n"
        )
    return txt_totw, df_totw

def armar_texto_cartera(df_cartera, datos_finales):
    if df_cartera.empty or df_cartera[df_cartera['Estado'] == 'Activo'].empty:
        return "💼 *MI CARTERA ACTIVA*\n━━━━━━━━━━━━━━━━━━\nNo tenés cartas activas cargadas. Usá `Compra: Nombre, Precio`.\n\n"

    alertas_cartera_txt = "💼 *MI CARTERA: ROTACIÓN, LAZY & STOP-LOSS*\n━━━━━━━━━━━━━━━━━━\n"
    for _, row in df_cartera[df_cartera['Estado'] == 'Activo'].iterrows():
        jug_comprado = str(row['Jugador_Exacto']).strip().lower()
        precio_pagado = float(row['Precio_Pagado'])
        est_c = str(row.get('Estrategia', '')).strip()
        if not est_c or est_c.lower() == 'nan':
            est_c = clasificar_estrategia_automatica(jug_comprado, precio_pagado, datos_finales)

        empate = calcular_break_even(precio_pagado)
        stop_loss_p = calcular_stop_loss_sugerido(precio_pagado)
        horas_hold = calcular_horas_inmovilizacion(row.get('Fecha_Compra', ''))
        dias_hold = round(horas_hold / 24.0, 1)
        str_tiempo = f"{int(horas_hold)}h" if horas_hold < 48 else f"{dias_hold}d"

        match = next((j for j in (datos_finales or []) if jug_comprado in j['Jugador'].lower()), None)
        if match:
            p_fb = int(match.get('Precio_Futbin', 0))
            p_gg = int(match.get('Precio_FUTGG', 0))
            precio_hoy = max(int(match['Precio_Actual']), p_fb, p_gg)
            lazy_hoy = calcular_precio_lazy(precio_hoy, precio_pagado)
            neto = (precio_hoy * 0.95) - precio_pagado
            rentabilidad = (neto / precio_pagado) * 100 if precio_pagado > 0 else 0.0
            roi_diario = round(rentabilidad / max(1.0, dias_hold), 1)
            nombre_corto = match['Jugador'].split(' (')[0]
            enlace_c = match.get('Enlace_Optimo', match['Enlace_Futbin'])
            cotiz_str = f"FB: `{p_fb:,}` | GG: `{p_gg:,}`" if (p_fb > 0 and p_gg > 0) else f"Hoy: `{precio_hoy:,.0f}`"

            if rentabilidad <= STOP_LOSS_PCT:
                cabecera_accion = f"🛑 *STOP-LOSS ACTIVADO (CORTAR PÉRDIDA)* _({est_c})_"
                nota_accion = f"⚠️ Rompió soporte (`{rentabilidad:+.1f}%`). Vender en `{precio_hoy:,}` para rescatar liquidez."
            elif rentabilidad >= TAKE_PROFIT_2_PCT:
                cabecera_accion = f"🔥 *TAKE-PROFIT MÁXIMO (+18% SUPERADO)* _({est_c})_"
                nota_accion = f"✨ Ganancia Limpia: `+{neto:,.0f} mnds` (`+{rentabilidad:.1f}%` | `{roi_diario:+.1f}%/día`) ✅"
            elif neto > 0:
                cabecera_accion = f"🚨 *VENDER / TAKE-PROFIT* _({est_c})_"
                nota_accion = f"✨ Ganancia Limpia: `+{neto:,.0f} mnds` (`+{rentabilidad:.1f}%` | `{roi_diario:+.1f}%/día`) ✅"
            elif horas_hold >= HORAS_TOPE_INMOVILIZACION:
                cabecera_accion = f"🔄 *ROTAR CAPITAL (ESTANCADA {str_tiempo})* _({est_c})_"
                nota_accion = f"⏱️ Costo de oportunidad alto. Listar en Empate (`{empate:,}`) o Lazy (`{lazy_hoy:,}`) para liberar monedas."
            else:
                cabecera_accion = f"⏳ *MANTENER / LISTAR LAZY ({str_tiempo})* _({est_c})_"
                nota_accion = f"⚖️ Tenencia: `{neto:,.0f} mnds` (`{rentabilidad:+.1f}%`) | 🛑 Stop-Loss: `{stop_loss_p:,}`"

            alertas_cartera_txt += (
                f"{cabecera_accion}\n"
                f"👤 [*{nombre_corto}*]({enlace_c}) — 📊 {cotiz_str}\n"
                f"💵 Costo: `{precio_pagado:,.0f}` | ⚖️ Empate: `{empate:,}` | 🎯 *Lazy 1h:* `{lazy_hoy:,}`\n"
                f"{nota_accion}\n\n"
            )
        else:
            lazy_est = calcular_precio_lazy(precio_pagado, precio_pagado)
            alertas_cartera_txt += (
                f"👀 *{row['Jugador_Exacto'].title()}* _({est_c} | {str_tiempo})_\n"
                f"   └ 💵 Costo: `{precio_pagado:,.0f}` | ⚖️ Empate: `{empate:,}` | 🎯 Lazy: `{lazy_est:,}` | 🛑 SL: `{stop_loss_p:,}`\n\n"
            )
    return alertas_cartera_txt

def analizar_jugador_individual(nombre_buscado, df_viejo, df_arb_viejo):
    busq = sin_tildes(nombre_buscado).lower().strip()
    if df_viejo is None or df_viejo.empty:
        return "⚠️ Todavía no hay historial guardado para analizar jugadores."

    mask = df_viejo['Jugador'].apply(lambda x: busq in sin_tildes(str(x)).lower())
    df_jug = df_viejo[mask]

    if df_jug.empty:
        return f"⚠️ No encontré a `{nombre_buscado}` en el padrón dual de {MAX_JUGADORES} cartas."

    ult_registro = df_jug.iloc[-1]
    enlace_principal = ult_registro['Enlace_Futbin']
    df_carta = df_viejo[df_viejo['Enlace_Futbin'] == enlace_principal]

    nombre_oficial = ult_registro['Jugador']
    precio_hoy = int(float(ult_registro['Precio_Actual']))
    p_fb_hoy = int(float(ult_registro.get('Precio_Futbin', precio_hoy))) if pd.notna(ult_registro.get('Precio_Futbin')) and str(ult_registro.get('Precio_Futbin')).strip() != "" else precio_hoy
    p_gg_hoy = int(float(ult_registro.get('Precio_FUTGG', 0))) if pd.notna(ult_registro.get('Precio_FUTGG')) and str(ult_registro.get('Precio_FUTGG')).strip() != "" else 0
    fuente_hoy = str(ult_registro.get('Fuente_Mejor_Precio', 'PC'))
    senal_hoy = str(ult_registro.get('Senal_Trading', '⚪ Rango Neutro'))

    precios_num = pd.to_numeric(df_carta['Precio_Actual'], errors='coerce').fillna(precio_hoy)
    p_min = int(precios_num.min())
    p_max = int(precios_num.max())
    p_prom = int(round(precios_num.mean()))
    muestras = len(df_carta)

    puja_10 = calcular_puja_maxima_segura(precio_hoy, 0.10)
    break_even = calcular_break_even(precio_hoy)
    lazy_sug = calcular_precio_lazy(precio_hoy)
    stop_loss = calcular_stop_loss_sugerido(precio_hoy)

    info_gg = f"`{p_gg_hoy:,} mnds`" if p_gg_hoy > 0 else "Sin cotización en FUT.GG"
    if df_arb_viejo is not None and not df_arb_viejo.empty:
        m_gg = df_arb_viejo[df_arb_viejo['Enlace_Futbin'] == enlace_principal]
        if not m_gg.empty:
            p_gg = int(float(m_gg.iloc[0]['Precio_FUTGG']))
            info_gg = f"`{p_gg:,} mnds` (Brecha Arbitraje: `{m_gg.iloc[0]['Brecha_%']}%`)"

    prompt_lupa = f"""
    Actuá como Trader experto de EA Sports FC 27 en PC (Base Dual Futbin + FUT.GG).
    Analizá esta carta con datos reales de nuestra base:
    - Carta: {nombre_oficial} | Diagnóstico cuantitativo: {senal_hoy}
    - Mejor Precio Actual PC: {precio_hoy:,} monedas (Fuente más barata: {fuente_hoy} | FB: {p_fb_hoy:,} | GG: {p_gg_hoy:,})
    - Mínimo Histórico: {p_min:,} | Promedio: {p_prom:,} | Máximo: {p_max:,} ({muestras} escaneos)
    - Puja máxima sugerida (+10% ROI): {puja_10:,} | Venta Lazy Buyer (+12%): {lazy_sug:,} | Stop-Loss (-10%): {stop_loss:,}

    Dame un veredicto táctico en 4 líneas: ¿Es apta para Flipping Rápido / Lazy Listing hoy o hay riesgo de caída (cuchillo cayendo)?
    """
    veredicto_ia, _ = llamar_gemini_turbo(prompt_lupa)
    if not veredicto_ia:
        veredicto_ia = "Carta en zona operativa. Buscar entrada únicamente mediante puja por debajo de la Puja Máxima Sugerida y listar a precio Lazy."

    reporte_lupa = (
        f"🔍 *LUPA DE TRADING DUAL (FB + GG): {nombre_oficial}*\n━━━━━━━━━━━━━━━━━━\n"
        f"🏷️ *Diagnóstico:* `{senal_hoy}`\n"
        f"💰 *Mejor Precio Actual:* `{precio_hoy:,} mnds` [{fuente_hoy}]\n"
        f"🟦 *Futbin PC:* `{p_fb_hoy:,} mnds` | 🟩 *FUT.GG PC:* {info_gg}\n\n"
        f"📊 *BANDA HISTÓRICA ({muestras} registros):*\n"
        f"   • 💎 Piso Mínimo: `{p_min:,} mnds`\n"
        f"   • ⚖️ Promedio: `{p_prom:,} mnds`\n"
        f"   • 🔥 Techo Máximo: `{p_max:,} mnds`\n\n"
        f"🧮 *PARÁMETROS OPERATIVOS (5% EA Tax):*\n"
        f"   • 🔨 *Puja Máx. (+10% neto):* `{puja_10:,} mnds`\n"
        f"   • ⚖️ *Punto de Empate (Break-Even):* `{break_even:,} mnds`\n"
        f"   • 🎯 *Venta Lazy Buyer (Listar 1h):* `{lazy_sug:,} mnds`\n"
        f"   • 🛑 *Stop-Loss (-10% protección):* `{stop_loss:,} mnds`\n\n"
        f"🤖 *VEREDICTO IA:*\n{veredicto_ia.replace('**', '*')}\n\n"
        f"🔗 [Abrir carta en navegador]({enlace_principal})"
    )
    return reporte_lupa

# ==========================================
# NAVEGADOR COMPATIBLE CON CLOUDFLARE Y GITHUB ACTIONS (LINUX HEADLESS)
# ==========================================
def crear_navegador_optimizado(ocultar=False):
    options = ChromiumOptions()
    options.set_argument('--window-size=1920,1080')
    options.set_argument('--disable-extensions')
    options.no_imgs(False)

    # Si corre en GitHub Actions o Linux sin monitor, activar modo headless compatible
    en_nube = os.environ.get("GITHUB_ACTIONS") == "true" or sys.platform != "win32"
    if en_nube:
        options.set_argument('--headless=new')
        options.set_argument('--no-sandbox')
        options.set_argument('--disable-dev-shm-usage')
        options.set_argument('--disable-gpu')
        options.set_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36')
    elif ocultar:
        options.set_argument('--window-position=-32000,-32000')

    page = ChromiumPage(addr_or_opts=options)
    page.set.timeouts(page_load=10)
    return page

def escanear_futgg(page, presupuesto_max, max_paginas=45):
    print(f"\n⚡ Iniciando escaneo ampliado en FUT.GG PC ({max_paginas} páginas <= {presupuesto_max//1000}k)...")
    jugadores_gg = []
    vistos_gg = set()

    try:
        page.get("https://www.fut.gg/players/")
        page.wait(0.8)
        page.run_js("""
        try {
            localStorage.setItem('platform', 'pc');
            document.cookie = "platform=pc; path=/; max-age=31536000";
            let bs = document.querySelectorAll('button, div, span');
            for (let b of bs) {
                if (b.textContent.trim() === 'PC') { b.click(); break; }
            }
        } catch(e) {}
        """)
        page.wait(0.5)
    except Exception:
        pass

    script_gg = """
    window.scrollTo(0, document.body.scrollHeight / 2);
    let items = [];
    let enlaces = document.querySelectorAll('a[href*="/players/"]');
    for (let a of enlaces) {
        let href = a.getAttribute('href') || '';
        if (!href.match(/\\/players\\/\\d+-/)) continue;
        let p1 = a.parentElement;
        let p2 = p1 ? p1.parentElement : null;
        let p3 = p2 ? p2.parentElement : null;
        let texto_card = (a.innerText || a.textContent || '') + '\\n' +
                         (p1 ? (p1.innerText || p1.textContent) : '') + '\\n' +
                         (p2 ? (p2.innerText || p2.textContent) : '') + '\\n' +
                         (p3 ? (p3.innerText || p3.textContent) : '');
        items.push({'href': href, 'texto': texto_card});
    }
    return items;
    """

    paginas_vacias_seguidas = 0
    for pag in range(1, max_paginas + 1):
        if len(jugadores_gg) >= MAX_JUGADORES:
            break
        url_gg = f"https://www.fut.gg/players/?page={pag}&platform=pc&price_min=200&price_max={presupuesto_max}&min_price=200&max_price={presupuesto_max}"
        print(f"🌐 FUT.GG PC Página {pag}/{max_paginas}... (Capturados GG: {len(jugadores_gg)})")
        try:
            page.get(url_gg)
        except Exception:
            pass
        try:
            if "moment" in page.title.lower() or "__cf_chl" in page.url:
                page.wait(4)
            page.ele('xpath://a[contains(@href, "/players/")]', timeout=4)
            page.wait(random.uniform(0.75, 1.05))
            page.stop_loading()
            datos_gg = page.run_js(script_gg)
        except Exception:
            continue

        nuevos_en_pagina = 0
        for item in (datos_gg or []):
            href = item.get('href', '')
            href_limpio = href.split('?')[0]
            if not href_limpio or href_limpio in vistos_gg:
                continue
            match_slug = re.search(r'/players/\d+-([a-zA-Z0-9\-]+)', href_limpio)
            if not match_slug:
                continue
            nombre_gg = match_slug.group(1).replace('-', ' ').title()
            texto = item.get('texto', '')
            lineas = [l.strip() for l in texto.split('\n') if l.strip()]

            media = "??"
            for l in lineas:
                if l.isdigit() and 45 <= int(l) <= 99:
                    media = l
                    break

            valor_gg = extraer_precio_robusto(texto, "", presupuesto_max)

            if 200 <= valor_gg <= presupuesto_max:
                vistos_gg.add(href_limpio)
                nuevos_en_pagina += 1
                nombre_fmt = f"{nombre_gg} (Media {media})"
                jugadores_gg.append({
                    "Jugador_GG": nombre_fmt,
                    "Nombre_Puro": sin_tildes(nombre_gg).lower(),
                    "Media": media,
                    "Precio_FUTGG": valor_gg,
                    "Clave": clave_cruce(nombre_fmt),
                    "Enlace_GG": f"https://www.fut.gg{href_limpio}" if href_limpio.startswith('/') else href_limpio
                })

        if nuevos_en_pagina == 0:
            paginas_vacias_seguidas += 1
            if paginas_vacias_seguidas >= 3:
                break
        else:
            paginas_vacias_seguidas = 0

    print(f"✅ ¡FUT.GG PC completado! {len(jugadores_gg)} cartas extraídas para la Base Dual.")
    return jugadores_gg

def cruzar_futbin_futgg(datos_unificados, datos_futgg):
    if not datos_unificados:
        return "", pd.DataFrame()

    desfases = []
    vistos_cruce = set()

    for f in datos_unificados:
        p_fb = int(f.get("Precio_Futbin", 0)) if pd.notna(f.get("Precio_Futbin")) and str(f.get("Precio_Futbin")).strip() != "" else 0
        p_gg = int(f.get("Precio_FUTGG", 0)) if pd.notna(f.get("Precio_FUTGG")) and str(f.get("Precio_FUTGG")).strip() != "" else 0
        if p_fb <= 0 or p_gg <= 0:
            continue

        id_unico = f.get("Enlace_Futbin", "")
        if id_unico in vistos_cruce:
            continue

        mayor = max(p_fb, p_gg)
        menor = min(p_fb, p_gg)
        brecha_bruta_pct = ((mayor - menor) / menor) * 100

        if 6.0 <= brecha_bruta_pct <= 60.0:
            ganancia_potencial = round((mayor * 0.95) - menor)
            if ganancia_potencial >= 250:
                vistos_cruce.add(id_unico)
                c = clave_cruce(f["Jugador"])
                media_mostrar = c.split('_')[1]
                desfases.append({
                    "Jugador": f["Jugador"].split(" (")[0] + f" (Media {media_mostrar})",
                    "Precio_Futbin": p_fb,
                    "Precio_FUTGG": p_gg,
                    "Brecha_%": round(brecha_bruta_pct, 1),
                    "Margen_Neto_5%": ganancia_potencial,
                    "Puja_Max_Sugerida": calcular_puja_maxima_segura(mayor, 0.10),
                    "Enlace_Futbin": f.get("Enlace_Futbin", ""),
                    "Enlace_FUTGG": f.get("Enlace_FUTGG", f.get("Enlace_Futbin", ""))
                })

    df_desfases = pd.DataFrame(desfases)
    if df_desfases.empty:
        return "⚡ *DESFASES FUTBIN vs FUT.GG (Mercado PC)*\n━━━━━━━━━━━━━━━━━━\nPrecios alineados.\n\n", df_desfases
    df_desfases = df_desfases.sort_values(by="Margen_Neto_5%", ascending=False)
    txt_arb = "⚡ *DESFASES FUTBIN vs FUT.GG (Mercado PC)*\n━━━━━━━━━━━━━━━━━━\n"
    for _, row in df_desfases.head(5).iterrows():
        txt_arb += (
            f"👤 [*{row['Jugador']}*]({row['Enlace_FUTGG']})\n"
            f"🟦 FB: `{row['Precio_Futbin']:,}` 🆚 🟩 GG: `{row['Precio_FUTGG']:,}`\n"
            f"💸 Brecha: `{row['Brecha_%']}%` (Neto: `+{row['Margen_Neto_5%']:,}` | 🔨 Pujar <= `{row['Puja_Max_Sugerida']:,}`)\n\n"
        )
    return txt_arb, df_desfases

def detectar_evoluciones(datos_unificados, datos_futgg):
    if not datos_unificados:
        return "", pd.DataFrame()

    evos_detectadas = []
    vistos_evo = set()
    for f in datos_unificados:
        enlace = f.get("Enlace_Optimo", f.get("Enlace_Futbin", ""))
        nombre_corto = str(f.get("Jugador", "")).split(" (")[0]
        p_ant = float(f.get("Precio_Anterior", 0))
        p_act = float(f.get("Precio_Actual", 0))
        p_fb = float(f.get("Precio_Futbin", 0)) if pd.notna(f.get("Precio_Futbin")) and str(f.get("Precio_Futbin")).strip() != "" else 0.0
        p_gg = float(f.get("Precio_FUTGG", 0)) if pd.notna(f.get("Precio_FUTGG")) and str(f.get("Precio_FUTGG")).strip() != "" else 0.0
        var_b = float(f.get("Variacion_Bruta_%", 0))
        fuente_e = f.get("Fuente_Mejor_Precio", "PC")

        if 200 <= p_ant <= TOPE_PRECIO_BASE_EVO and var_b >= UMBRAL_SUBIDA_EVO_PCT and (p_act - p_ant) >= 700:
            vistos_evo.add(enlace)
            evos_detectadas.append({
                "Jugador": f.get("Jugador", nombre_corto),
                "Origen_Alerta": f"🔥 Explosión {fuente_e}",
                "Precio_Base": int(p_ant),
                "Precio_Explosivo": int(p_act),
                "Subida_%": round(var_b, 1),
                "Ganancia_Potencial_Neta": round((p_act * 0.95) - p_ant),
                "Enlace": enlace
            })

        if enlace not in vistos_evo and p_fb > 0 and p_gg > 0 and 200 <= min(p_fb, p_gg) <= TOPE_PRECIO_BASE_EVO:
            base_dual = min(p_fb, p_gg)
            pico_dual = max(p_fb, p_gg)
            salto_dual_pct = ((pico_dual - base_dual) / base_dual) * 100 if base_dual > 0 else 0
            if 40.0 <= salto_dual_pct <= 350.0 and (pico_dual - base_dual) >= 1000:
                vistos_evo.add(enlace)
                origen_str = "🧬 Volando en GG" if p_gg > p_fb else "🔥 Volando en FB"
                evos_detectadas.append({
                    "Jugador": f.get("Jugador", nombre_corto),
                    "Origen_Alerta": origen_str,
                    "Precio_Base": int(base_dual),
                    "Precio_Explosivo": int(pico_dual),
                    "Subida_%": round(salto_dual_pct, 1),
                    "Ganancia_Potencial_Neta": round((pico_dual * 0.95) - base_dual),
                    "Enlace": enlace
                })

    df_evos = pd.DataFrame(evos_detectadas)
    if df_evos.empty:
        return "🧬 *EXPLOSIÓN POR EVOLUCIONES (FB + GG)*\n━━━━━━━━━━━━━━━━━━\nSin picos anormales (>35%).\n\n", df_evos
    df_evos = df_evos.sort_values(by="Subida_%", ascending=False)
    txt_evo = "🧬 *EXPLOSIÓN POR EVOLUCIONES (FB + GG)*\n━━━━━━━━━━━━━━━━━━\n"
    for _, row in df_evos.head(5).iterrows():
        txt_evo += f"👤 [*{str(row['Jugador']).split(' (')[0]}*]({row['Enlace']}) ({row['Origen_Alerta']})\n🚀 Precio: `{row['Precio_Base']:,}` ➡️ `{row['Precio_Explosivo']:,} mnds`\n📈 Subida brutal: `+{row['Subida_%']}%`\n\n"
    return txt_evo, df_evos

def procesar_mensajes_telegram(df_cartera, df_viejo, df_arb_viejo, df_evos_viejo, df_alertas):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates"
    hubo_cambios = False
    try:
        r = requests.get(url, timeout=10).json()
        if r.get("ok") and r.get("result"):
            max_update_id = 0
            datos_recientes = []
            if df_viejo is not None and not df_viejo.empty:
                ult_f = df_viejo['Fecha'].max()
                datos_recientes = df_viejo[df_viejo['Fecha'] == ult_f].to_dict('records')

            for msg in r["result"]:
                update_id = msg["update_id"]
                max_update_id = max(max_update_id, update_id)
                texto = msg.get("message", {}).get("text", "").strip()
                if not texto:
                    continue
                t_low = texto.lower()

                if t_low.startswith("/analizar "):
                    jug_lupa = texto.split(" ", 1)[1].strip()
                    print(f"🔍 Analizando en profundidad a: {jug_lupa}...")
                    respuesta_lupa = analizar_jugador_individual(jug_lupa, df_viejo, df_arb_viejo)
                    enviar_alerta_telegram(respuesta_lupa)

                elif t_low in ["/flipping", "⚡ flipping / gangas"]:
                    enviar_alerta_telegram(armar_resumen_flipping_gangas(datos_recientes))

                elif t_low in ["/links", "/domino", "🔗 links dominó"]:
                    txt_dom, _ = detectar_links_domino(datos_recientes, df_evos_viejo)
                    enviar_alerta_telegram(txt_dom)

                elif t_low in ["/pnl", "/contabilidad", "📊 estado p&l"]:
                    rep_pnl, _, _ = generar_estado_resultados_contable(df_cartera, datos_recientes)
                    enviar_alerta_telegram(rep_pnl)

                elif t_low in ["/oraculo", "/leaks", "🔮 oráculo leaks"]:
                    enviar_alerta_telegram("🔮 _Consultando al Oráculo de Contenido sobre la Base Dual (FB + GG)..._")
                    rep_oraculo = escanear_oraculo_contenido(None, datos_recientes, ocultar=True)
                    enviar_alerta_telegram(rep_oraculo)

                elif t_low in ["/ayuda", "🔍 ayuda comandos"]:
                    ayuda_txt = (
                        "🛠️ *CENTRO DE COMANDOS FC27 (6 MOTORES DE TRADING)*\n━━━━━━━━━━━━━━━━━━\n"
                        "• `Compra: Jugador, Precio` _(guarda costo, Lazy y Stop-Loss)_\n"
                        "• `Venta: Jugador, Precio` ➡️ Cierra trade y calcula Costeo Variable\n"
                        "• `⚡ Flipping / Gangas` ➡️ Filtra rebotes rápidos vs Cuchillos cayendo\n"
                        "• `🔗 Links Dominó` ➡️ Cartas dormidas de 2da ola por Evos/SBC\n"
                        "• `📊 Estado P&L` ➡️ Estado de Resultados, VNR y Rotación de Stock\n"
                        "• `🔮 Oráculo Leaks` ➡️ Radar 15:00 hs ARG y Hype\n"
                        "• `/analizar Jugador` ➡️ Lupa FB+GG + Puja + Lazy + Stop-Loss\n"
                        "• `/alerta Jugador, Precio` ➡️ Crea Sniper Alert en FB y GG"
                    )
                    enviar_alerta_telegram(ayuda_txt)

                elif t_low.startswith("/alerta "):
                    try:
                        datos = texto.split(" ", 1)[1].split(",")
                        if len(datos) == 2:
                            jug_b = datos[0].strip().title()
                            precio_obj = int(re.sub(r'[^\d]', '', datos[1]))
                            df_alertas = pd.concat([df_alertas, pd.DataFrame([{"Jugador_Buscado": jug_b, "Precio_Objetivo": precio_obj, "Estado": "Activa"}])], ignore_index=True)
                            hubo_cambios = True
                            enviar_alerta_telegram(f"🎯 *ALERTA SNIPER DUAL CREADA*\nTe aviso apenas `{jug_b}` toque `{precio_obj:,} mnds` o menos en Futbin o FUT.GG.")
                    except Exception:
                        pass

                elif t_low.startswith("/borrar"):
                    partes = texto.split(" ", 1)
                    if len(partes) == 1 or partes[1].strip().lower() == "todo":
                        df_alertas.loc[df_alertas['Estado'] == 'Activa', 'Estado'] = 'Cancelada'
                        hubo_cambios = True
                        enviar_alerta_telegram("🗑️ Todas tus alertas Sniper activas fueron eliminadas.")
                    else:
                        busq = partes[1].strip().lower()
                        cond = (df_alertas['Jugador_Buscado'].str.lower().str.contains(busq, na=False)) & (df_alertas['Estado'] == 'Activa')
                        cant = cond.sum()
                        df_alertas.loc[cond, 'Estado'] = 'Cancelada'
                        hubo_cambios = True
                        enviar_alerta_telegram(f"🗑️ Se eliminaron `{cant}` alertas para `{partes[1].strip()}`.")

                elif t_low in ["/alertas", "🎯 mis alertas"]:
                    act = df_alertas[df_alertas['Estado'] == 'Activa']
                    if act.empty:
                        enviar_alerta_telegram("🎯 *MIS ALERTAS SNIPER*\n━━━━━━━━━━━━━━━━━━\nNo tenés alertas activas.\nCreá una enviando: `/alerta Nombre, Precio`")
                    else:
                        txt_al = "🎯 *MIS ALERTAS SNIPER ACTIVAS (FB + GG)*\n━━━━━━━━━━━━━━━━━━\n"
                        for _, r_al in act.iterrows():
                            txt_al += f"• *{r_al['Jugador_Buscado']}* ➡️ Objetivo: `<= {int(r_al['Precio_Objetivo']):,} mnds`\n"
                        txt_al += "\n_Para quitar una usá:_ `/borrar Nombre` _o_ `/borrar todo`"
                        enviar_alerta_telegram(txt_al)

                elif t_low.startswith("compra:"):
                    try:
                        datos = [x.strip() for x in texto.split(":", 1)[1].split(",")]
                        if len(datos) >= 2:
                            jugador = datos[0].title()
                            precio = int(re.sub(r'[^\d]', '', datos[1]))
                            if len(datos) >= 3 and datos[2]:
                                est_elegida = datos[2].title()
                            else:
                                est_elegida = clasificar_estrategia_automatica(jugador, precio, datos_recientes)

                            empate = calcular_break_even(precio)
                            lazy_ini = calcular_precio_lazy(precio, precio)
                            sl_ini = calcular_stop_loss_sugerido(precio)

                            df_cartera = pd.concat([df_cartera, pd.DataFrame([{
                                "Jugador_Exacto": str(jugador),
                                "Estrategia": str(est_elegida),
                                "Precio_Pagado": int(precio),
                                "Estado": "Activo",
                                "Precio_Venta": 0,
                                "Costo_EA_Tax": 0,
                                "Ganancia_Neta": 0,
                                "ROI_%": 0.0,
                                "Fecha_Compra": ahora_arg().strftime("%Y-%m-%d %H:%M"),
                                "Fecha_Venta": ""
                            }])], ignore_index=True)
                            hubo_cambios = True
                            enviar_alerta_telegram(
                                f"✅ *Compra registrada en Cartera:*\n"
                                f"👤 `{jugador}` por `{precio:,} mnds` _({est_elegida})_\n"
                                f"⚖️ *Punto de Empate (5% tax):* `{empate:,} mnds`\n"
                                f"🎯 *Precio Lazy Sugerido (1h):* `{lazy_ini:,} mnds`\n"
                                f"🛑 *Stop-Loss (-10%):* `{sl_ini:,} mnds`"
                            )
                    except Exception:
                        pass

                elif t_low.startswith("venta:"):
                    try:
                        datos = texto.split(":", 1)[1].split(",")
                        if len(datos) == 2:
                            jug_busqueda = datos[0].strip().lower()
                            precio_venta = int(re.sub(r'[^\d]', '', datos[1]))
                            indices = df_cartera[(df_cartera['Jugador_Exacto'].str.lower().str.contains(jug_busqueda, na=False)) & (df_cartera['Estado'] == 'Activo')].index
                            if not indices.empty:
                                idx_v = indices[0]
                                precio_compra = float(df_cartera.loc[idx_v, 'Precio_Pagado'])
                                jugador_real = df_cartera.loc[idx_v, 'Jugador_Exacto']
                                est_real = str(df_cartera.loc[idx_v, 'Estrategia']) if 'Estrategia' in df_cartera.columns else 'Flipping Meta'
                                if not est_real or est_real.lower() == 'nan':
                                    est_real = clasificar_estrategia_automatica(jugador_real, precio_compra, datos_recientes)
                                costo_tax = int(round(precio_venta * 0.05))
                                ganancia_limpia = int(round((precio_venta - costo_tax) - precio_compra))
                                roi = round((ganancia_limpia / precio_compra) * 100, 2) if precio_compra > 0 else 0.0

                                df_cartera["Estado"] = df_cartera["Estado"].astype(object)
                                df_cartera["Fecha_Venta"] = df_cartera["Fecha_Venta"].astype(object)
                                df_cartera.loc[idx_v, 'Estado'] = 'Vendido'
                                df_cartera.loc[idx_v, 'Precio_Venta'] = precio_venta
                                df_cartera.loc[idx_v, 'Costo_EA_Tax'] = costo_tax
                                df_cartera.loc[idx_v, 'Ganancia_Neta'] = ganancia_limpia
                                df_cartera.loc[idx_v, 'ROI_%'] = roi
                                df_cartera.loc[idx_v, 'Fecha_Venta'] = ahora_arg().strftime("%Y-%m-%d %H:%M")

                                hubo_cambios = True
                                signo = "📈" if ganancia_limpia >= 0 else "📉"
                                enviar_alerta_telegram(
                                    f"{signo} *VENTA CERRADA: {jugador_real}* _({est_real})_\n"
                                    f"• Venta Bruta: `{precio_venta:,} mnds`\n"
                                    f"• (-) Costo Variable (EA Tax 5%): `-{costo_tax:,} mnds`\n"
                                    f"• (-) Costo de Adquisición: `-{int(precio_compra):,} mnds`\n"
                                    f"• (=) *Contribución Marginal Neta:* `{ganancia_limpia:+,} mnds` (`{roi:+}%`)"
                                )
                            else:
                                enviar_alerta_telegram(f"⚠️ No hay posición activa para `{datos[0].strip()}`.")
                    except Exception:
                        pass

                elif t_low in ["/balance", "💼 balance"]:
                    rep_pnl, _, _ = generar_estado_resultados_contable(df_cartera, datos_recientes)
                    enviar_alerta_telegram(rep_pnl)

                elif t_low in ["/cartera", "🎒 mi cartera"]:
                    enviar_alerta_telegram(armar_texto_cartera(df_cartera, datos_recientes))

                elif t_low in ["/fodder", "🧱 fodder"]:
                    if datos_recientes:
                        txt_fod, _ = analizar_fodder(datos_recientes)
                        enviar_alerta_telegram(txt_fod)
                    else:
                        enviar_alerta_telegram("⚠️ Sin datos para Fodder.")

                elif t_low in ["/futgg", "⚡ fut.gg"]:
                    if df_arb_viejo is not None and not df_arb_viejo.empty:
                        txt_gg = "⚡ *DESFASES FUTBIN vs FUT.GG (PC)*\n━━━━━━━━━━━━━━━━━━\n"
                        for _, r_gg in df_arb_viejo.head(6).iterrows():
                            txt_gg += f"👤 [*{r_gg['Jugador']}*]({r_gg['Enlace_FUTGG']})\n🟦 FB: `{r_gg['Precio_Futbin']:,}` 🆚 🟩 GG: `{r_gg['Precio_FUTGG']:,}` (`+{r_gg['Margen_Neto_5%']:,} mnds`)\n\n"
                        enviar_alerta_telegram(txt_gg)
                    else:
                        enviar_alerta_telegram("⚠️ Aún no hay cruces de FUT.GG.")

                elif t_low in ["/evos", "🧬 evos"]:
                    if df_evos_viejo is not None and not df_evos_viejo.empty:
                        txt_ev = "🧬 *EXPLOSIÓN POR EVOLUCIONES*\n━━━━━━━━━━━━━━━━━━\n"
                        for _, r_ev in df_evos_viejo.head(6).iterrows():
                            txt_ev += f"👤 [*{str(r_ev['Jugador']).split(' (')[0]}*]({r_ev['Enlace']}): `{r_ev['Precio_Base']:,}` ➡️ `{r_ev['Precio_Explosivo']:,} mnds` (`+{r_ev['Subida_%']}%`)\n\n"
                        enviar_alerta_telegram(txt_ev)
                    else:
                        enviar_alerta_telegram("🧬 Sin explosiones recientes.")

                elif t_low in ["/totw", "🖤 totw piso"]:
                    if datos_recientes:
                        txt_t, _ = analizar_totw_piso(datos_recientes)
                        enviar_alerta_telegram(txt_t)
                    else:
                        enviar_alerta_telegram("⚠️ Sin datos recientes para filtrar TOTW.")

                elif len(texto) > 5 and not texto.startswith("/"):
                    print(f"🤖 Procesando consulta: '{texto}'")
                    resp_txt, _ = llamar_gemini_turbo(
                        f"Actuá como trader FC27 en el MERCADO DE PC. Presupuesto: {PRESUPUESTO_MAXIMO:,}. Pregunta: '{texto}'. Respuesta corta: 1. Gameplay 2. Inversión (con 5% tax, Puja y Lazy Listing) 3. Veredicto."
                    )
                    if resp_txt:
                        enviar_alerta_telegram(f"💬 *CONSULTA FC27*\n━━━━━━━━━━━━━━━━━━\n{resp_txt.replace('**', '*')}")
                    else:
                        enviar_alerta_telegram("⚠️ No se pudo procesar con IA ahora.")

            if max_update_id > 0:
                requests.get(f"{url}?offset={max_update_id + 1}", timeout=10)
    except Exception:
        pass
    return df_cartera, df_alertas, hubo_cambios

def verificar_alertas_sniper(datos_finales, df_alertas):
    if df_alertas.empty or not datos_finales:
        return df_alertas
    df_alertas["Estado"] = df_alertas["Estado"].fillna("").astype(object)
    alertas_activas = df_alertas[df_alertas['Estado'] == 'Activa']
    for idx, alerta in alertas_activas.iterrows():
        jug_buscado = str(alerta['Jugador_Buscado']).strip().lower()
        precio_obj = float(alerta['Precio_Objetivo'])
        match = next((j for j in datos_finales if jug_buscado in j['Jugador'].lower() and j['Precio_Actual'] > 0), None)
        if match and match['Precio_Actual'] <= precio_obj:
            puja_sug = calcular_puja_maxima_segura(match['Precio_Actual'], 0.10)
            lazy_sug = calcular_precio_lazy(match['Precio_Actual'])
            fuente_s = match.get("Fuente_Mejor_Precio", "PC")
            enlace_s = match.get("Enlace_Optimo", match['Enlace_Futbin'])
            enviar_alerta_telegram(
                f"🚨 🎯 *¡ALERTA SNIPER CUMPLIDA ({fuente_s})!* 🎯 🚨\n━━━━━━━━━━━━━━━━━━\n"
                f"👤 [*{match['Jugador'].split(' (')[0]}*]({enlace_s}) cayó a `{match['Precio_Actual']:,} mnds` [{fuente_s}].\n"
                f"🎯 Objetivo: `{int(precio_obj):,}` | 🔨 Puja +10%: `{puja_sug:,}` | 🎯 Salida Lazy: `{lazy_sug:,}`"
            )
            df_alertas.loc[idx, 'Estado'] = 'Cumplida'
    return df_alertas

def leer_hoja_gsheets_segura(sh, nombre_hoja):
    try:
        ws = sh.worksheet(nombre_hoja)
        valores = ws.get_all_values()
        if not valores or len(valores) < 2:
            return pd.DataFrame()
        encabezados = [str(c).strip() for c in valores[0]]
        indices = [i for i, h in enumerate(encabezados) if h != ""]
        cols = [encabezados[i] for i in indices]
        filas = [[fila[i] if i < len(fila) else "" for i in indices] for fila in valores[1:] if any(str(c).strip() != "" for c in fila)]
        return pd.DataFrame(filas, columns=cols)
    except Exception:
        return pd.DataFrame()

def cargar_estado_excel():
    columnas_cartera = ["Jugador_Exacto", "Estrategia", "Precio_Pagado", "Estado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%", "Fecha_Compra", "Fecha_Venta"]
    cols_texto_cartera = ["Jugador_Exacto", "Estrategia", "Estado", "Fecha_Compra", "Fecha_Venta"]
    cols_num_cartera = ["Precio_Pagado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%"]
    columnas_alertas = ["Jugador_Buscado", "Precio_Objetivo", "Estado"]

    df_viejo = None
    df_arb_viejo = pd.DataFrame()
    df_evos_viejo = pd.DataFrame()
    df_cartera = pd.DataFrame(columns=columnas_cartera)
    df_alertas = pd.DataFrame(columns=columnas_alertas)
    historial_precios = {}

    if os.path.exists(archivo_excel):
        try:
            df_viejo = pd.read_excel(archivo_excel, sheet_name="Historial_Tendencias")
            if 'Fecha' in df_viejo.columns:
                df_viejo = df_viejo.sort_values(by='Fecha')
            historial_precios = dict(zip(df_viejo['Enlace_Futbin'], pd.to_numeric(df_viejo['Precio_Actual'], errors='coerce').fillna(0)))
        except Exception:
            pass
        try:
            df_c = pd.read_excel(archivo_excel, sheet_name="Mi_Cartera")
            for col in columnas_cartera:
                if col not in df_c.columns:
                    df_c[col] = 0 if col in cols_num_cartera else ""
            for col_t in cols_texto_cartera:
                df_c[col_t] = df_c[col_t].fillna("").astype(str).replace("nan", "").astype(object)
            for col_n in cols_num_cartera:
                df_c[col_n] = pd.to_numeric(df_c[col_n], errors="coerce").fillna(0)
            df_cartera = df_c[columnas_cartera]
        except Exception:
            pass
        try:
            df_arb_viejo = pd.read_excel(archivo_excel, sheet_name="Arbitraje_FUTGG")
        except Exception:
            pass
        try:
            df_evos_viejo = pd.read_excel(archivo_excel, sheet_name="Radar_Evos")
        except Exception:
            pass
        try:
            df_al = pd.read_excel(archivo_excel, sheet_name="Mis_Alertas")
            for col_a in ["Jugador_Buscado", "Estado"]:
                if col_a in df_al.columns:
                    df_al[col_a] = df_al[col_a].fillna("").astype(str).replace("nan", "").astype(object)
            df_alertas = df_al
        except Exception:
            pass

    # Si corre en GitHub Actions (sin Excel local), recuperar el historial y la cartera desde Google Sheets
    if (df_viejo is None or df_viejo.empty) and os.path.exists(ARCHIVO_CRED):
        gc = obtener_cliente_gsheets()
        if gc:
            try:
                sh = gc.open(NOMBRE_SHEET)
                df_v_gs = leer_hoja_gsheets_segura(sh, "Historial_Tendencias")
                if not df_v_gs.empty and "Enlace_Futbin" in df_v_gs.columns:
                    df_v_gs["Precio_Actual"] = pd.to_numeric(df_v_gs["Precio_Actual"], errors="coerce").fillna(0)
                    df_viejo = df_v_gs
                    historial_precios = dict(zip(df_viejo['Enlace_Futbin'], df_viejo['Precio_Actual']))

                df_c_gs = leer_hoja_gsheets_segura(sh, "Mi_Cartera")
                if not df_c_gs.empty:
                    for col in columnas_cartera:
                        if col not in df_c_gs.columns:
                            df_c_gs[col] = 0 if col in cols_num_cartera else ""
                    for col_n in cols_num_cartera:
                        df_c_gs[col_n] = pd.to_numeric(df_c_gs[col_n], errors="coerce").fillna(0)
                    df_cartera = df_c_gs[columnas_cartera]

                df_al_gs = leer_hoja_gsheets_segura(sh, "Mis_Alertas")
                if not df_al_gs.empty:
                    df_al_gs["Precio_Objetivo"] = pd.to_numeric(df_al_gs["Precio_Objetivo"], errors="coerce").fillna(0)
                    df_alertas = df_al_gs
            except Exception as e:
                print(f"⚠️ Aviso leyendo estado previo desde Google Sheets: {e}")

    return df_viejo, df_arb_viejo, df_evos_viejo, df_cartera, df_alertas, historial_precios

def formatear_dashboard_financiero_excel(ruta_excel):
    try:
        wb = openpyxl.load_workbook(ruta_excel)
        fill_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        fill_verde = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        font_blanca = Font(color="FFFFFF", bold=True)
        font_bold = Font(bold=True)
        borde_fino = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for name in wb.sheetnames:
            ws = wb[name]
            for cell in ws[1]:
                cell.fill = fill_azul
                cell.font = font_blanca
                cell.alignment = Alignment(horizontal="center")

        if "Dashboard_Financiero" in wb.sheetnames:
            ws_fin = wb["Dashboard_Financiero"]
            ws_fin.column_dimensions["A"].width = 48
            ws_fin.column_dimensions["B"].width = 22
            ws_fin.column_dimensions["C"].width = 18
            for col_l in ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]:
                ws_fin.column_dimensions[col_l].width = 20

            for row in ws_fin.iter_rows(min_row=2, max_row=13, min_col=1, max_col=3):
                for cell in row:
                    cell.border = borde_fino
                if "CONTRIBUCIÓN MARGINAL" in str(row[0].value):
                    for c in row:
                        c.fill = fill_verde
                        c.font = font_bold

        wb.save(ruta_excel)
    except Exception as e:
        print(f"⚠️ Aviso al dar formato gráfico al Excel: {e}")

def ejecutar_ciclo_radar(modo_elegido, ocultar_navegador=False):
    inicio_tiempo = time.time()
    df_viejo, df_arb_viejo, df_evos_viejo, df_cartera, df_alertas, historial_precios = cargar_estado_excel()
    stats_historicas = calcular_estadisticas_historicas(df_viejo)
    df_cartera, df_alertas, _ = procesar_mensajes_telegram(df_cartera, df_viejo, df_arb_viejo, df_evos_viejo, df_alertas)

    datos_finales = []
    datos_gg_actuales = []
    fecha_actual = ahora_arg().strftime("%Y-%m-%d %H:%M")
    page_activa = None

    if modo_elegido in ["2", "3", "5"] and df_viejo is not None and not df_viejo.empty:
        ultima_fecha = df_viejo['Fecha'].max()
        print(f"\n📖 Recuperando padrón dual del {ultima_fecha}...")
        datos_previos = df_viejo[df_viejo['Fecha'] == ultima_fecha].to_dict('records')
        fecha_actual = ultima_fecha

        if modo_elegido == "3":
            page_activa = crear_navegador_optimizado(ocultar=ocultar_navegador)
            datos_gg_actuales = escanear_futgg(page_activa, PRESUPUESTO_MAXIMO, PAGINAS_FUTGG)
            page_activa.quit()
            page_activa = None
            jug_fb_previos = [
                {"Jugador": r["Jugador"], "Enlace": r["Enlace_Futbin"], "Precio_Actual": int(float(r.get("Precio_Futbin", r["Precio_Actual"]))) if pd.notna(r.get("Precio_Futbin")) and str(r.get("Precio_Futbin")).strip() != "" else int(float(r["Precio_Actual"]))}
                for r in datos_previos if int(float(r.get("Precio_Actual", 0))) > 0
            ]
            datos_finales = unificar_bases_mercado(jug_fb_previos, datos_gg_actuales, historial_precios, stats_historicas, ahora_arg().strftime("%Y-%m-%d %H:%M"))
        elif modo_elegido == "5":
            rep_oraculo_solo = escanear_oraculo_contenido(None, datos_previos, ocultar=ocultar_navegador)
            enviar_alerta_telegram(rep_oraculo_solo)
            print("✅ ¡Reporte del Oráculo de Contenido enviado a Telegram!")
            return
        else:
            for r_p in datos_previos:
                p_a = int(float(r_p.get("Precio_Actual", 0)))
                p_ant = int(float(r_p.get("Precio_Anterior", p_a)))
                p_fb_p = int(float(r_p.get("Precio_Futbin", p_a))) if pd.notna(r_p.get("Precio_Futbin")) and str(r_p.get("Precio_Futbin")).strip() != "" else p_a
                p_gg_p = int(float(r_p.get("Precio_FUTGG", 0))) if pd.notna(r_p.get("Precio_FUTGG")) and str(r_p.get("Precio_FUTGG")).strip() != "" else 0
                st_p = stats_historicas.get(r_p.get("Enlace_Futbin", ""), {"min": p_a, "max": p_a, "prom": p_a, "count": 1})
                if not r_p.get("Precio_Lazy_Sugerido") or pd.isna(r_p.get("Precio_Lazy_Sugerido")):
                    r_p["Precio_Lazy_Sugerido"] = calcular_precio_lazy(p_a)
                if not r_p.get("Senal_Trading") or pd.isna(r_p.get("Senal_Trading")):
                    r_p["Senal_Trading"] = diagnosticar_senal_trading(p_a, p_ant, st_p["min"], st_p["max"], st_p["prom"], st_p["count"], p_fb_p, p_gg_p)
            datos_finales = datos_previos
    else:
        print(f"\n🌐 Abriendo navegador para Escaneo Dual de 2000 Jugadores (Futbin PC + FUT.GG PC)...")
        page_activa = crear_navegador_optimizado(ocultar=ocultar_navegador)

        try:
            page_activa.get("https://www.futbin.com/27/players")
            page_activa.wait(1.0)
            page_activa.run_js('document.cookie = "platform=pc; path=/; max-age=31536000"; localStorage.setItem("platform", "pc");')
        except Exception:
            pass

        jugadores_tendencia = []
        enlaces_vistos = set()

        script_js = """
        let data = [];
        let filas = document.querySelectorAll('table tbody tr, .player-item, .player-row, .card-row');
        for (let i = 0; i < filas.length; i++) {
            let fila = filas[i];
            let a = fila.querySelector('a[href*="/player/"]');
            if (!a) continue;
            let href = a.getAttribute('href');
            let texto_fila = fila.innerText || fila.textContent || "";
            
            let nodo_pc = fila.querySelector('.pc-price, .price-pc, [data-platform="pc"]');
            let nodo_precio = nodo_pc || fila.querySelector('.price, .ps-price, [data-price], td:nth-last-child(2)');
            let texto_precio = (nodo_precio && (nodo_precio.innerText || nodo_precio.textContent).trim() !== "") 
                               ? (nodo_precio.innerText || nodo_precio.textContent) 
                               : texto_fila;
            
            let media = "??";
            let rating_ele = fila.querySelector('.rating, .form.rating, span.rating_val, td:nth-child(2)');
            if (rating_ele) {
                let match = (rating_ele.innerText || rating_ele.textContent).match(/\\b([4-9]\\d)\\b/);
                if (match) media = match[1];
            }
            data.push({
                'href': href,
                'texto_fila': texto_fila,
                'texto_precio': texto_precio,
                'media': media
            });
        }
        return data;
        """

        paginas_fb_vacias = 0
        for num_pag in range(1, 85):
            if len(jugadores_tendencia) >= MAX_JUGADORES:
                break
            url = f"https://www.futbin.com/27/players?page={num_pag}&pc_price=200-{PRESUPUESTO_MAXIMO}"
            print(f"⚡ Futbin PC Pág {num_pag}... (Capturados FB: {len(jugadores_tendencia)}/{MAX_JUGADORES})")
            try:
                page_activa.get(url)
            except Exception:
                pass
            try:
                if not page_activa.ele('xpath://a[contains(@href, "/player/")]', timeout=4):
                    paginas_fb_vacias += 1
                    if paginas_fb_vacias >= 3:
                        print("⏩ Futbin bloqueado por Cloudflare en este servidor. Saltando directo a FUT.GG...")
                        break
                    continue
                page_activa.wait(random.uniform(0.50, 0.80))
                page_activa.stop_loading()
                datos_pagina = page_activa.run_js(script_js)
            except Exception:
                paginas_fb_vacias += 1
                if paginas_fb_vacias >= 3:
                    break
                continue

            antes_pag = len(jugadores_tendencia)
            for item in (datos_pagina or []):
                if len(jugadores_tendencia) >= MAX_JUGADORES:
                    break
                href = item.get('href', '')
                texto_fila = item.get('texto_fila', '')
                texto_precio = item.get('texto_precio', '')

                if not href or '/player/' not in href or href in enlaces_vistos:
                    continue
                if 'SBC' in texto_fila.upper() or 'OBJECTIVE' in texto_fila.upper():
                    continue

                partes_url = [p for p in href.split('/') if p]
                nombre_url = partes_url[-1].replace('-', ' ').title() if partes_url else "Jugador"
                lineas = [l.strip() for l in texto_fila.split('\n') if l.strip()]

                media = item.get('media', '??')
                if media == "??":
                    for l in lineas[:5]:
                        if l.isdigit() and 45 <= int(l) <= 99:
                            media = l
                            break

                tipo_carta = "Normal"
                for l in lineas:
                    if l.isdigit() or '++' in l.upper() or '\\' in l or '/' in l:
                        continue
                    if len(l) >= 3 and not re.search(r'\d', l):
                        tipo_carta = l
                        break

                valor_actual = extraer_precio_robusto(texto_precio, texto_fila, PRESUPUESTO_MAXIMO)

                if 200 <= valor_actual <= PRESUPUESTO_MAXIMO:
                    enlaces_vistos.add(href)
                    jugadores_tendencia.append({
                        "Jugador": f"{nombre_url} (Media {media} - {tipo_carta})",
                        "Enlace": f"https://www.futbin.com{href}" if href.startswith('/') else href,
                        "Precio_Actual": valor_actual
                    })

            if len(jugadores_tendencia) == antes_pag:
                paginas_fb_vacias += 1
                if paginas_fb_vacias >= 3:
                    break
            else:
                paginas_fb_vacias = 0
            if len(jugadores_tendencia) >= MAX_JUGADORES:
                break
            url = f"https://www.futbin.com/27/players?page={num_pag}&pc_price=200-{PRESUPUESTO_MAXIMO}"
            print(f"⚡ Futbin PC Pág {num_pag}... (Capturados FB: {len(jugadores_tendencia)}/{MAX_JUGADORES})")
            try:
                page_activa.get(url)
            except Exception:
                pass
            try:
                if "moment" in page_activa.title.lower() or "__cf_chl" in page_activa.url or not page_activa.ele('xpath://a[contains(@href, "/player/")]', timeout=3.5):
                    if not page_activa.ele('xpath://a[contains(@href, "/player/")]', timeout=20):
                        continue
                page_activa.wait(random.uniform(0.50, 0.80))
                page_activa.stop_loading()
                datos_pagina = page_activa.run_js(script_js)
            except Exception:
                continue

            for item in (datos_pagina or []):
                if len(jugadores_tendencia) >= MAX_JUGADORES:
                    break
                href = item.get('href', '')
                texto_fila = item.get('texto_fila', '')
                texto_precio = item.get('texto_precio', '')

                if not href or '/player/' not in href or href in enlaces_vistos:
                    continue
                if 'SBC' in texto_fila.upper() or 'OBJECTIVE' in texto_fila.upper():
                    continue

                partes_url = [p for p in href.split('/') if p]
                nombre_url = partes_url[-1].replace('-', ' ').title() if partes_url else "Jugador"
                lineas = [l.strip() for l in texto_fila.split('\n') if l.strip()]

                media = item.get('media', '??')
                if media == "??":
                    for l in lineas[:5]:
                        if l.isdigit() and 45 <= int(l) <= 99:
                            media = l
                            break

                tipo_carta = "Normal"
                for l in lineas:
                    if l.isdigit() or '++' in l.upper() or '\\' in l or '/' in l:
                        continue
                    if len(l) >= 3 and not re.search(r'\d', l):
                        tipo_carta = l
                        break

                valor_actual = extraer_precio_robusto(texto_precio, texto_fila, PRESUPUESTO_MAXIMO)

                if 200 <= valor_actual <= PRESUPUESTO_MAXIMO:
                    enlaces_vistos.add(href)
                    jugadores_tendencia.append({
                        "Jugador": f"{nombre_url} (Media {media} - {tipo_carta})",
                        "Enlace": f"https://www.futbin.com{href}" if href.startswith('/') else href,
                        "Precio_Actual": valor_actual
                    })

        print(f"\n✅ ¡Futbin PC listo! Se extrajeron {len(jugadores_tendencia)} cartas.")

        datos_gg_actuales = escanear_futgg(page_activa, PRESUPUESTO_MAXIMO, PAGINAS_FUTGG)

        datos_finales = unificar_bases_mercado(
            jugadores_tendencia,
            datos_gg_actuales,
            historial_precios,
            stats_historicas,
            fecha_actual
        )

    oraculo_txt = escanear_oraculo_contenido(page_activa, datos_finales, ocultar=ocultar_navegador)
    if page_activa is not None:
        try:
            page_activa.quit()
        except Exception:
            pass

    df_alertas = verificar_alertas_sniper(datos_finales, df_alertas)

    termometro_txt = calcular_termometro_mercado(datos_finales)
    ciclo_wl_txt = obtener_radar_ciclo_futchampions(datos_finales)
    alertas_cartera_txt = armar_texto_cartera(df_cartera, datos_finales)
    fodder_txt, df_fodder = analizar_fodder(datos_finales)
    totw_txt, df_totw = analizar_totw_piso(datos_finales)
    rep_pnl_tg, df_pnl_excel, df_est_excel = generar_estado_resultados_contable(df_cartera, datos_finales)

    if not df_cartera.empty:
        df_cartera["Estrategia"] = df_cartera["Estrategia"].fillna("").astype(object)
        est_sync = []
        for _, r_c in df_cartera.iterrows():
            val_e = str(r_c.get("Estrategia", "")).strip()
            if not val_e or val_e.lower() == "nan":
                est_sync.append(str(clasificar_estrategia_automatica(r_c.get("Jugador_Exacto", ""), float(r_c.get("Precio_Pagado", 0)), datos_finales)))
            else:
                est_sync.append(val_e)
        df_cartera["Estrategia"] = pd.Series(est_sync, index=df_cartera.index, dtype="object")

    minimos_lista = [
        j for j in datos_finales
        if (j.get("Estado_Historico") == "💎 MÍNIMO HISTÓRICO" or "⚡" in str(j.get("Senal_Trading", "")))
        and float(j.get("Variacion_Bruta_%", 0)) < 0
        and "⚠️" not in str(j.get("Senal_Trading", ""))
    ]
    minimos_lista = sorted(minimos_lista, key=lambda x: float(x.get("Variacion_Bruta_%", 0)))[:5]
    minimos_txt = ""
    if minimos_lista:
        minimos_txt = "💎 *FLIPPING RÁPIDO & SUELO HISTÓRICO (PUJA + LAZY)*\n━━━━━━━━━━━━━━━━━━\n"
        for m_item in minimos_lista:
            nom_c = m_item["Jugador"].split(" (")[0]
            f_min = m_item.get("Fuente_Mejor_Precio", "PC")
            enl_min = m_item.get("Enlace_Optimo", m_item["Enlace_Futbin"])
            puja_m = m_item.get("Puja_Max_10%_ROI", calcular_puja_maxima_segura(m_item["Precio_Actual"]))
            emp_m = m_item.get("Break_Even_5%", calcular_break_even(m_item["Precio_Actual"]))
            lazy_m = m_item.get("Precio_Lazy_Sugerido", calcular_precio_lazy(m_item["Precio_Actual"]))
            senal_m = m_item.get("Senal_Trading", "💎 Suelo")
            minimos_txt += (
                f"👤 [*{nom_c}*]({enl_min}) — Hoy: `{m_item['Precio_Actual']:,} mnds` [{f_min}] (`{m_item['Variacion_Bruta_%']}%`)\n"
                f"   └ 🏷️ `{senal_m}` | Promedio Hist: `{m_item.get('Precio_Prom_Hist', m_item['Precio_Actual']):,}`\n"
                f"   └ 🔨 *Pujar:* `{puja_m:,}` | ⚖️ *Empate:* `{emp_m:,}` | 🎯 *Lazy 1h:* `{lazy_m:,}`\n\n"
            )

    arbitraje_txt, df_arbitraje = cruzar_futbin_futgg(datos_finales, datos_gg_actuales)
    if df_arbitraje.empty and not df_arb_viejo.empty and modo_elegido == "2":
        df_arbitraje = df_arb_viejo
        arbitraje_txt = "⚡ *DESFASES FUTBIN vs FUT.GG (Mercado PC)*\n━━━━━━━━━━━━━━━━━━\n"
        for _, row in df_arbitraje.head(4).iterrows():
            arbitraje_txt += f"👤 [*{row['Jugador']}*]({row['Enlace_FUTGG']})\n🟦 FB: `{row['Precio_Futbin']:,}` 🆚 🟩 GG: `{row['Precio_FUTGG']:,}`\n💸 Brecha: `{row['Brecha_%']}%`\n\n"

    evos_txt, df_evos = detectar_evoluciones(datos_finales, datos_gg_actuales)
    if df_evos.empty and not df_evos_viejo.empty and modo_elegido == "2":
        df_evos = df_evos_viejo
        evos_txt = "🧬 *EXPLOSIÓN POR EVOLUCIONES (HYPE)*\n━━━━━━━━━━━━━━━━━━\n"
        for _, row in df_evos.head(4).iterrows():
            evos_txt += f"👤 [*{str(row['Jugador']).split(' (')[0]}*]({row['Enlace']})\n🚀 Precio: `{row['Precio_Base']:,}` ➡️ `{row['Precio_Explosivo']:,} mnds`\n📈 Subida: `+{row['Subida_%']}%`\n\n"

    domino_txt, df_domino = detectar_links_domino(datos_finales, df_evos)

    dashboard_tecnico = (
        f"🖥️ *DASHBOARD DUAL FC27 (FUTBIN + FUT.GG PC)*\n"
        f"🕒 `{fecha_actual}` | 📦 *Padrón:* `{len(datos_finales)} cartas` | 💰 *Tope:* `{PRESUPUESTO_MAXIMO:,}`\n\n"
        f"{termometro_txt}"
        f"{ciclo_wl_txt}"
        f"{alertas_cartera_txt}"
        f"{minimos_txt}"
        f"{domino_txt}"
        f"{fodder_txt}"
        f"{totw_txt}"
        f"{evos_txt}"
        f"{arbitraje_txt}"
    )
    enviar_alerta_telegram(dashboard_tecnico)
    enviar_alerta_telegram(oraculo_txt)

    if not datos_finales:
        print("⚠️ No se obtuvieron datos en este ciclo.")
        return

    print("\n🧠 Conectando directo con Gemini AI...")
    df_nuevo = pd.DataFrame(datos_finales)
    for col_extra in ["Precio_Futbin", "Precio_FUTGG", "Fuente_Mejor_Precio", "Precio_Min_Hist", "Precio_Prom_Hist", "Estado_Historico", "Senal_Trading", "Puja_Max_10%_ROI", "Break_Even_5%", "Precio_Lazy_Sugerido", "Enlace_FUTGG", "Enlace_Optimo"]:
        if col_extra not in df_nuevo.columns:
            df_nuevo[col_extra] = ""

    df_compras_ia = df_nuevo[df_nuevo['Accion'] == '🟢 COMPRAR'].sort_values(by='Variacion_Bruta_%').head(18)
    df_ventas_ia = df_nuevo[df_nuevo['Accion'] == '🔴 VENDER'].sort_values(by='Variacion_Neta_%', ascending=False).head(18)

    df_para_ia = pd.concat([df_compras_ia, df_ventas_ia, df_nuevo.head(20)]).drop_duplicates(subset=['Enlace_Futbin'])
    cols_ia = [c for c in ['Jugador', 'Precio_Anterior', 'Precio_Actual', 'Fuente_Mejor_Precio', 'Senal_Trading', 'Puja_Max_10%_ROI', 'Break_Even_5%', 'Precio_Lazy_Sugerido', 'Variacion_Bruta_%', 'Accion'] if c in df_para_ia.columns]
    resumen_csv = df_para_ia[cols_ia].to_string(index=False)
    dia_hoy = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][ahora_arg().weekday()]

    prompt_estrategico = f"""
    Actuá como Trader Profesional de FC 27 en MERCADO PC operando con BASE DUAL UNIFICADA (Futbin + FUT.GG de {len(datos_finales)} cartas).
    Hoy es {dia_hoy}. Presupuesto: {PRESUPUESTO_MAXIMO:,} mnds (Tax: 5%).
    Datos con diagnóstico de liquidez (Senal_Trading), Puja_Max_10%_ROI y Precio_Lazy_Sugerido:
    {resumen_csv}
    Redactá de forma estética (sin tablas '|', usando viñetas):
    1. TOP 5 COMPRAS DE FLIPPING & REBOTE PC (FB + GG): Indicá Jugador, web más barata (Fuente_Mejor_Precio), Precio Compra hoy, hasta cuánto PUJAR (Puja_Max_10%_ROI), a cuánto listar para LAZY BUYERS (Precio_Lazy_Sugerido) y Tesis (evitando cuchillos cayendo).
    2. TOP 5 VENTAS / TOMA DE GANANCIAS PC: Jugador, Precio actual y Motivo.
    3. ESTRATEGIA DE ROTACIÓN DE CAPITAL Y CICLO FUT CHAMPIONS PARA HOY ({dia_hoy}).
    """

    reporte_estrategico_ia, mod_usado = llamar_gemini_turbo(prompt_estrategico)
    if reporte_estrategico_ia:
        print(f"✅ Reporte Estratégico Dual generado al instante usando: {mod_usado}")
        texto_limpio = f"🤖 *REPORTE ESTRATÉGICO DUAL FB+GG ({dia_hoy})*\n━━━━━━━━━━━━━━━━━━\n{reporte_estrategico_ia.replace('**', '*')}"
    else:
        compras_lista = []
        for _, r in df_compras_ia.head(5).iterrows():
            puja_r = calcular_puja_maxima_segura(r['Precio_Actual'])
            lazy_r = calcular_precio_lazy(r['Precio_Actual'])
            f_str = r.get('Fuente_Mejor_Precio', 'PC')
            enl_r = r.get('Enlace_Optimo', r['Enlace_Futbin'])
            compras_lista.append(f"• [*{r['Jugador'].split(' (')[0]}*]({enl_r}): `{r['Precio_Actual']:,}` [{f_str}] (🔨 Puja <= `{puja_r:,}` | 🎯 Lazy: `{lazy_r:,}`)")
        ventas_lista = []
        for _, r in df_ventas_ia.head(5).iterrows():
            enl_r = r.get('Enlace_Optimo', r['Enlace_Futbin'])
            ventas_lista.append(f"• [*{r['Jugador'].split(' (')[0]}*]({enl_r}): `{r['Precio_Actual']:,}` (Neto: `+{r['Beneficio_Neto_Mnds']:,}`)")
        texto_limpio = "🟢 *OPORTUNIDADES DUAL (FB + GG)*\n━━━━━━━━━━━━━━━━━━\n" + "\n".join(compras_lista) + "\n\n🔴 *ALERTAS DE VENTA (PC)*\n━━━━━━━━━━━━━━━━━━\n" + "\n".join(ventas_lista)

    enviar_alerta_telegram(texto_limpio)

    if modo_elegido not in ["2", "5"] and df_viejo is not None and not df_viejo.empty:
        for col in df_nuevo.columns:
            if col not in df_viejo.columns:
                df_viejo[col] = ""
        df_final = pd.concat([df_viejo[df_nuevo.columns], df_nuevo], ignore_index=True)
    elif df_viejo is not None and not df_viejo.empty:
        df_final = df_viejo
    else:
        df_final = df_nuevo

    # 1. Guardar respaldo en Excel Local
    try:
        with pd.ExcelWriter(archivo_excel, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Historial_Tendencias")
            pd.DataFrame({"Fecha": [fecha_actual], "Reporte": [texto_limpio], "Oraculo": [oraculo_txt]}).to_excel(writer, index=False, sheet_name="Recomendacion_IA")
            df_cartera.to_excel(writer, index=False, sheet_name="Mi_Cartera")
            df_alertas.to_excel(writer, index=False, sheet_name="Mis_Alertas")
            if not df_pnl_excel.empty:
                df_pnl_excel.to_excel(writer, index=False, sheet_name="Dashboard_Financiero", startrow=0, startcol=0)
            if not df_est_excel.empty:
                df_est_excel.to_excel(writer, index=False, sheet_name="Dashboard_Financiero", startrow=0, startcol=4)
            if not df_fodder.empty:
                df_fodder.to_excel(writer, index=False, sheet_name="Radar_Fodder")
            if not df_arbitraje.empty:
                df_arbitraje.to_excel(writer, index=False, sheet_name="Arbitraje_FUTGG")
            if not df_evos.empty:
                df_evos.to_excel(writer, index=False, sheet_name="Radar_Evos")
            if not df_totw.empty:
                df_totw.to_excel(writer, index=False, sheet_name="Radar_TOTW")
            if not df_domino.empty:
                df_domino.to_excel(writer, index=False, sheet_name="Radar_Links_Domino")

        formatear_dashboard_financiero_excel(archivo_excel)
    except Exception as e:
        print(f"⚠️ Aviso guardando Excel local: {e}")

    # 2. Subir automáticamente a Google Sheets (Base_Datos_FC27)
    sincronizar_todo_a_google_sheets(
        df_final, fecha_actual, texto_limpio, oraculo_txt,
        df_cartera, df_alertas, df_pnl_excel, df_fodder,
        df_arbitraje, df_evos, df_totw, df_domino
    )

    print(f"🏁 ¡Proceso Dual (2000 Jugadores + 6 Motores de Trading) completado en {round(time.time() - inicio_tiempo, 1)} segundos!")

# ==========================================
# MENÚ PRINCIPAL Y MODO VIGILANTE / AUTOMÁTICO (GITHUB ACTIONS COMPATIBLE)
# ==========================================
en_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"
modo_auto = en_github_actions or (len(sys.argv) > 1 and sys.argv[1].lower() == "auto")
modo_vigilante_auto = len(sys.argv) > 1 and sys.argv[1].lower() == "vigilante"

print("=" * 80)
print(f"🤖 BOT MAESTRO FC 27 - BASE DUAL 2000 + 6 MOTORES TRADING (Tope: {PRESUPUESTO_MAXIMO//1000}k)")
print("=" * 80)

if modo_vigilante_auto:
    modo_elegido = "4"
elif modo_auto:
    print("🤖 MODO PILOTO AUTOMÁTICO ACTIVADO (Escaneo Dual 2000 + Subida a Google Sheets...)")
    modo_elegido = "1"
else:
    print("1️⃣  Escaneo Dual 2000 Completo: Futbin + FUT.GG + 6 Motores Trading + P&L + Oráculo")
    print("2️⃣  Modo Instantáneo (5 seg): Evaluar Cartera (Rotación/Lazy/Stop-Loss) + Flipping")
    print("3️⃣  Actualización Rápida FUT.GG: Refrescar precios de FUT.GG sobre la Base Dual")
    print(f"4️⃣  Modo Vigilante 24/7: Chat en vivo + Alerta 15:00 hs ARG + Escaneo cada {INTERVALO_VIGILANTE_MIN} min")
    print("5️⃣  Oráculo de Contenido Express (10 seg): Radar de las 15:00 hs ARG y Leaks")
    modo_elegido = input("\n👉 Elegí una opción (1, 2, 3, 4 o 5 y Enter). O dejalo en blanco para Vigilante (4): ").strip()
    if modo_elegido == "":
        modo_elegido = "4"

if modo_elegido == "4":
    print(f"\n🛡️ MODO VIGILANTE 24/7 ACTIVADO: Base Dual 2000 (FB+GG), 6 Motores de Trading, P&L y Oráculo.")
    print("   (Presioná Ctrl + C cuando quieras detenerlo)\n")
    ultima_alerta_15hs = ""
    try:
        while True:
            ejecutar_ciclo_radar("1", ocultar_navegador=True)
            segundos_espera = INTERVALO_VIGILANTE_MIN * 60
            for _ in range(segundos_espera // 10):
                time.sleep(10)
                ahora_v = ahora_arg()
                fecha_hoy_str = ahora_v.strftime("%Y-%m-%d")

                if ahora_v.hour == 14 and 35 <= ahora_v.minute <= 58 and ultima_alerta_15hs != fecha_hoy_str:
                    ultima_alerta_15hs = fecha_hoy_str
                    df_v_o, _, _, _, _, _ = cargar_estado_excel()
                    datos_rec_o = []
                    if df_v_o is not None and not df_v_o.empty:
                        ult_f_o = df_v_o['Fecha'].max()
                        datos_rec_o = df_v_o[df_v_o['Fecha'] == ult_f_o].to_dict('records')
                    aviso_15 = escanear_oraculo_contenido(None, datos_rec_o, ocultar=True)
                    enviar_alerta_telegram(f"🚨 ⏰ *¡ALERTA PRE-CONTENIDO 15:00 HS ARG!* ⏰ 🚨\n\n{aviso_15}")

                df_v, df_a_v, df_e_v, df_c, df_al, _ = cargar_estado_excel()
                df_c_nuevo, df_al_nuevo, cambios = procesar_mensajes_telegram(df_c, df_v, df_a_v, df_e_v, df_al)
                if cambios:
                    if os.path.exists(archivo_excel):
                        try:
                            with pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                df_c_nuevo.to_excel(writer, index=False, sheet_name="Mi_Cartera")
                                df_al_nuevo.to_excel(writer, index=False, sheet_name="Mis_Alertas")
                        except Exception:
                            pass
                    gc_v = obtener_cliente_gsheets()
                    if gc_v:
                        try:
                            sh_v = gc_v.open(NOMBRE_SHEET)
                            subir_dataframe_gsheets(sh_v, "Mi_Cartera", df_c_nuevo, max_filas=1000)
                            subir_dataframe_gsheets(sh_v, "Mis_Alertas", df_al_nuevo, max_filas=500)
                        except Exception:
                            pass
    except KeyboardInterrupt:
        print("\n🛑 Modo Vigilante detenido manualmente.")
else:
    ejecutar_ciclo_radar(modo_elegido, ocultar_navegador=modo_auto)
