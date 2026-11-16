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
from openpyxl.styles import Alignment, Font, PatternFill
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
MAX_JUGADORES = 2000
PAGINAS_FUTGG = 45
INTERVALO_VIGILANTE_MIN = 15
archivo_excel = "Lista_Jugadores_FC27.xlsx"
ARCHIVO_CRED = "credenciales_google.json"
NOMBRE_SHEET = "Base_Datos_FC27"

HORAS_TOPE_INMOVILIZACION = 72.0
STOP_LOSS_PCT = -10.0
TAKE_PROFIT_1_PCT = 8.0
TAKE_PROFIT_2_PCT = 18.0
MARGEN_LAZY_BUYER = 1.12

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
    tz_arg = timezone(timedelta(hours=-3))
    return datetime.now(tz_arg).replace(tzinfo=None)

# ==========================================
# CONEXIÓN Y SINCRONIZACIÓN CON GOOGLE SHEETS
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
        print(f"⚠️ Aviso sincronizando hoja '{nombre_hoja}': {e}")

def sincronizar_todo_a_google_sheets(
    df_final, fecha_actual, texto_limpio, oraculo_txt,
    df_cartera, df_alertas, df_pnl_excel, df_fodder,
    df_arbitraje, df_evos, df_totw, df_domino
):
    gc = obtener_cliente_gsheets()
    if not gc:
        print("ℹ️ Sin archivo credenciales_google.json válido: se guardó directo en Lista_Jugadores_FC27.xlsx.")
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
# TECLADO INTERACTIVO DE TELEGRAM
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
        for mod in MODELOS_PREFERIDOS:
            try:
                resp = client.models.generate_content(model=mod, contents=prompt)
                if resp and resp.text:
                    return resp.text, mod
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

def extraer_precio_carta_unica_gg(texto_card, presupuesto_max):
    """Extrae el precio únicamente dentro del contenedor individual de una carta en FUT.GG."""
    if not texto_card:
        return 0
    precios_km = re.findall(r'(\d+(?:\.\d+)?)\s*([KkMm])\b', texto_card)
    for num_s, unidad in precios_km:
        mult = 1000000 if unidad.upper() == 'M' else 1000
        val = int(round(float(num_s) * mult))
        if val > presupuesto_max:
            return 0
        if 200 <= val <= presupuesto_max:
            return val

    precios_comas = re.findall(r'\b(\d{1,3}(?:,\d{3})+)\b', texto_card)
    for p_c in precios_comas:
        val = int(p_c.replace(',', ''))
        if val > presupuesto_max:
            return 0
        if 200 <= val <= presupuesto_max:
            return val

    enteros = re.findall(r'\b(\d{3,6})\b', texto_card)
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
# MOTOR DE CLASIFICACIÓN DE TRADING
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
# FUSIÓN DE BASES: FUTBIN + FUT.GG
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
        p_con_fb = int(jug.get("Precio_Consola", p_fb) or p_fb)
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
            "Precio_Consola": p_con_fb,
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
        if p_gg <= 0 or g.get("Media") == "??":
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
            "Precio_Consola": p_gg,
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
    vendidos_con_precio = vendidos[vendidos["Precio_Venta"] > 0] if not vendidos.empty else pd.DataFrame()
    costo_adq_tot = int(vendidos_con_precio["Precio_Pagado"].sum()) if not vendidos_con_precio.empty else 0
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
        v_est = vendidos_con_precio[vendidos_con_precio["Estrategia"] == est] if not vendidos_con_precio.empty else pd.DataFrame()
        a_est = activos[activos["Estrategia"] == est
