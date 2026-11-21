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
# CONFIGURACIÓN GENERAL Y CREDENCIALES (BASE EXCLUSIVA FUT.GG — 10,000 CARTAS + 15M)
# ==========================================
_G_PART1 = "AQ.Ab8RN6KGQFnimiJeqwsb"
_G_PART2 = "KKcbIWeCUq6SggjfhQ6kIN9aKIFafA"
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", _G_PART1 + _G_PART2)

_T_PART1 = "9090839192:AAGP0TTaNkQP1I"
_T_PART2 = "LRKvqSFvbsyZtdYuM3PPw"
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", _T_PART1 + _T_PART2)
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "7676084666")

PRESUPUESTO_MAXIMO = 15000000  # Tope ampliado a 15 Millones de monedas
MAX_JUGADORES = 10000          # Base exclusiva FUT.GG escalada a 10,000 jugadores
PAGINAS_FUTGG = 350            # Hasta 350 páginas en FUT.GG (~10,000+ cartas)
INTERVALO_VIGILANTE_MIN = 15
archivo_excel = "Lista_Jugadores_FC27.xlsx"
ARCHIVO_CRED = "credenciales_google.json"
NOMBRE_SHEET = "Base_Datos_FC27"

# Parámetros de los 6 Motores de Trading Avanzado
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
COLUMNAS_RADAR_TOTW = ["Jugador", "Media", "Precio_Actual", "Fuente", "Break_Even_5%", "Precio_Lazy_1h", "Riesgo", "Enlace_Futbin"]

def ahora_arg():
    """Devuelve siempre la fecha y hora exacta de Argentina (UTC-3)."""
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
            print(f"⚠️ Aviso leyendo {ARCHIVO_CRED}: {e}")
    return None

def subir_dataframe_gsheets(sh, nombre_hoja, df_subir, max_filas=25000):
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
    df_evos, df_totw, df_domino
):
    gc = obtener_cliente_gsheets()
    if not gc:
        print("ℹ️ Guardado directo en Lista_Jugadores_FC27.xlsx (GitHub Auto-Sync activo).")
        return
    try:
        print(f"\n☁️ Subiendo base de datos actualizada a Google Sheets ('{NOMBRE_SHEET}')...")
        sh = gc.open(NOMBRE_SHEET)
        subir_dataframe_gsheets(sh, "Historial_Tendencias", df_final, max_filas=25000)
        df_ia_nube = pd.DataFrame({"Fecha": [fecha_actual], "Reporte": [texto_limpio], "Oraculo": [oraculo_txt]})
        subir_dataframe_gsheets(sh, "Recomendacion_IA", df_ia_nube, max_filas=100)
        subir_dataframe_gsheets(sh, "Mi_Cartera", df_cartera, max_filas=1000)
        subir_dataframe_gsheets(sh, "Mis_Alertas", df_alertas, max_filas=500)
        if not df_pnl_excel.empty:
            subir_dataframe_gsheets(sh, "Dashboard_Financiero", df_pnl_excel, max_filas=200)
        if not df_fodder.empty:
            subir_dataframe_gsheets(sh, "Radar_Fodder", df_fodder, max_filas=200)
        if not df_evos.empty:
            subir_dataframe_gsheets(sh, "Radar_Evos", df_evos, max_filas=500)
        df_totw_sync = df_totw if (df_totw is not None and not df_totw.empty) else pd.DataFrame(columns=COLUMNAS_RADAR_TOTW)
        subir_dataframe_gsheets(sh, "Radar_TOTW", df_totw_sync, max_filas=500)
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

UMBRAL_DESFASE_VERSION_PCT = 45.0

def sin_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn')

def extraer_media_exacta(texto):
    """Extrae una media válida (45-99) de un nombre formateado. Devuelve '' si falta o es '??'."""
    if not texto:
        return ""
    m = re.search(r'Media\s*(\d{2})\b', str(texto))
    if m:
        val = int(m.group(1))
        if 45 <= val <= 99:
            return str(val)
    return ""

def categoria_version_carta(texto_nombre_o_tipo):
    """Clasifica estrictamente la carta en 'Normal/Oro', 'TOTW/IF' o 'Promo/Especial' para impedir cruces falsos."""
    t = sin_tildes(str(texto_nombre_o_tipo)).upper()
    if any(k in t for k in ["TOTW", "IN-FORM", "INFORM", "TEAM OF THE WEEK", "RARITY-TOTW", "RARITY_TOTW"]) or re.search(r'(?:-\s*IF\b|\(IF\b|\bIF\))', t):
        return "TOTW/IF"
    palabras_promo = [
        "ICON", "HERO", "EVO", "EVOLUTION", "PROMO", "SPECIAL", "RTTK", "RTTF",
        "TRAILBLAZER", "CENTURION", "THUNDER", "TOTY", "TOTS", "FUTTIES",
        "WINTER", "WILDCARD", "FUTURE STAR", "BIRTHDAY", "FANTASY", "POTM",
        "FLASHBACK", "SHOWDOWN", "RUSH", "SHAPESHIFTER", "GRASSROOT", "IMMORTAL",
        "DREAMCHASER", "NUMERO", "TRACK STAR", "GLOBETROTTER", "SUCCESSION",
        "CORNERSTONE", "SCREAM", "FC PRO", "DYNASTY", "RADIOACTIVE", "GOLAZO",
        "PATH TO GLORY", "MAKE YOUR MARK", "GREATS OF THE GAME", "END OF AN ERA",
        "ONES TO WATCH", "OTW", "RULEBREAKER", "HEADLINER", "SIGNATURE",
        "OUT OF POSITION", "VERSUS", "MOMENTS", "STORYLINE", "FOUNDATIONS",
        "WORLD TOUR", "MODE MASTERY", "UCL", "UEL", "UECL", "LIBERTADORES", "SUDAMERICANA"
    ]
    if any(k in t for k in palabras_promo):
        return "Promo/Especial"
    return "Normal/Oro"

def normalizar_nombre_completo(nombre_completo):
    """Normaliza el nombre completo sin tildes ni sufijos de paréntesis para emparejamiento exacto."""
    nombre_base = str(nombre_completo).split(' (')[0]
    limpio = sin_tildes(nombre_base).lower()
    limpio = re.sub(r'\bjr\.?\b', 'junior', limpio)
    limpio = re.sub(r'\bsr\.?\b', 'senior', limpio)
    limpio = re.sub(r'\bst\.?\b', 'saint', limpio)
    palabras = [p for p in re.findall(r'[a-z]+', limpio) if len(p) >= 2]
    if not palabras:
        palabras = re.findall(r'[a-z]+', limpio)
    return "_".join(palabras)

def clave_cruce(nombre_completo):
    """Clave primaria quirúrgica: nombre completo normalizado + media exacta + categoría de carta."""
    media = extraer_media_exacta(nombre_completo)
    if not media:
        return ""
    nom_norm = normalizar_nombre_completo(nombre_completo)
    if not nom_norm:
        return ""
    cat_ver = categoria_version_carta(nombre_completo)
    return f"{nom_norm}|{media}|{cat_ver}"

def clave_cruce_apellido_media(nombre_completo):
    """Clave secundaria controlada: inicial + apellido + media exacta + categoría (nunca sin media ni cruzando categorías)."""
    media = extraer_media_exacta(nombre_completo)
    if not media:
        return ""
    nom_norm = normalizar_nombre_completo(nombre_completo)
    palabras = [p for p in nom_norm.split('_') if p]
    if not palabras:
        return ""
    cat_ver = categoria_version_carta(nombre_completo)
    if len(palabras) >= 2:
        identificador = f"{palabras[0][0]}_{palabras[-1]}"
    else:
        identificador = palabras[0]
    return f"{identificador}|{media}|{cat_ver}"

# ==========================================
# MATEMÁTICA FINANCIERA EA TAX, LAZY & STOP-LOSS (SOPORTE 15M)
# ==========================================
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

def es_multiplo_escalon_ea(precio, presupuesto_max=PRESUPUESTO_MAXIMO):
    """Valida que un precio sea estrictamente múltiplo de los escalones oficiales del mercado EA FC."""
    try:
        p = int(precio)
    except Exception:
        return False
    if not (200 <= p <= presupuesto_max):
        return False
    return (p % paso_escalon_ea(p)) == 0

def redondear_escalon_ea(precio, hacia_arriba=False):
    if precio <= 0:
        return 0
    paso = paso_escalon_ea(precio)
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

def calcular_precio_lazy(precio_referencia, precio_pagado=0, max_hist=0, muestras=0, margen_lazy=MARGEN_LAZY_BUYER):
    """Calcula el precio Lazy respetando EA Tax (5%), rentabilidad mínima y techo histórico firme (max_hist)."""
    if precio_referencia <= 0:
        return 0
    base_lazy = precio_referencia * margen_lazy
    if precio_pagado > 0:
        min_rentable = (precio_pagado * 1.10) / 0.95
        base_lazy = max(base_lazy, min_rentable)
    piso_empate = calcular_break_even(precio_pagado if precio_pagado > 0 else precio_referencia)
    lazy_ea = redondear_escalon_ea(base_lazy, hacia_arriba=True)

    # Respetar techo firme reciente (no sugerir Lazy por encima de max_hist si hay historial suficiente)
    if muestras >= 3 and max_hist > 0 and lazy_ea > max_hist:
        techo_ea = redondear_escalon_ea(max_hist, hacia_arriba=False)
        if techo_ea >= piso_empate:
            return techo_ea
        return piso_empate

    return max(lazy_ea, piso_empate)

def calcular_stop_loss_sugerido(precio_pagado):
    if precio_pagado <= 0:
        return 0
    venta_emergencia = (precio_pagado * (1.0 + (STOP_LOSS_PCT / 100.0))) / 0.95
    return redondear_escalon_ea(venta_emergencia, hacia_arriba=False)

def ganancia_neta_minima_real(precio_actual):
    """Umbral mínimo de ganancia neta proyectada Post-Tax (5%) para eliminar ruido en cartas baratas."""
    if precio_actual < 5000:
        return 350
    elif precio_actual <= 25000:
        return 750
    return 1500

def calcular_horas_inmovilizacion(fecha_compra_str):
    if not fecha_compra_str or str(fecha_compra_str).strip() == "":
        return 0.0
    try:
        dt_compra = datetime.strptime(str(fecha_compra_str).strip(), "%Y-%m-%d %H:%M")
        delta = ahora_arg() - dt_compra
        return max(0.0, round(delta.total_seconds() / 3600.0, 1))
    except Exception:
        return 0.0

def _limpiar_ruido_dom_precio(texto):
    """Elimina tokens típicos de estadísticas in-game, likes, contratos, medidas e IDs antes de parsear monedas."""
    t = str(texto or "")
    t = re.sub(r'\b(?:IGS|BGS|PAC|SHO|PAS|DRI|DEF|PHY|OVR|RATING|WF|SM|SKILL|CONTRACTS?|LIKES?|VIEWS?|COMMENTS?|XP|PTS|ID)\s*[:#]?\s*\d+(?:[.,]\d+)?[KkMm]?\b', ' ', t, flags=re.IGNORECASE)
    t = re.sub(r'\b\d+(?:[.,]\d+)?[KkMm]?\s*(?:IGS|BGS|LIKES?|VIEWS?|COMMENTS?|CONTRACTS?|CM|KG|FT|LB|XP|PTS|%)\b', ' ', t, flags=re.IGNORECASE)
    t = re.sub(r'#\d+\b', ' ', t)
    return t

def _parsear_valor_monedas_km(fuente_texto, presupuesto_max):
    """Extrae precios K/M, con comas o enteros exigiendo estrictamente múltiplos válidos de los escalones EA FC."""
    if not fuente_texto:
        return 0
    texto_norm = _limpiar_ruido_dom_precio(fuente_texto).strip()
    if not texto_norm:
        return 0

    # 1. Sufijos K y M (ej. 14.5K -> 14500, 1.25M -> 1250000)
    precios_km = re.findall(r'(?<!\d)(\d+(?:[.,]\d+)?)\s*([KkMm])\b', texto_norm)
    for num_s, unidad in precios_km:
        num_clean = float(num_s.replace(',', '.'))
        mult = 1000000 if unidad.upper() == 'M' else 1000
        val = int(round(num_clean * mult))
        if val > presupuesto_max:
            return -1
        if es_multiplo_escalon_ea(val, presupuesto_max):
            return val
        # Tolerancia controlada solo para abreviaturas K con 1 decimal en rango 10k-50k (ej. 14.3K -> 14250)
        if unidad.upper() == 'K' and 10000 <= val < 50000 and ('.' in num_s or ',' in num_s):
            ajustado = int(round(val / 250.0) * 250)
            if abs(ajustado - val) <= 50 and es_multiplo_escalon_ea(ajustado, presupuesto_max):
                return ajustado

    # 2. Números con separadores de miles por comas (ej. 14,500 o 1,250,000) -> estricto múltiplo EA
    precios_comas = re.findall(r'\b(\d{1,3}(?:,\d{3})+)\b', texto_norm)
    for p_c in precios_comas:
        val = int(p_c.replace(',', ''))
        if val > presupuesto_max:
            return -1
        if es_multiplo_escalon_ea(val, presupuesto_max):
            return val

    # 3. Enteros directos -> estricto múltiplo de escalón EA (descarta IGS como 2150, contratos e IDs)
    enteros = re.findall(r'\b(\d{3,8})\b', texto_norm)
    for ent in enteros:
        val = int(ent)
        if val > presupuesto_max:
            continue
        if es_multiplo_escalon_ea(val, presupuesto_max):
            return val
    return 0

def extraer_precio_robusto(texto_precio, texto_fila, presupuesto_max):
    """Parsea primero el nodo aislado de precio y solo recurre a texto_fila si el nodo estaba vacío."""
    if texto_precio:
        val_p = _parsear_valor_monedas_km(texto_precio, presupuesto_max)
        if val_p == -1:
            return 0
        if val_p > 0:
            return val_p
    if texto_fila and texto_fila != texto_precio:
        val_f = _parsear_valor_monedas_km(texto_fila, presupuesto_max)
        if val_f > 0:
            return val_f
    return 0

def extraer_precio_carta_unica_gg(texto_card, presupuesto_max, texto_precio=""):
    """Extrae el precio únicamente dentro de la caja individual de cada carta en FUT.GG validando escalón EA."""
    if texto_precio:
        val_p = _parsear_valor_monedas_km(texto_precio, presupuesto_max)
        if val_p == -1:
            return 0
        if val_p > 0:
            return val_p
    if not texto_card:
        return 0
    val = _parsear_valor_monedas_km(texto_card, presupuesto_max)
    return val if val > 0 else 0

def calcular_estadisticas_historicas(df_viejo):
    stats = {}
    if df_viejo is None or df_viejo.empty:
        return stats
    col_key = 'Enlace_Optimo' if 'Enlace_Optimo' in df_viejo.columns else ('Enlace_Futbin' if 'Enlace_Futbin' in df_viejo.columns else None)
    if not col_key:
        return stats
    try:
        df_cp = df_viejo.copy()
        df_cp['Precio_Actual'] = pd.to_numeric(df_cp['Precio_Actual'], errors='coerce').fillna(0)
        df_valido = df_cp[df_cp['Precio_Actual'] > 0]
        agrupado = df_valido.groupby(col_key)['Precio_Actual'].agg(['min', 'max', 'mean', 'count'])
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
# MOTOR DE CLASIFICACIÓN DE TRADING (BASE EXCLUSIVA FUT.GG)
# ==========================================
def diagnosticar_senal_trading(p_act, p_ant, min_hist, max_hist, prom_hist, muestras):
    if p_act <= 0:
        return "⚪ Rango Neutro"

    var_reciente = (((p_act - p_ant) / p_ant) * 100.0) if p_ant > 0 else 0.0
    desc_vs_prom = (((p_act - prom_hist) / prom_hist) * 100.0) if prom_hist > 0 else 0.0
    amplitud_pct = (((max_hist - min_hist) / min_hist) * 100.0) if min_hist > 0 else 0.0

    # 1. Protección prioritaria contra Cuchillo Cayendo (nunca etiquetar compra en caída libre)
    if (var_reciente <= -22.0 and p_act < min_hist and muestras >= 3) or (var_reciente <= -25.0 and muestras >= 2):
        return "⚠️ Cuchillo Cayendo (Esperar Piso)"

    # 2. Exigencia de ganancia neta mínima real Post-Tax (evita ruido de monedas bajas)
    min_ganancia = ganancia_neta_minima_real(p_act)
    ref_rebote_prom = min(prom_hist, max_hist) if (prom_hist > 0 and max_hist > 0) else prom_hist
    ganancia_proy_prom = int(round((ref_rebote_prom * 0.95) - p_act)) if ref_rebote_prom > 0 else 0
    ganancia_proy_ant = int(round((p_ant * 0.95) - p_act)) if p_ant > 0 else 0

    if muestras >= 2 and amplitud_pct >= 12.0 and desc_vs_prom <= -5.0 and var_reciente >= -18.0 and ganancia_proy_prom >= min_ganancia:
        return "⚡ Flipping Rápido (Alta Liquidez)"

    if muestras >= 2 and p_act <= int(min_hist * 1.03) and desc_vs_prom <= -6.0 and var_reciente > -22.0 and ganancia_proy_prom >= min_ganancia:
        return "💎 Rebote en Suelo Firme"

    if var_reciente <= -5.0 and max(ganancia_proy_prom, ganancia_proy_ant) >= min_ganancia:
        return "🟢 Caída Moderada (Puja)"

    ganancia_subida = int(round((p_act * 0.95) - p_ant)) if p_ant > 0 else 0
    if var_reciente >= 6.0 and ganancia_subida >= min_ganancia:
        return "🔴 En Subida (Tomar Ganancia)"

    return "⚪ Rango Neutro"

def obtener_radar_ciclo_futchampions(datos_jugadores):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    ahora = ahora_arg()
    dia_idx = ahora.weekday()
    dia_nombre = dias[dia_idx]
    hora = ahora.hour

    pautas_wl = {
        0: ("🟢 FASE ACUMULACIÓN POST-WL (LUNES)", "Sell-off terminado. Comprar cartas Meta en piso; suben miércoles/jueves."),
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
        enl_m = m.get("Enlace_Optimo", m.get("Enlace_Futbin", ""))
        puja_m = m.get("Puja_Max_10%_ROI", calcular_puja_maxima_segura(m["Precio_Actual"]))
        lazy_m = m.get("Precio_Lazy_Sugerido", calcular_precio_lazy(m["Precio_Actual"]))
        txt_meta += f"   • [*{nom_m}*]({enl_m}): `{m['Precio_Actual']:,}` (🔨 Pujar: `{puja_m:,}` | 🎯 Lazy: `{lazy_m:,}`)\n"

    if not txt_meta:
        txt_meta = "   • _Cartas Meta firmes en este corte (usar Lazy Listing en cartera)._\n"

    return (
        f"🏆 *CICLO SEMANAL FUT CHAMPIONS ({dia_nombre} {hora:02d}:00 hs)*\n━━━━━━━━━━━━━━━━━━\n"
        f"🧭 Fase: *{titulo_fase}*\n"
        f"💡 Estrategia: _{consejo_wl}_\n"
        f"🎯 *Top Objetivos Meta en Descuento:*\n{txt_meta}\n"
    )

def detectar_links_domino(datos_gg_procesados, df_evos):
    if not datos_gg_procesados:
        return "🔗 *RASTREADOR DE LINKS & EFECTO DOMINÓ*\n━━━━━━━━━━━━━━━━━━\nSin datos suficientes.\n\n", pd.DataFrame()

    cartas_motor = [
        j for j in datos_gg_procesados
        if float(j.get("Variacion_Bruta_%", 0)) >= 20.0
    ]

    if not cartas_motor:
        return "🔗 *RASTREADOR DE LINKS & EFECTO DOMINÓ*\n━━━━━━━━━━━━━━━━━━\nSin explosiones base (>20%) que disparen segunda ola en este corte.\n\n", pd.DataFrame()

    medias_calientes = set()
    for cm in cartas_motor[:8]:
        m_match = extraer_media_exacta(cm.get("Jugador", ""))
        if m_match:
            medias_calientes.add(int(m_match))

    if not medias_calientes:
        medias_calientes = {82, 83, 84}

    candidatos_domino = []
    for j in datos_gg_procesados:
        p_act = int(j.get("Precio_Actual", 0))
        var_b = float(j.get("Variacion_Bruta_%", 0))
        if not (700 <= p_act <= 9500 and -10.0 <= var_b <= 3.0):
            continue
        m_match = extraer_media_exacta(j.get("Jugador", ""))
        med_j = int(m_match) if m_match else 0
        if med_j in medias_calientes:
            puja_d = calcular_puja_maxima_segura(p_act, 0.10)
            lazy_d = calcular_precio_lazy(
                p_act,
                max_hist=int(j.get("Precio_Max_Hist", 0) or 0),
                muestras=int(j.get("Muestras_Hist", 0) or 0)
            )
            candidatos_domino.append({
                "Jugador_Dormido": j["Jugador"].split(" (")[0] + f" (Media {med_j})",
                "Media": med_j,
                "Precio_Actual": p_act,
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
            f"👤 [*{row['Jugador_Dormido']}*]({row['Enlace']}) — Piso: `{row['Precio_Actual']:,} mnds`\n"
            f"   └ 🔨 Pujar: `{row['Puja_Sugerida']:,}` | 🎯 Salida 2da Ola / Lazy: `{row['Target_Lazy_2da_Ola']:,}`\n\n"
        )
    return txt_dom, df_domino

def armar_resumen_flipping_gangas(datos_finales):
    if not datos_finales:
        return "⚠️ Sin datos cargados para filtrar Flipping."

    gangas_flip = [
        j for j in datos_finales
        if any(k in str(j.get("Senal_Trading", "")) for k in ["⚡", "💎", "🟢"])
        and "⚠️" not in str(j.get("Senal_Trading", ""))
        and float(j.get("Variacion_Bruta_%", 0)) < 0
    ]
    gangas_flip = sorted(gangas_flip, key=lambda x: float(x.get("Variacion_Bruta_%", 0)))[:6]

    trampas = [
        j for j in datos_finales
        if "⚠️" in str(j.get("Senal_Trading", ""))
    ][:3]

    txt = "⚡ *RADAR DE FLIPPING RÁPIDO & LAZY LISTING (FUT.GG)*\n━━━━━━━━━━━━━━━━━━\n"
    if gangas_flip:
        for g in gangas_flip:
            nom = g["Jugador"].split(" (")[0]
            enl = g.get("Enlace_Optimo", g.get("Enlace_Futbin", ""))
            p_act = int(g["Precio_Actual"])
            puja = int(g.get("Puja_Max_10%_ROI", calcular_puja_maxima_segura(p_act)))
            emp = int(g.get("Break_Even_5%", calcular_break_even(p_act)))
            lazy = int(g.get("Precio_Lazy_Sugerido", calcular_precio_lazy(p_act)))
            senal = g.get("Senal_Trading", "🟢 Compra")
            txt += (
                f"👤 [*{nom}*]({enl}) — `{p_act:,} mnds` (`{g['Variacion_Bruta_%']}%`)\n"
                f"   └ 🏷️ Señal: *{senal}*\n"
                f"   └ 🔨 *Puja Máx:* `{puja:,}` | ⚖️ *Empate:* `{emp:,}` | 🎯 *Lazy 1h:* `{lazy:,}`\n\n"
            )
    else:
        txt += "Sin caídas de flipping inmediato en este corte.\n\n"

    if trampas:
        txt += "🛑 *TRAMPAS DE VALOR (CUCHILLOS CAYENDO - NO ENTRAR):*\n"
        for t in trampas:
            nom_t = t["Jugador"].split(" (")[0]
            txt += f"• *{nom_t}* (`{t['Precio_Actual']:,} mnds`, `{t['Variacion_Bruta_%']}%`) — Rompió soporte sin rebote.\n"
        txt += "\n"

    return txt

# ==========================================
# PROCESADOR EXCLUSIVO FUT.GG (HASTA 10,000 CARTAS)
# ==========================================
def procesar_datos_futgg(datos_gg, historial_precios, stats_historicas, fecha_actual):
    """
    Procesa hasta 10,000 cartas extraídas exclusivamente desde FUT.GG, calculando métricas
    financieras (EA Tax 5%, Puja Máx, Lazy 1h), bandas históricas y señales de trading.
    Mantiene el alias Enlace_Futbin apuntando a Enlace_Optimo (FUT.GG) por compatibilidad histórica.
    """
    datos_procesados = []
    vistos_enl = set()

    for g in (datos_gg or []):
        enlace_gg = str(g.get("Enlace_GG", "") or "").strip()
        if not enlace_gg or enlace_gg in vistos_enl:
            continue

        p_act = int(g.get("Precio_FUTGG", 0) or 0)
        media_g = extraer_media_exacta(g.get("Jugador_GG", "")) or (
            str(g.get("Media", "")) if str(g.get("Media", "")).isdigit() and 45 <= int(g.get("Media", "0")) <= 99 else ""
        )
        if not es_multiplo_escalon_ea(p_act, PRESUPUESTO_MAXIMO) or not media_g:
            continue

        vistos_enl.add(enlace_gg)
        etiqueta_gg = str(g.get("Tipo_Carta", "") or "Normal").strip()
        nombre_base = str(g.get("Jugador_GG", "Jugador")).split(" (")[0].strip()
        nombre_jug_final = f"{nombre_base} (Media {media_g} - {etiqueta_gg} - FUT.GG)"

        es_nueva = enlace_gg not in historial_precios
        precio_anterior = int(historial_precios.get(enlace_gg, p_act) or p_act)
        var_b = (((p_act - precio_anterior) / precio_anterior) * 100.0) if precio_anterior > 0 else 0.0
        ganancia_neta = int(round((p_act * 0.95) - precio_anterior)) if precio_anterior > 0 and p_act > precio_anterior else 0
        var_n = (ganancia_neta / precio_anterior) * 100.0 if precio_anterior > 0 else 0.0

        st_h = stats_historicas.get(enlace_gg) or {"min": p_act, "max": p_act, "prom": p_act, "count": 1}
        min_hist = min(st_h["min"], p_act)
        max_hist = max(st_h["max"], p_act)
        prom_hist = st_h["prom"]
        muestras = int(st_h["count"])
        descuento_vs_prom = (((p_act - prom_hist) / prom_hist) * 100.0) if prom_hist > 0 else 0.0

        if muestras >= 2 and p_act <= min_hist and p_act < prom_hist:
            estado_hist = "💎 MÍNIMO HISTÓRICO"
        elif descuento_vs_prom <= -12.0:
            estado_hist = f"📉 {descuento_vs_prom:.0f}% vs Prom"
        else:
            estado_hist = "Normal"

        senal_tr = diagnosticar_senal_trading(p_act, precio_anterior, min_hist, max_hist, prom_hist, muestras)

        if "⚠️" in senal_tr:
            accion = "⚠️ EVITAR (CAÍDA)"
        elif any(k in senal_tr for k in ("⚡", "💎", "🟢")):
            accion = "🟢 COMPRAR"
        elif "🔴" in senal_tr or (var_n >= 5.0 and ganancia_neta >= ganancia_neta_minima_real(precio_anterior)):
            accion = "🔴 VENDER"
        elif es_nueva:
            accion = "🆕 NUEVO"
        else:
            accion = "⚪ MANTENER"

        datos_procesados.append({
            "Fecha": fecha_actual,
            "Jugador": nombre_jug_final,
            "Categoria_Carta": categoria_version_carta(nombre_jug_final),
            "Precio_Anterior": precio_anterior,
            "Precio_Actual": p_act,
            "Precio_Consola": p_act,
            "Precio_Min_Hist": min_hist,
            "Precio_Max_Hist": max_hist,
            "Precio_Prom_Hist": prom_hist,
            "Muestras_Hist": muestras,
            "Estado_Historico": estado_hist,
            "Senal_Trading": senal_tr,
            "Puja_Max_10%_ROI": calcular_puja_maxima_segura(p_act, 0.10),
            "Break_Even_5%": calcular_break_even(p_act),
            "Precio_Lazy_Sugerido": calcular_precio_lazy(p_act, max_hist=max_hist, muestras=muestras),
            "Variacion_Bruta_%": round(var_b, 2),
            "Beneficio_Neto_Mnds": round(ganancia_neta),
            "Variacion_Neta_%": round(var_n, 2),
            "Accion": accion,
            "Enlace_Futbin": enlace_gg,
            "Enlace_Optimo": enlace_gg
        })

    print(f"⚡ Base Exclusiva FUT.GG procesada: {len(datos_procesados):,} cartas válidas listas.")
    return datos_procesados

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

def precio_mercado_seguro_carta(match, fallback_precio=0):
    """Devuelve el precio de liquidación vigente en FUT.GG (o el fallback si no hay cotización)."""
    if not match:
        return int(fallback_precio)
    p_act = int(float(match.get("Precio_Actual", 0) or 0))
    return p_act if p_act > 0 else int(fallback_precio)

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
    if "Cantidad" not in df_c.columns:
        df_c["Cantidad"] = 1

    df_c["Estrategia"] = df_c["Estrategia"].fillna("").astype(object)
    df_c["Cantidad"] = pd.to_numeric(df_c["Cantidad"], errors="coerce").fillna(1).clip(lower=1).astype(int)
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

    ventas_brutas_tot = int((vendidos["Precio_Venta"] * vendidos["Cantidad"]).sum()) if not vendidos.empty else 0
    vendidos_con_precio = vendidos[vendidos["Precio_Venta"] > 0] if not vendidos.empty else pd.DataFrame()
    costo_adq_tot = int((vendidos_con_precio["Precio_Pagado"] * vendidos_con_precio["Cantidad"]).sum()) if not vendidos_con_precio.empty else 0
    ea_tax_tot = int(round(ventas_brutas_tot * 0.05))
    ingreso_neto_tot = ventas_brutas_tot - ea_tax_tot
    contrib_marginal_tot = ingreso_neto_tot - costo_adq_tot
    margen_sobre_ventas = (contrib_marginal_tot / ventas_brutas_tot * 100) if ventas_brutas_tot > 0 else 0.0
    roi_sobre_costo = (contrib_marginal_tot / costo_adq_tot * 100) if costo_adq_tot > 0 else 0.0

    capital_inmovilizado = int((activos["Precio_Pagado"] * activos["Cantidad"]).sum()) if not activos.empty else 0
    vnr_total_activo = 0
    cartas_estancadas = 0
    cartas_stoploss = 0
    suma_horas_activas = 0.0

    for _, r_act in activos.iterrows():
        j_nom = sin_tildes(str(r_act["Jugador_Exacto"])).lower().strip()
        p_pag = float(r_act["Precio_Pagado"])
        cant_act = max(1, int(r_act.get("Cantidad", 1)))
        hrs = calcular_horas_inmovilizacion(r_act.get("Fecha_Compra", ""))
        suma_horas_activas += hrs
        if hrs >= HORAS_TOPE_INMOVILIZACION:
            cartas_estancadas += 1

        cands_m = [j for j in (datos_recientes or []) if j_nom in sin_tildes(j.get("Jugador", "")).lower()]
        match = min(cands_m, key=lambda j: abs(int(j.get("Precio_Actual", 0)) - p_pag)) if cands_m else None
        p_ref = precio_mercado_seguro_carta(match, fallback_precio=p_pag)
        vnr_carta = int(round(p_ref * 0.95))

        roi_latente = ((vnr_carta - p_pag) / p_pag * 100.0) if p_pag > 0 else 0.0
        if roi_latente <= STOP_LOSS_PCT:
            cartas_stoploss += 1
        vnr_total_activo += (vnr_carta * cant_act)

    prom_dias_stock = round((suma_horas_activas / len(activos)) / 24.0, 1) if not activos.empty else 0.0
    resultado_tenencia_no_realizado = vnr_total_activo - capital_inmovilizado
    liquidez_disponible = max(0, 200000 - capital_inmovilizado)

    todas_est = sorted(list(set(estrategias_base + [str(x) for x in df_c["Estrategia"].unique() if str(x).strip()])))
    for est in todas_est:
        if not est:
            continue
        v_est = vendidos_con_precio[vendidos_con_precio["Estrategia"] == est] if not vendidos_con_precio.empty else pd.DataFrame()
        a_est = activos[activos["Estrategia"] == est]
        if v_est.empty and a_est.empty:
            continue

        vb = int((v_est["Precio_Venta"] * v_est["Cantidad"]).sum()) if not v_est.empty else 0
        ca = int((v_est["Precio_Pagado"] * v_est["Cantidad"]).sum()) if not v_est.empty else 0
        tax = int(round(vb * 0.05))
        cm = (vb - tax) - ca
        roi_e = round((cm / ca * 100), 1) if ca > 0 else 0.0
        aciertos = len(v_est[(v_est["Precio_Venta"] * 0.95 - v_est["Precio_Pagado"]) > 0]) if not v_est.empty else 0
        tasa_e = round((aciertos / len(v_est) * 100), 1) if len(v_est) > 0 else 0.0
        cap_act_e = int((a_est["Precio_Pagado"] * a_est["Cantidad"]).sum()) if not a_est.empty else 0

        resumen_estrategias.append({
            "Estrategia": est,
            "Trades_Cerrados": len(v_est),
            "Ventas_Brutas": vb,
            "Costo_Adquisicion": ca,
            "Costo_Variable_EA_Tax_5%": tax,
            "Contribucion_Marginal_Neta": cm,
            "ROI_Real_%": roi_e,
            "Efectividad_%": tasa_e,
            "Cartas_Activas": int(a_est["Cantidad"].sum()) if not a_est.empty else 0,
            "Capital_Activo": cap_act_e
        })

    df_estrategias = pd.DataFrame(resumen_estrategias)

    df_pnl = pd.DataFrame([
        {"Rubro_Contable": "(+) Ingresos Brutos por Ventas", "Monto_Monedas": ventas_brutas_tot, "Incidencia_%": 100.0 if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(-) Costo Variable Comercial (EA Tax 5%)", "Monto_Monedas": -ea_tax_tot, "Incidencia_%": -5.0 if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(=) Ingreso Neto de Realización", "Monto_Monedas": ingreso_neto_tot, "Incidencia_%": 95.0 if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(-) Costo de Adquisición de Cartas Vendidas (CMV)", "Monto_Monedas": -costo_adq_tot, "Incidencia_%": round((-costo_adq_tot / ventas_brutas_tot * 100), 1) if ventas_brutas_tot > 0 else 0.0},
        {"Rubro_Contable": "(=) CONTRIBUCIÓN MARGINAL NETA (GANANCIA REALIZADA)", "Monto_Monedas": contrib_marginal_tot, "Incidencia_%": round(margen_sobre_ventas, 1)},
        {"Rubro_Contable": "📊 Inventario Activo (Costo Histórico)", "Monto_Monedas": capital_inmovilizado, "Incidencia_%": 0.0},
        {"Rubro_Contable": "📈 Inventario Activo a VNR FUT.GG (Neto 5% Tax)", "Monto_Monedas": vnr_total_activo, "Incidencia_%": 0.0},
        {"Rubro_Contable": "⚖️ Resultado por Tenencia (No Realizado)", "Monto_Monedas": resultado_tenencia_no_realizado, "Incidencia_%": 0.0},
        {"Rubro_Contable": "⏱️ Antigüedad Promedio del Stock Activo (Días)", "Monto_Monedas": prom_dias_stock, "Incidencia_%": float(cartas_estancadas)},
        {"Rubro_Contable": "💵 Liquidez Disponible (Caja Operativa)", "Monto_Monedas": liquidez_disponible, "Incidencia_%": 0.0}
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
        f"• Antigüedad Promedio Stock: `{prom_dias_stock} días` (`{cartas_estancadas}` estancadas >72h | `{cartas_stoploss}` en Stop-Loss)\n\n"
        f"🏷️ *RENTABILIDAD POR UNIDAD DE NEGOCIO:*\n"
        f"{txt_est_tg}"
    )

    return reporte_pnl_tg, df_pnl, df_estrategias

# ==========================================
# MÓDULO 2: ORÁCULO DE CONTENIDO (15:00 ARG)
# ==========================================
def escanear_oraculo_contenido(page, datos_finales, ocultar=False):
    print("\n🔮 Ejecutando Oráculo de Contenido (Base Exclusiva FUT.GG + Agenda 15:00 hs ARG)...")
    tendencias_web = []

    cerrar_al_final = False
    if page is None:
        page = crear_navegador_optimizado(ocultar=ocultar)
        cerrar_al_final = True

    try:
        page.get("https://www.fut.gg/players/?sort=popular")
        if page.ele('xpath://a[contains(@href, "/players/")]', timeout=3):
            page.wait(0.35)
            _detener_carga_segura(page)
            script_pop = """
            let pop = [];
            let vistos = new Set();
            let enlaces = document.querySelectorAll('a[href*="/players/"]');
            for (let a of enlaces) {
                let href = (a.getAttribute('href') || '').split('?')[0];
                let m = href.match(/\\/players\\/\\d+-([a-zA-Z0-9\\-]+)/);
                if (!m) continue;
                let nom = m[1].replace(/-/g, ' ');
                if (!vistos.has(nom)) {
                    vistos.add(nom);
                    pop.push(nom);
                }
                if (pop.length >= 10) break;
            }
            return pop;
            """
            pop_extraidos = page.run_js(script_pop) or []
            for p_nom in pop_extraidos[:8]:
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
                subidas_fuertes.append(f"{j['Jugador']} ({j['Precio_Actual']:,} mnds, +{j['Variacion_Bruta_%']}%)")

        for med, lim in sorted(REFERENCIA_FODDER.items()):
            cands = [x for x in datos_finales if f"(Media {med}" in str(x.get("Jugador", "")) and int(x.get("Precio_Actual", 0) or 0) > 0]
            if cands:
                min_c = min(cands, key=lambda x: int(x["Precio_Actual"]))
                fodder_barato.append(f"Media {med}: {min_c['Jugador'].split(' (')[0]} a {int(min_c['Precio_Actual']):,} mnds")

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
    Actuá como el Oráculo de Contenido y Filtraciones (Leaks) de EA Sports FC 27 para el MERCADO DE PC (Base Exclusiva FUT.GG de {len(datos_finales):,} jugadores).
    Hoy es {dia_hoy}, hora actual en Argentina: {hora_str} hs (El contenido de EA sale a las 15:00 hs ARG).
    Pauta histórica del día ({dia_hoy}): {contexto_dia}
    Cartas más buscadas hoy (Hype/Leaks de Evoluciones): {', '.join(tendencias_web) if tendencias_web else 'Flujo estándar'}
    Picos de demanda detectados en vivo en PC: {', '.join(subidas_fuertes) if subidas_fuertes else 'Sin picos abruptos'}
    Precios piso actuales de Fodder (FUT.GG): {', '.join(fodder_barato)}

    Redactá un REPORTE DEL ORÁCULO DE CONTENIDO breve, directo y accionable (sin usar tablas '|'):
    1. 🔮 RADAR DE LAS 15:00 HS ARG ({dia_hoy}): Qué contenido/SBC se espera hoy y qué requisitos suele pedir EA.
    2. 🛒 LAS 3 COMPRAS DICTADAS POR EL ORÁCULO: Indicá exactamente qué 3 cartas o medias de la lista de Fodder conviene comprar AHORA MISMO antes de que salgan los SBCs.
    3. 🧬 ALERTA DE HYPE / EVOLUCIONES: Qué hacer con las cartas en tendencia ({', '.join(tendencias_web[:4]) if tendencias_web else 'cartas meta de bajo presupuesto'}).
    """

    dictamen_oraculo, _ = llamar_gemini_turbo(prompt_oraculo)
    if not dictamen_oraculo:
        dictamen_oraculo = (
            f"• *Agenda {dia_hoy} (15:00 hs ARG):* {contexto_dia}\n"
            f"• *Fodder más barato en FUT.GG:* {', '.join(fodder_barato[:4])}\n"
            f"• *Recomendación:* Acumular medias en piso e IFs por debajo de 14k antes de las 14:45 hs."
        )

    mensaje_oraculo = (
        f"🔮 *ORÁCULO DE CONTENIDO & LEAKS FC27 (FUT.GG)*\n"
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
        return f"🌡️ *TERMÓMETRO FUT.GG ({total_padron:,} cartas):* `⚖️ ESTABLE (0.0%)`\n_Recién inicializando base comparativa._\n\n"

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
        f"🌡️ *TERMÓMETRO DEL MERCADO FUT.GG ({total_padron:,} Cartas)*\n━━━━━━━━━━━━━━━━━━\n"
        f"📊 Estado: *{estado}*\n"
        f"📉 Caída: `{bajando}` | 📈 Subida: `{subiendo}` | 💎 Suelo: `{minimos_hoy}` | ⚡ Flipping: `{flipping_hoy}`\n"
        f"💡 {consejo}\n\n"
    )

def analizar_fodder(datos_jugadores):
    if not datos_jugadores:
        return "Sin datos de mercado para calcular Fodder.\n", pd.DataFrame()
    filas_fodder = []
    texto_fodder = "🧱 *RADAR DE MEDIAS SBC (Base Exclusiva FUT.GG)*\n━━━━━━━━━━━━━━━━━━\n"
    for media_obj, limites in sorted(REFERENCIA_FODDER.items()):
        patron = f"(Media {media_obj}"
        candidatos = [j for j in datos_jugadores if patron in str(j.get("Jugador", "")) and int(j.get("Precio_Actual", 0) or 0) > 0]

        if candidatos:
            mas_barato = min(candidatos, key=lambda x: int(x["Precio_Actual"]))
            precio_min = int(mas_barato["Precio_Actual"])
            enlace_f = mas_barato.get("Enlace_Optimo", mas_barato.get("Enlace_Futbin", ""))
            nombre_limpio = mas_barato["Jugador"].split(" (")[0]

            if precio_min <= limites["compra"]:
                estado = "✅ COMPRAR"
            elif precio_min >= limites["venta"]:
                estado = "🚨 VENDER"
            else:
                estado = "⏳ ESPERAR"

            texto_fodder += (
                f"🔹 *Media {media_obj}:* `{precio_min:,} mnds` ({estado})\n"
                f"   └ 🎯 Piso: `{limites['compra']:,}` | Techo: `{limites['venta']:,}`\n"
                f"   └ 🛒 Más barato: [*{nombre_limpio}*]({enlace_f})\n\n"
            )
            filas_fodder.append({
                "Media": media_obj,
                "Precio_Mas_Barato": precio_min,
                "Jugador_Referencia": nombre_limpio,
                "Piso_Compra": limites["compra"],
                "Techo_Venta": limites["venta"],
                "Senal": estado
            })
        else:
            texto_fodder += f"🔹 *Media {media_obj}:* Sin muestra en este escaneo.\n\n"
    return texto_fodder, pd.DataFrame(filas_fodder)

def es_carta_base_url_gg(enlace):
    """En FUT.GG una carta base (Oro/Plata/Bronce) tiene el mismo ID en el slug y en '/27-{id}/'; los TOTW siempre tienen ID versionado distinto."""
    m = re.search(r'/players/(\d+)-[^/]+/27-(\d+)(?:/|$)', str(enlace or ""))
    if m and m.group(1) == m.group(2):
        return True
    return False

def es_etiqueta_totw_estricta(nombre_full, enlace=""):
    """Exige etiqueta explícita TOTW/Team of the Week/In-Form/- IF) y descarta explícitamente Normal/Gold/Rare/Common/Hero/Icon/POTM/TOTY/TOTS/SBC/Objective."""
    if not nombre_full:
        return False
    if es_carta_base_url_gg(enlace):
        return False
    nom_up = sin_tildes(str(nombre_full)).upper()
    terminos_excluidos = ("NORMAL", "GOLD", "RARE", "COMMON", "HERO", "ICON", "POTM", "TOTY", "TOTS", "SBC", "OBJECTIVE")
    if re.search(r'\b(?:' + "|".join(terminos_excluidos) + r')\b', nom_up):
        return False
    tiene_etiqueta_totw = (
        ("TOTW" in nom_up)
        or ("TEAM OF THE WEEK" in nom_up)
        or ("IN-FORM" in nom_up)
        or ("- IF)" in nom_up)
        or bool(re.search(r'(?:-\s*IF\b|\(IF\s+\d{2}\))', nom_up))
    )
    return tiene_etiqueta_totw

def analizar_totw_piso(datos_jugadores):
    if not datos_jugadores:
        return "", pd.DataFrame(columns=COLUMNAS_RADAR_TOTW)
    candidatos_totw = []
    for j in datos_jugadores:
        nombre_full = str(j.get("Jugador", ""))
        precio = int(j.get("Precio_Actual", 0) or 0)
        enlace_t = j.get("Enlace_Optimo", j.get("Enlace_Futbin", ""))
        media_str = extraer_media_exacta(nombre_full)
        media = int(media_str) if media_str else 0
        es_totw = es_etiqueta_totw_estricta(nombre_full, enlace=enlace_t)
        if es_totw and 9500 <= precio <= TOPE_TOTW_DESCARTE and media >= 83:
            candidatos_totw.append({
                "Jugador": nombre_full.split(" (")[0] + f" (Media {media} - TOTW)",
                "Media": media,
                "Precio_Actual": precio,
                "Fuente": "FUT.GG",
                "Break_Even_5%": calcular_break_even(precio),
                "Precio_Lazy_1h": calcular_precio_lazy(
                    precio,
                    max_hist=int(j.get("Precio_Max_Hist", 0) or 0),
                    muestras=int(j.get("Muestras_Hist", 0) or 0)
                ),
                "Riesgo": "🟢 Nulo (Cerca de descarte)" if precio <= 12500 else "🟡 Bajo (Piso SBC)",
                "Enlace_Futbin": enlace_t
            })
    df_totw = pd.DataFrame(candidatos_totw, columns=COLUMNAS_RADAR_TOTW)
    if df_totw.empty:
        return "🖤 *RADAR TOTW EN PISO (FUT.GG)*\n━━━━━━━━━━━━━━━━━━\nSin cartas TOTW de media 83+ en piso (9,500–16,500 mnds) en este corte.\n\n", df_totw
    df_totw = df_totw.sort_values(by=["Precio_Actual", "Media"], ascending=[True, False]).reset_index(drop=True)
    txt_totw = "🖤 *RADAR TOTW EN PISO (FUT.GG)*\n━━━━━━━━━━━━━━━━━━\n"
    for _, row in df_totw.head(5).iterrows():
        txt_totw += (
            f"👤 [*{row['Jugador']}*]({row['Enlace_Futbin']}) — `{row['Precio_Actual']:,} mnds`\n"
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

        jug_norm = sin_tildes(jug_comprado).lower()
        cands_c = [j for j in (datos_finales or []) if jug_norm in sin_tildes(j.get('Jugador', '')).lower()]
        match = min(cands_c, key=lambda j: abs(int(j.get('Precio_Actual', 0)) - precio_pagado)) if cands_c else None
        if match:
            precio_hoy = precio_mercado_seguro_carta(match, fallback_precio=precio_pagado)
            lazy_hoy = calcular_precio_lazy(
                precio_hoy,
                precio_pagado,
                max_hist=int(match.get('Precio_Max_Hist', 0) or 0),
                muestras=int(match.get('Muestras_Hist', 0) or 0)
            )
            neto = int(round((precio_hoy * 0.95) - precio_pagado))
            rentabilidad = (neto / precio_pagado) * 100 if precio_pagado > 0 else 0.0
            roi_diario = round(rentabilidad / max(1.0, dias_hold), 1)
            nombre_corto = match['Jugador'].split(' (')[0]
            enlace_c = match.get('Enlace_Optimo', match.get('Enlace_Futbin', ''))
            cotiz_str = f"FUT.GG: `{precio_hoy:,.0f} mnds`"

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

def analizar_jugador_individual(nombre_buscado, df_viejo):
    busq = sin_tildes(nombre_buscado).lower().strip()
    if df_viejo is None or df_viejo.empty:
        return "⚠️ Todavía no hay historial guardado para analizar jugadores."

    mask = df_viejo['Jugador'].apply(lambda x: busq in sin_tildes(str(x)).lower())
    df_jug = df_viejo[mask]

    if df_jug.empty:
        return f"⚠️ No encontré a `{nombre_buscado}` en el padrón FUT.GG."

    ult_registro = df_jug.iloc[-1]
    col_enl = 'Enlace_Optimo' if 'Enlace_Optimo' in df_viejo.columns else 'Enlace_Futbin'
    enlace_principal = ult_registro.get('Enlace_Optimo', ult_registro.get('Enlace_Futbin', ''))
    df_carta = df_viejo[df_viejo[col_enl] == enlace_principal]

    nombre_oficial = ult_registro['Jugador']
    precio_hoy = int(float(ult_registro['Precio_Actual']))
    senal_hoy = str(ult_registro.get('Senal_Trading', '⚪ Rango Neutro'))

    precios_num = pd.to_numeric(df_carta['Precio_Actual'], errors='coerce').fillna(precio_hoy)
    p_min = int(precios_num.min())
    p_max = int(precios_num.max())
    p_prom = int(round(precios_num.mean()))
    muestras = len(df_carta)

    puja_10 = calcular_puja_maxima_segura(precio_hoy, 0.10)
    break_even = calcular_break_even(precio_hoy)
    lazy_sug = calcular_precio_lazy(precio_hoy, max_hist=p_max, muestras=muestras)
    stop_loss = calcular_stop_loss_sugerido(precio_hoy)

    prompt_lupa = f"""
    Actuá como Trader experto de EA Sports FC 27 en PC (Base Exclusiva FUT.GG).
    Analizá esta carta con datos reales de nuestra base:
    - Carta: {nombre_oficial} | Diagnóstico cuantitativo: {senal_hoy}
    - Precio Actual FUT.GG: {precio_hoy:,} monedas
    - Mínimo Histórico: {p_min:,} | Promedio: {p_prom:,} | Máximo: {p_max:,} ({muestras} escaneos)
    - Puja máxima sugerida (+10% ROI): {puja_10:,} | Venta Lazy Buyer (+12%): {lazy_sug:,} | Stop-Loss (-10%): {stop_loss:,}

    Dame un veredicto táctico en 4 líneas: ¿Es apta para Flipping Rápido / Lazy Listing hoy o hay riesgo de caída (cuchillo cayendo)?
    """
    veredicto_ia, _ = llamar_gemini_turbo(prompt_lupa)
    if not veredicto_ia:
        veredicto_ia = "Carta en zona operativa. Buscar entrada únicamente mediante puja por debajo de la Puja Máxima Sugerida y listar a precio Lazy."

    reporte_lupa = (
        f"🔍 *LUPA DE TRADING FUT.GG: {nombre_oficial}*\n━━━━━━━━━━━━━━━━━━\n"
        f"🏷️ *Diagnóstico:* `{senal_hoy}`\n"
        f"💰 *Precio Actual (FUT.GG):* `{precio_hoy:,} mnds`\n\n"
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
        f"🔗 [Abrir carta en FUT.GG]({enlace_principal})"
    )
    return reporte_lupa

# ==========================================
# NAVEGADOR ULTRA-RÁPIDO Y ESCÁNER MASIVO FUT.GG (10,000 CARTAS)
# ==========================================
def crear_navegador_optimizado(ocultar=False):
    options = ChromiumOptions()
    options.set_argument('--window-size=1920,1080')
    options.set_argument('--disable-extensions')
    options.set_argument('--disable-blink-features=AutomationControlled')
    options.set_argument('--blink-settings=imagesEnabled=false')
    options.no_imgs(True)
    try:
        options.set_load_mode('eager')
    except Exception:
        pass

    en_nube = os.environ.get("GITHUB_ACTIONS") == "true" or sys.platform != "win32"
    if en_nube:
        options.set_argument('--headless=new')
        options.set_argument('--no-sandbox')
        options.set_argument('--disable-dev-shm-usage')
        options.set_argument('--disable-gpu')
        options.set_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36')
    elif ocultar:
        options.set_argument('--window-position=-32000,-32000')

    page = ChromiumPage(addr_or_opts=options)
    page.set.timeouts(page_load=10)
    try:
        page.run_cdp('Page.addScriptToEvaluateOnNewDocument', source="""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        """)
    except Exception:
        pass
    try:
        page.run_cdp('Network.enable')
        page.run_cdp('Network.setBlockedURLs', urls=[
            "*.png", "*.jpg", "*.jpeg", "*.webp", "*.gif", "*.svg", "*.ico",
            "*.woff", "*.woff2", "*.ttf", "*.otf", "*.mp4", "*.webm",
            "*doubleclick.net*", "*googlesyndication.com*", "*googletagmanager.com*",
            "*google-analytics.com*", "*facebook.net*", "*adservice*", "*amazon-adsystem*"
        ])
    except Exception:
        pass
    return page

def _detener_carga_segura(page):
    try:
        page.stop_loading()
    except Exception:
        pass

def detectar_tipo_carta_gg(texto_card, clases_card="", texto_alts="", texto_badges="", href=""):
    """
    Detecta la rareza real de una carta en FUT.GG revisando img[alt], clases/badges de rareza
    y texto del contenedor individual (`box`). Nunca etiqueta una carta base (base_id == card_id)
    ni cartas Gold/Normal/Hero/Icon/Promo como TOTW.
    """
    if es_carta_base_url_gg(href):
        return "Normal"

    combinado = sin_tildes(f"{texto_alts} {texto_badges} {clases_card} {texto_card}").upper()
    exclusiones_no_totw = ("HERO", "ICON", "POTM", "TOTY", "TOTS", "SBC", "OBJECTIVE", "EVOLUTION")
    tiene_exclusion_promo = bool(re.search(r'\b(?:' + "|".join(exclusiones_no_totw) + r')\b', combinado))
    tiene_marca_base = bool(re.search(r'\b(?:RARE GOLD|COMMON GOLD|GOLD RARE|GOLD COMMON|COMMON SILVER|RARE SILVER|COMMON BRONZE|RARE BRONZE)\b', combinado))

    es_totw_real = (
        not tiene_exclusion_promo
        and not tiene_marca_base
        and any(k in combinado for k in ["TEAM OF THE WEEK", "TOTW", "IN-FORM", "RARITY-TOTW", "RARITY_TOTW", "RARITY-3", "RARITY_3"])
    )
    if es_totw_real:
        return "TOTW"
    if tiene_exclusion_promo or "EVO " in combinado or categoria_version_carta(combinado) == "Promo/Especial":
        return "Promo/Especial"
    return "Normal"

def extraer_media_futgg(texto_inner, texto_alts, valor_gg=0, min_media=65):
    """
    Detecta la valoración general (OVR) en FUT.GG leyendo PRIMERO la primera línea numérica
    independiente de 2 dígitos (min_media..99) dentro de box.innerText antes de revisar img[alt],
    evitando confundir estadísticas individuales (PAC/SHO/PAS/DRI/DEF/PHY) con la media.
    """
    # 1. Primera línea numérica independiente de 2 dígitos en box.innerText
    for linea in str(texto_inner or "").splitlines():
        l_clean = linea.strip()
        if re.fullmatch(r'[6-9]\d', l_clean):
            val_m = int(l_clean)
            if min_media <= val_m <= 99 and val_m != valor_gg:
                return str(val_m)

    # 2. Fallback controlado a img[alt]
    for cand_m in re.findall(r'\b([6-9]\d)\b', str(texto_alts or "")):
        val_m = int(cand_m)
        if min_media <= val_m <= 99 and val_m != valor_gg:
            return str(val_m)

    # 3. Último recurso: tokens limpios en texto_inner excluyendo etiquetas de stats
    inner_sin_stats = _limpiar_ruido_dom_precio(texto_inner)
    for cand_m in re.findall(r'\b([6-9]\d)\b', inner_sin_stats):
        val_m = int(cand_m)
        if min_media <= val_m <= 99 and val_m != valor_gg:
            return str(val_m)
    return "??"

def escanear_futgg(page, presupuesto_max, max_paginas=PAGINAS_FUTGG):
    print(f"\n⚡ Iniciando escaneo masivo de hasta {MAX_JUGADORES} cartas en FUT.GG PC ({max_paginas} páginas <= {presupuesto_max:,} mnds)...")
    jugadores_gg = []
    vistos_gg = set()

    try:
        page.get("https://www.fut.gg/players/")
        page.wait(0.8)
        page.run_js("""
        try {
            localStorage.setItem('platform', 'pc');
            document.cookie = "platform=pc; path=/; max-age=31536000";
        } catch(e) {}
        """)
    except Exception:
        pass

    # Regla estricta: aislar únicamente la caja individual de cada jugador (contarLinksJugador === 1)
    script_gg = """
    window.scrollTo(0, document.body.scrollHeight / 2);
    function contarLinksJugador(el) {
        if (!el) return 99;
        let ls = el.querySelectorAll('a[href*="/players/"]');
        let s = new Set();
        for (let l of ls) {
            let h = (l.getAttribute('href') || '').split('?')[0];
            if (h.match(/\\/players\\/\\d+-/)) s.add(h);
        }
        return s.size;
    }
    let items = [];
    let enlaces = document.querySelectorAll('a[href*="/players/"]');
    for (let a of enlaces) {
        let href = a.getAttribute('href') || '';
        if (!href.match(/\\/players\\/\\d+-/)) continue;
        let box = a;
        while (box.parentElement && contarLinksJugador(box.parentElement) === 1) {
            box = box.parentElement;
        }
        let alts = [];
        for (let img of box.querySelectorAll('img')) {
            let alt_txt = img.getAttribute('alt') || '';
            let title_txt = img.getAttribute('title') || '';
            let src_txt = img.getAttribute('src') || '';
            if (alt_txt) alts.push(alt_txt);
            if (title_txt) alts.push(title_txt);
            let src_low = src_txt.toLowerCase();
            if (src_low.includes('team-of-the-week') || src_low.includes('totw') || src_low.includes('in-form') || src_low.includes('/rarity/3') || src_low.includes('rarity-3')) alts.push('Team of the Week TOTW');
            else if (src_low.includes('icon')) alts.push('ICON');
            else if (src_low.includes('hero')) alts.push('HERO');
            else if (src_low.includes('toty')) alts.push('TOTY');
            else if (src_low.includes('tots')) alts.push('TOTS');
            else if (src_low.includes('potm')) alts.push('POTM');
            else if (src_low.includes('promo') || src_low.includes('special') || src_low.includes('rttk')) alts.push('PROMO');
        }
        let badges = [];
        for (let b of box.querySelectorAll('[class*="rarity"], [class*="badge"], [class*="version"], [data-rarity], [data-card-type]')) {
            let bt = (b.innerText || b.textContent || '').trim();
            let bc = b.className || '';
            let br = b.getAttribute('data-rarity') || '';
            let ba = b.getAttribute('aria-label') || b.getAttribute('title') || '';
            if (bt) badges.push(bt);
            if (typeof bc === 'string' && bc) badges.push(bc);
            if (br) badges.push('rarity-' + br);
            if (ba) badges.push(ba);
        }
        let nodo_p = box.querySelector('[class*="price"], [data-testid*="price"]');
        let texto_precio = nodo_p ? (nodo_p.innerText || nodo_p.textContent || '') : '';
        let clases_box = (box.className || '') + ' ' + (a.className || '') + ' ' + (box.getAttribute('data-rarity') || '');
        let texto_inner = box.innerText || box.textContent || '';
        let texto_alts = alts.join(' | ');
        let texto_badges = badges.join(' | ');
        let texto_card = texto_alts + '\\n' + texto_inner;
        items.push({
            'href': href,
            'texto': texto_card,
            'texto_inner': texto_inner,
            'texto_alts': texto_alts,
            'texto_badges': texto_badges,
            'texto_precio': texto_precio,
            'clases': clases_box
        });
    }
    return items;
    """

    # Pase 0: Escaneo relámpago dedicado de cartas TOTW en piso (<= TOPE_TOTW_DESCARTE), validando rareza real en DOM
    urls_totw_dedicadas = [
        f"https://www.fut.gg/players/?page=1&rarities=3&price__gte=9000&price__lte={TOPE_TOTW_DESCARTE}",
        f"https://www.fut.gg/players/?page=2&rarities=3&price__gte=9000&price__lte={TOPE_TOTW_DESCARTE}"
    ]
    for url_t in urls_totw_dedicadas:
        try:
            page.get(url_t)
            if page.ele('xpath://a[contains(@href, "/players/")]', timeout=4):
                page.wait(0.4)
                _detener_carga_segura(page)
                datos_t = page.run_js(script_gg) or []
                for item in datos_t:
                    href_limpio = (item.get('href', '') or '').split('?')[0]
                    if not href_limpio or href_limpio in vistos_gg:
                        continue
                    match_slug = re.search(r'/players/\d+-([a-zA-Z0-9\-]+)', href_limpio)
                    if not match_slug:
                        continue
                    nombre_gg = match_slug.group(1).replace('-', ' ').title()
                    texto = item.get('texto', '')
                    texto_inner = item.get('texto_inner', texto)
                    texto_alts = item.get('texto_alts', '')
                    texto_badges = item.get('texto_badges', '')
                    texto_p = item.get('texto_precio', '')
                    clases_box = item.get('clases', '')
                    valor_gg = extraer_precio_carta_unica_gg(texto, TOPE_TOTW_DESCARTE, texto_precio=texto_p)
                    if not (9000 <= valor_gg <= TOPE_TOTW_DESCARTE) or not es_multiplo_escalon_ea(valor_gg, TOPE_TOTW_DESCARTE):
                        continue
                    # Verificar rareza REAL dentro de la caja individual (jamás asumir TOTW por la URL)
                    tipo_real_t = detectar_tipo_carta_gg(
                        texto_inner, clases_card=clases_box,
                        texto_alts=texto_alts, texto_badges=texto_badges,
                        href=href_limpio
                    )
                    if tipo_real_t != "TOTW":
                        continue
                    media = extraer_media_futgg(texto_inner, texto_alts, valor_gg=valor_gg, min_media=75)
                    if media == "??":
                        continue
                    vistos_gg.add(href_limpio)
                    nombre_fmt = f"{nombre_gg} (Media {media} - TOTW - FUT.GG)"
                    jugadores_gg.append({
                        "Jugador_GG": nombre_fmt,
                        "Nombre_Puro": sin_tildes(nombre_gg).lower(),
                        "Media": media,
                        "Tipo_Carta": "TOTW",
                        "Precio_FUTGG": valor_gg,
                        "Clave": clave_cruce(nombre_fmt),
                        "Enlace_GG": f"https://www.fut.gg{href_limpio}" if href_limpio.startswith('/') else href_limpio
                    })
        except Exception:
            continue

    paginas_vacias_seguidas = 0
    for pag in range(1, max_paginas + 1):
        if len(jugadores_gg) >= MAX_JUGADORES:
            break
        url_gg = f"https://www.fut.gg/players/?page={pag}&price__gte=200&price__lte={presupuesto_max}"
        print(f"🌐 FUT.GG PC Página {pag}/{max_paginas}... (Capturados GG: {len(jugadores_gg):,}/{MAX_JUGADORES:,})")
        datos_gg = None
        for intento in range(2):
            try:
                page.get(url_gg)
                if page.ele('xpath://a[contains(@href, "/players/")]', timeout=5):
                    page.wait(random.uniform(0.25, 0.45))
                    _detener_carga_segura(page)
                    datos_gg = page.run_js(script_gg)
                    break
            except Exception:
                if intento == 0:
                    time.sleep(0.4)
                    continue
        if not datos_gg:
            paginas_vacias_seguidas += 1
            if paginas_vacias_seguidas >= 5:
                break
            continue

        nuevos_en_pagina = 0
        for item in (datos_gg or []):
            if len(jugadores_gg) >= MAX_JUGADORES:
                break
            href = item.get('href', '')
            href_limpio = href.split('?')[0]
            if not href_limpio or href_limpio in vistos_gg:
                continue
            match_slug = re.search(r'/players/\d+-([a-zA-Z0-9\-]+)', href_limpio)
            if not match_slug:
                continue
            nombre_gg = match_slug.group(1).replace('-', ' ').title()
            texto = item.get('texto', '')
            texto_inner = item.get('texto_inner', texto)
            texto_alts = item.get('texto_alts', '')
            texto_badges = item.get('texto_badges', '')
            texto_p = item.get('texto_precio', '')
            clases_box = item.get('clases', '')

            valor_gg = extraer_precio_carta_unica_gg(texto, presupuesto_max, texto_precio=texto_p)
            if not es_multiplo_escalon_ea(valor_gg, presupuesto_max):
                continue

            media = extraer_media_futgg(texto_inner, texto_alts, valor_gg=valor_gg, min_media=65)
            if media == "??":
                continue

            tipo_carta_gg = detectar_tipo_carta_gg(
                texto_inner, clases_card=clases_box,
                texto_alts=texto_alts, texto_badges=texto_badges,
                href=href_limpio
            )
            vistos_gg.add(href_limpio)
            nuevos_en_pagina += 1
            nombre_fmt = f"{nombre_gg} (Media {media} - {tipo_carta_gg} - FUT.GG)"
            jugadores_gg.append({
                "Jugador_GG": nombre_fmt,
                "Nombre_Puro": sin_tildes(nombre_gg).lower(),
                "Media": media,
                "Tipo_Carta": tipo_carta_gg,
                "Precio_FUTGG": valor_gg,
                "Clave": clave_cruce(nombre_fmt),
                "Enlace_GG": f"https://www.fut.gg{href_limpio}" if href_limpio.startswith('/') else href_limpio
            })

        if nuevos_en_pagina == 0:
            paginas_vacias_seguidas += 1
            if paginas_vacias_seguidas >= 5:
                break
        else:
            paginas_vacias_seguidas = 0

    print(f"✅ ¡FUT.GG PC completado! {len(jugadores_gg):,} cartas reales extraídas.")
    return jugadores_gg

def detectar_evoluciones(datos_procesados):
    if not datos_procesados:
        return "", pd.DataFrame()

    evos_detectadas = []
    vistos_evo = set()
    for f in datos_procesados:
        enlace = f.get("Enlace_Optimo", f.get("Enlace_Futbin", ""))
        if not enlace or enlace in vistos_evo:
            continue
        nombre_corto = str(f.get("Jugador", "")).split(" (")[0]
        p_ant = float(f.get("Precio_Anterior", 0))
        p_act = float(f.get("Precio_Actual", 0))
        var_b = float(f.get("Variacion_Bruta_%", 0))

        if 200 <= p_ant <= TOPE_PRECIO_BASE_EVO and var_b >= UMBRAL_SUBIDA_EVO_PCT and (p_act - p_ant) >= 700:
            vistos_evo.add(enlace)
            evos_detectadas.append({
                "Jugador": f.get("Jugador", nombre_corto),
                "Origen_Alerta": "🔥 Explosión FUT.GG",
                "Precio_Base": int(p_ant),
                "Precio_Explosivo": int(p_act),
                "Subida_%": round(var_b, 1),
                "Ganancia_Potencial_Neta": int(round((p_act * 0.95) - p_ant)),
                "Enlace": enlace
            })

    df_evos = pd.DataFrame(evos_detectadas)
    if df_evos.empty:
        return "🧬 *EXPLOSIÓN POR EVOLUCIONES (FUT.GG)*\n━━━━━━━━━━━━━━━━━━\nSin picos anormales (>35%).\n\n", df_evos
    df_evos = df_evos.sort_values(by="Subida_%", ascending=False)
    txt_evo = "🧬 *EXPLOSIÓN POR EVOLUCIONES (FUT.GG)*\n━━━━━━━━━━━━━━━━━━\n"
    for _, row in df_evos.head(5).iterrows():
        txt_evo += f"👤 [*{str(row['Jugador']).split(' (')[0]}*]({row['Enlace']}) ({row['Origen_Alerta']})\n🚀 Precio: `{row['Precio_Base']:,}` ➡️ `{row['Precio_Explosivo']:,} mnds`\n📈 Subida brutal: `+{row['Subida_%']}%`\n\n"
    return txt_evo, df_evos

def procesar_mensajes_telegram(df_cartera, df_viejo, df_evos_viejo, df_alertas):
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
                    respuesta_lupa = analizar_jugador_individual(jug_lupa, df_viejo)
                    enviar_alerta_telegram(respuesta_lupa)

                elif t_low in ["/flipping", "⚡ flipping / gangas", "⚡ fut.gg"]:
                    enviar_alerta_telegram(armar_resumen_flipping_gangas(datos_recientes))

                elif t_low in ["/links", "/domino", "🔗 links dominó"]:
                    txt_dom, _ = detectar_links_domino(datos_recientes, df_evos_viejo)
                    enviar_alerta_telegram(txt_dom)

                elif t_low in ["/pnl", "/contabilidad", "📊 estado p&l", "/balance", "💼 balance"]:
                    rep_pnl, _, _ = generar_estado_resultados_contable(df_cartera, datos_recientes)
                    enviar_alerta_telegram(rep_pnl)

                elif t_low in ["/oraculo", "/leaks", "🔮 oráculo leaks"]:
                    enviar_alerta_telegram("🔮 _Consultando al Oráculo de Contenido sobre la Base Exclusiva FUT.GG..._")
                    rep_oraculo = escanear_oraculo_contenido(None, datos_recientes, ocultar=True)
                    enviar_alerta_telegram(rep_oraculo)

                elif t_low in ["/ayuda", "🔍 ayuda comandos"]:
                    ayuda_txt = (
                        "🛠️ *CENTRO DE COMANDOS FC27 (BASE EXCLUSIVA FUT.GG 10,000 CARTAS)*\n━━━━━━━━━━━━━━━━━━\n"
                        "• `Compra: Jugador, Precio` _(guarda costo, Lazy y Stop-Loss)_\n"
                        "• `Venta: Jugador, Precio` ➡️ Cierra trade y calcula Costeo Variable\n"
                        "• `⚡ Flipping / Gangas` ➡️ Filtra rebotes rápidos vs Cuchillos cayendo\n"
                        "• `🔗 Links Dominó` ➡️ Cartas dormidas de 2da ola por Evos/SBC\n"
                        "• `📊 Estado P&L` ➡️ Estado de Resultados, VNR y Rotación de Stock\n"
                        "• `🔮 Oráculo Leaks` ➡️ Radar 15:00 hs ARG y Hype\n"
                        "• `/analizar Jugador` ➡️ Lupa FUT.GG + Puja + Lazy + Stop-Loss\n"
                        "• `/alerta Jugador, Precio` ➡️ Crea Sniper Alert en FUT.GG"
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
                            enviar_alerta_telegram(f"🎯 *ALERTA SNIPER CREADA (FUT.GG)*\nTe aviso apenas `{jug_b}` toque `{precio_obj:,} mnds` o menos.")
                    except Exception:
                        pass

                elif t_low.startswith("compra:"):
                    try:
                        datos = [x.strip() for x in texto.split(":", 1)[1].split(",")]
                        if len(datos) >= 2:
                            jugador = datos[0].title()
                            precio = int(re.sub(r'[^\d]', '', datos[1]))
                            est_elegida = datos[2].title() if (len(datos) >= 3 and datos[2]) else clasificar_estrategia_automatica(jugador, precio, datos_recientes)
                            empate = calcular_break_even(precio)
                            lazy_ini = calcular_precio_lazy(precio, precio)
                            sl_ini = calcular_stop_loss_sugerido(precio)

                            df_cartera = pd.concat([df_cartera, pd.DataFrame([{
                                "Jugador_Exacto": str(jugador),
                                "Estrategia": str(est_elegida),
                                "Cantidad": 1,
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
                                f"⚖️ *Empate (5%):* `{empate:,}` | 🎯 *Lazy:* `{lazy_ini:,}` | 🛑 *SL:* `{sl_ini:,}`"
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
                                cant_v = max(1, int(pd.to_numeric(df_cartera.loc[idx_v, 'Cantidad'], errors='coerce') or 1)) if 'Cantidad' in df_cartera.columns else 1
                                precio_compra = float(df_cartera.loc[idx_v, 'Precio_Pagado'])
                                jugador_real = df_cartera.loc[idx_v, 'Jugador_Exacto']
                                venta_bruta_lote = precio_venta * cant_v
                                costo_lote = precio_compra * cant_v
                                costo_tax = int(round(venta_bruta_lote * 0.05))
                                ganancia_limpia = int(round((venta_bruta_lote - costo_tax) - costo_lote))
                                roi = round((ganancia_limpia / costo_lote) * 100, 2) if costo_lote > 0 else 0.0

                                df_cartera["Estado"] = df_cartera["Estado"].astype(object)
                                df_cartera["Fecha_Venta"] = df_cartera["Fecha_Venta"].astype(object)
                                df_cartera.loc[idx_v, 'Estado'] = 'Vendido'
                                df_cartera.loc[idx_v, 'Precio_Venta'] = precio_venta
                                df_cartera.loc[idx_v, 'Costo_EA_Tax'] = costo_tax
                                df_cartera.loc[idx_v, 'Ganancia_Neta'] = ganancia_limpia
                                df_cartera.loc[idx_v, 'ROI_%'] = roi
                                df_cartera.loc[idx_v, 'Fecha_Venta'] = ahora_arg().strftime("%Y-%m-%d %H:%M")

                                hubo_cambios = True
                                enviar_alerta_telegram(f"📈 *VENTA CERRADA: {cant_v}x {jugador_real}*\n• Neto Limpio: `{ganancia_limpia:+,} mnds` (`{roi:+}%`)")
                    except Exception:
                        pass

                elif t_low in ["/cartera", "🎒 mi cartera"]:
                    enviar_alerta_telegram(armar_texto_cartera(df_cartera, datos_recientes))

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
        jug_buscado = sin_tildes(str(alerta['Jugador_Buscado'])).strip().lower()
        precio_obj = float(alerta['Precio_Objetivo'])
        cands = [j for j in datos_finales if jug_buscado in sin_tildes(j['Jugador']).lower() and j['Precio_Actual'] > 0]
        match = min(cands, key=lambda j: j['Precio_Actual']) if cands else None
        if match and match['Precio_Actual'] <= precio_obj:
            puja_sug = calcular_puja_maxima_segura(match['Precio_Actual'], 0.10)
            lazy_sug = calcular_precio_lazy(
                match['Precio_Actual'],
                max_hist=int(match.get("Precio_Max_Hist", 0) or 0),
                muestras=int(match.get("Muestras_Hist", 0) or 0)
            )
            enlace_s = match.get("Enlace_Optimo", match.get("Enlace_Futbin", ""))
            enviar_alerta_telegram(
                f"🚨 🎯 *¡ALERTA SNIPER CUMPLIDA (FUT.GG)!* 🎯 🚨\n━━━━━━━━━━━━━━━━━━\n"
                f"👤 [*{match['Jugador'].split(' (')[0]}*]({enlace_s}) cayó a `{match['Precio_Actual']:,} mnds`.\n"
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
    columnas_cartera = ["Jugador_Exacto", "Estrategia", "Cantidad", "Precio_Pagado", "Estado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%", "Fecha_Compra", "Fecha_Venta"]
    cols_texto_cartera = ["Jugador_Exacto", "Estrategia", "Estado", "Fecha_Compra", "Fecha_Venta"]
    cols_num_cartera = ["Cantidad", "Precio_Pagado", "Precio_Venta", "Costo_EA_Tax", "Ganancia_Neta", "ROI_%"]
    columnas_alertas = ["Jugador_Buscado", "Precio_Objetivo", "Estado"]

    df_viejo = None
    df_evos_viejo = pd.DataFrame()
    df_cartera = pd.DataFrame(columns=columnas_cartera)
    df_alertas = pd.DataFrame(columns=columnas_alertas)
    historial_precios = {}

    if os.path.exists(archivo_excel):
        try:
            df_v_raw = pd.read_excel(archivo_excel, sheet_name="Historial_Tendencias")
            if not df_v_raw.empty and "Jugador" in df_v_raw.columns and "Precio_Actual" in df_v_raw.columns:
                mask_bad = df_v_raw["Jugador"].astype(str).str.contains(r"Media \?\?", na=False) & (pd.to_numeric(df_v_raw["Precio_Actual"], errors="coerce") == 36000)
                df_v_raw = df_v_raw[~mask_bad].copy()
            df_viejo = df_v_raw
            if 'Fecha' in df_viejo.columns:
                df_viejo = df_viejo.sort_values(by='Fecha')
            col_key = 'Enlace_Optimo' if 'Enlace_Optimo' in df_viejo.columns else 'Enlace_Futbin'
            if col_key in df_viejo.columns:
                df_hist_pos = df_viejo[pd.to_numeric(df_viejo['Precio_Actual'], errors='coerce').fillna(0) > 0].drop_duplicates(subset=[col_key], keep='last')
                historial_precios = dict(zip(df_hist_pos[col_key], pd.to_numeric(df_hist_pos['Precio_Actual'], errors='coerce').fillna(0)))
        except Exception:
            pass
        try:
            df_c = pd.read_excel(archivo_excel, sheet_name="Mi_Cartera")
            for col in columnas_cartera:
                if col not in df_c.columns:
                    df_c[col] = 1 if col == "Cantidad" else (0 if col in cols_num_cartera else "")
            for col_t in cols_texto_cartera:
                df_c[col_t] = df_c[col_t].fillna("").astype(str).replace("nan", "").astype(object)
            for col_n in cols_num_cartera:
                df_c[col_n] = pd.to_numeric(df_c[col_n], errors="coerce").fillna(1 if col_n == "Cantidad" else 0)
            df_c["Cantidad"] = df_c["Cantidad"].clip(lower=1).astype(int)
            df_cartera = df_c[columnas_cartera]
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

    if (df_viejo is None or df_viejo.empty) and os.path.exists(ARCHIVO_CRED):
        gc = obtener_cliente_gsheets()
        if gc:
            try:
                sh = gc.open(NOMBRE_SHEET)
                df_v_gs = leer_hoja_gsheets_segura(sh, "Historial_Tendencias")
                col_key_gs = "Enlace_Optimo" if "Enlace_Optimo" in df_v_gs.columns else "Enlace_Futbin"
                if not df_v_gs.empty and col_key_gs in df_v_gs.columns:
                    df_v_gs["Precio_Actual"] = pd.to_numeric(df_v_gs["Precio_Actual"], errors="coerce").fillna(0)
                    if "Fecha" in df_v_gs.columns:
                        df_v_gs = df_v_gs.sort_values(by="Fecha")
                    df_viejo = df_v_gs
                    df_gs_pos = df_viejo[df_viejo["Precio_Actual"] > 0].drop_duplicates(subset=[col_key_gs], keep="last")
                    historial_precios = dict(zip(df_gs_pos[col_key_gs], df_gs_pos['Precio_Actual']))
                df_c_gs = leer_hoja_gsheets_segura(sh, "Mi_Cartera")
                if not df_c_gs.empty:
                    for col in columnas_cartera:
                        if col not in df_c_gs.columns:
                            df_c_gs[col] = 1 if col == "Cantidad" else (0 if col in cols_num_cartera else "")
                    for col_n in cols_num_cartera:
                        df_c_gs[col_n] = pd.to_numeric(df_c_gs[col_n], errors="coerce").fillna(1 if col_n == "Cantidad" else 0)
                    df_c_gs["Cantidad"] = df_c_gs["Cantidad"].clip(lower=1).astype(int)
                    df_cartera = df_c_gs[columnas_cartera]
            except Exception:
                pass

    return df_viejo, df_evos_viejo, df_cartera, df_alertas, historial_precios

def formatear_dashboard_financiero_excel(ruta_excel):
    try:
        wb = openpyxl.load_workbook(ruta_excel)
        fill_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        font_blanca = Font(color="FFFFFF", bold=True)
        for name in wb.sheetnames:
            ws = wb[name]
            for cell in ws[1]:
                cell.fill = fill_azul
                cell.font = font_blanca
                cell.alignment = Alignment(horizontal="center")
        wb.save(ruta_excel)
    except Exception:
        pass

def ejecutar_ciclo_radar(modo_elegido, ocultar_navegador=False):
    inicio_tiempo = time.time()
    df_viejo, df_evos_viejo, df_cartera, df_alertas, historial_precios = cargar_estado_excel()
    stats_historicas = calcular_estadisticas_historicas(df_viejo)
    df_cartera, df_alertas, _ = procesar_mensajes_telegram(df_cartera, df_viejo, df_evos_viejo, df_alertas)

    datos_finales = []
    fecha_actual = ahora_arg().strftime("%Y-%m-%d %H:%M")
    page_activa = None

    if modo_elegido in ["2", "3", "5"] and df_viejo is not None and not df_viejo.empty:
        ultima_fecha = df_viejo['Fecha'].max()
        datos_previos = df_viejo[df_viejo['Fecha'] == ultima_fecha].to_dict('records')
        fecha_actual = ultima_fecha
        datos_finales = datos_previos
    else:
        print(f"\n🌐 Abriendo navegador ultra-rápido para Base Exclusiva FUT.GG (Hasta {MAX_JUGADORES:,} cartas | Tope: {PRESUPUESTO_MAXIMO:,} mnds)...")
        page_activa = crear_navegador_optimizado(ocultar=ocultar_navegador)

        datos_gg_actuales = escanear_futgg(page_activa, PRESUPUESTO_MAXIMO, PAGINAS_FUTGG)
        datos_finales = procesar_datos_futgg(
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
    evos_txt, df_evos = detectar_evoluciones(datos_finales)
    domino_txt, df_domino = detectar_links_domino(datos_finales, df_evos)

    dashboard_tecnico = (
        f"🖥️ *DASHBOARD MASIVO FC27 (BASE EXCLUSIVA FUT.GG — HASTA {MAX_JUGADORES:,} CARTAS)*\n"
        f"🕒 `{fecha_actual}` | 📦 *Padrón:* `{len(datos_finales):,} cartas`\n\n"
        f"{termometro_txt}{ciclo_wl_txt}{alertas_cartera_txt}{domino_txt}{fodder_txt}{totw_txt}{evos_txt}"
    )
    enviar_alerta_telegram(dashboard_tecnico)
    enviar_alerta_telegram(oraculo_txt)

    if not datos_finales:
        print("⚠️ No se obtuvieron datos en este ciclo.")
        return

    df_nuevo = pd.DataFrame(datos_finales)
    df_compras_ia = df_nuevo[df_nuevo['Accion'] == '🟢 COMPRAR'].sort_values(by='Variacion_Bruta_%').head(18)
    df_ventas_ia = df_nuevo[df_nuevo['Accion'] == '🔴 VENDER'].sort_values(by='Variacion_Neta_%', ascending=False).head(18)
    df_para_ia = pd.concat([df_compras_ia, df_ventas_ia, df_nuevo.head(20)]).drop_duplicates(subset=['Enlace_Optimo'])
    cols_ia = [c for c in ['Jugador', 'Precio_Anterior', 'Precio_Actual', 'Senal_Trading', 'Puja_Max_10%_ROI', 'Break_Even_5%', 'Precio_Lazy_Sugerido', 'Variacion_Bruta_%', 'Accion'] if c in df_para_ia.columns]
    resumen_csv = df_para_ia[cols_ia].to_string(index=False)
    dia_hoy = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][ahora_arg().weekday()]

    prompt_estrategico = f"""
    Actuá como Trader Profesional de FC 27 operando con BASE EXCLUSIVA FUT.GG ({len(datos_finales):,} cartas).
    Hoy es {dia_hoy}. Tax: 5%.
    Datos:
    {resumen_csv}
    Redactá sin tablas '|':
    1. TOP 5 COMPRAS DE FLIPPING & REBOTE (FUT.GG).
    2. TOP 5 VENTAS / TOMA DE GANANCIAS.
    3. ESTRATEGIA DE ROTACIÓN Y CICLO FUT CHAMPIONS PARA HOY ({dia_hoy}).
    """

    reporte_estrategico_ia, _ = llamar_gemini_turbo(prompt_estrategico)
    texto_limpio = f"🤖 *REPORTE ESTRATÉGICO FUT.GG ({dia_hoy})*\n━━━━━━━━━━━━━━━━━━\n{reporte_estrategico_ia.replace('**', '*')}" if reporte_estrategico_ia else "🟢 Escaneo masivo FUT.GG completado."
    enviar_alerta_telegram(texto_limpio)

    df_clubes_cloud_prev = pd.DataFrame()
    if os.path.exists(archivo_excel):
        try:
            df_clubes_cloud_prev = pd.read_excel(archivo_excel, sheet_name="Clubes_Cloud")
        except Exception:
            df_clubes_cloud_prev = pd.DataFrame()

    if modo_elegido not in ["2", "5"] and df_viejo is not None and not df_viejo.empty:
        for col in df_nuevo.columns:
            if col not in df_viejo.columns:
                df_viejo[col] = ""
        df_combinado = pd.concat([df_viejo[df_nuevo.columns], df_nuevo], ignore_index=True)

        # Compactación inteligente: 48h en alta resolución + 1 cierre diario por carta para días previos
        try:
            corte_48h_str = (ahora_arg() - timedelta(hours=48)).strftime("%Y-%m-%d %H:%M")
            fechas_str = df_combinado["Fecha"].astype(str)
            df_48h = df_combinado[fechas_str >= corte_48h_str].copy()
            df_antiguos = df_combinado[fechas_str < corte_48h_str].copy()
            if not df_antiguos.empty:
                df_antiguos["_Dia"] = df_antiguos["Fecha"].astype(str).str.slice(0, 10)
                col_dedup = "Enlace_Optimo" if "Enlace_Optimo" in df_antiguos.columns else "Enlace_Futbin"
                df_cierres_diarios = df_antiguos.drop_duplicates(subset=["_Dia", col_dedup], keep="last").drop(columns=["_Dia"])
                df_compacto = pd.concat([df_cierres_diarios, df_48h], ignore_index=True)
            else:
                df_compacto = df_48h
        except Exception:
            df_compacto = df_combinado

        df_final = df_compacto.tail(25000)
    else:
        df_final = df_nuevo

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
            if not df_evos.empty:
                df_evos.to_excel(writer, index=False, sheet_name="Radar_Evos")
            df_totw_excel = df_totw if (df_totw is not None and not df_totw.empty) else pd.DataFrame(columns=COLUMNAS_RADAR_TOTW)
            df_totw_excel.to_excel(writer, index=False, sheet_name="Radar_TOTW")
            if not df_domino.empty:
                df_domino.to_excel(writer, index=False, sheet_name="Radar_Links_Domino")
            if not df_clubes_cloud_prev.empty:
                df_clubes_cloud_prev.to_excel(writer, index=False, sheet_name="Clubes_Cloud")
        formatear_dashboard_financiero_excel(archivo_excel)
    except Exception as e:
        print(f"⚠️ Aviso guardando Excel local: {e}")

    sincronizar_todo_a_google_sheets(
        df_final, fecha_actual, texto_limpio, oraculo_txt,
        df_cartera, df_alertas, df_pnl_excel, df_fodder,
        df_evos, df_totw, df_domino
    )
    print(f"🏁 ¡Proceso Masivo FUT.GG ({len(datos_finales):,} cartas) completado en {round(time.time() - inicio_tiempo, 1)} segundos!")

# ==========================================
# EJECUCIÓN PRINCIPAL (COMPATIBLE CON GITHUB ACTIONS Y PC)
# ==========================================
if __name__ == "__main__":
    en_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"
    modo_auto = en_github_actions or (len(sys.argv) > 1 and sys.argv[1].lower() == "auto")
    modo_vigilante_auto = len(sys.argv) > 1 and sys.argv[1].lower() == "vigilante"

    if modo_vigilante_auto:
        modo_elegido = "4"
    elif modo_auto:
        modo_elegido = "1"
    else:
        modo_elegido = input("\n👉 Elegí una opción (1, 2, 3, 4 o 5 y Enter). O dejalo en blanco para Vigilante (4): ").strip() or "4"

    if modo_elegido == "4":
        while True:
            ejecutar_ciclo_radar("1", ocultar_navegador=True)
            time.sleep(INTERVALO_VIGILANTE_MIN * 60)
    else:
        ejecutar_ciclo_radar(modo_elegido, ocultar_navegador=modo_auto)
