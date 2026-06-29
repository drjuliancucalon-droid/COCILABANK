"""
parsers/banco_pdf.py — Parser de extracto bancario en PDF
Soporta Bancolombia, Davivienda, BBVA, Banco de Bogotá, Nequi, Colpatria,
Occidente, Popular y cualquier banco colombiano.
Formatos de fecha: d/MM (A), dd-MM-yyyy (B), dd/MM/yyyy (C).
Con fallback OCR.
"""
import re
import logging
from datetime import datetime

import pdfplumber
import pandas as pd

from utils.pdf_diagnostico import ocr_pdf_page, OCR_AVAILABLE
from storage.db import registrar_formato_pdf, buscar_formato_pdf

log = logging.getLogger(__name__)


def limpiar_num(t):
    t = str(t or '').strip()
    if not t:
        return None
    neg = t.startswith('-') or (t.startswith('(') and t.endswith(')'))
    t = re.sub(r'[\$\(\)\s]', '', t).replace(',', '')
    try:
        v = float(t)
        return -abs(v) if neg else v
    except Exception:
        return None


# --- Patrones de fecha reconocidos ---
_PAT_FECHA_A = re.compile(r'^\d{1,2}/\d{2}$')       # d/MM  (Bancolombia formato A: 2/01, 15/01)
_PAT_FECHA_B = re.compile(r'^\d{2}-\d{2}-\d{4}$')   # dd-MM-yyyy (formato B: 05-01-2026)
_PAT_FECHA_C = re.compile(r'^\d{2}/\d{2}/\d{4}$')   # dd/MM/yyyy (alternativo: 05/01/2026)


def es_fecha_banco(t):
    s = str(t or '').strip()
    return bool(_PAT_FECHA_A.match(s) or _PAT_FECHA_B.match(s) or _PAT_FECHA_C.match(s))


def _parse_fecha(fecha_raw, anio_extracto):
    """Convierte fecha_raw a datetime segun el formato detectado."""
    s = str(fecha_raw).strip()
    if _PAT_FECHA_A.match(s):
        # Formato d/MM: requiere el anio del extracto
        return pd.to_datetime(f'{anio_extracto}/{s}', format='%Y/%d/%m', errors='coerce')
    if _PAT_FECHA_B.match(s):
        # Formato dd-MM-yyyy: el anio esta en la propia fecha
        return pd.to_datetime(s, format='%d-%m-%Y', errors='coerce')
    if _PAT_FECHA_C.match(s):
        # Formato dd/MM/yyyy: el anio esta en la propia fecha
        return pd.to_datetime(s, format='%d/%m/%Y', errors='coerce')
    return pd.NaT


def _detectar_anio_extracto(texto_pag0):
    """
    Extrae el anio del periodo desde la primera pagina del extracto.

    PRIORIDAD:
      1. Fecha HASTA  (fin del periodo — el extracto pertenece a este anio/mes)
      2. Fecha DESDE con logica: si mes_desde == 12, el extracto es de Enero anio+1
      3. Primer anio encontrado en el texto

    Bancolombia cabecera tipica:
        PERIODO: DESDE: 2025/12/31  HASTA: 2026/01/31
    Queremos 2026 (HASTA), no 2025 (DESDE).
    """
    # 1. Buscar HASTA explicito
    m = re.search(r'HASTA\s*:?\s*(20\d{2})[/\-](\d{2})[/\-]\d{2}', texto_pag0, re.IGNORECASE)
    if m:
        log.debug("[banco_pdf] Anio desde HASTA: %s", m.group(1))
        return int(m.group(1))

    # 2. Buscar DESDE (fallback con logica de diciembre)
    m = re.search(r'DESDE\s*:?\s*(20\d{2})[/\-](\d{2})[/\-]\d{2}', texto_pag0, re.IGNORECASE)
    if m:
        anio_desde = int(m.group(1))
        mes_desde  = int(m.group(2))
        if mes_desde == 12:
            # DESDE es fin del periodo anterior (dic) → extracto pertenece a enero anio+1
            log.debug("[banco_pdf] DESDE mes=12, infiriendo anio+1: %d", anio_desde + 1)
            return anio_desde + 1
        return anio_desde

    # 3. Fallback: primer anio encontrado
    m = re.search(r'\b(20\d{2})\b', texto_pag0)
    if m:
        return int(m.group(1))

    return datetime.now().year


# Patrones para resumen del extracto — keywords de Bancolombia, Davivienda, BBVA, Bogotá, etc.
_PAT_RESUMEN = {
    'SALDO_ANTERIOR': r'SALDO\s+(?:ANTERIOR|INICIAL)\s*:?\s*\$?\s*([\d,\.]+)',
    'TOTAL_ABONOS'  : r'TOTAL\s+(?:ABONOS|CR[EÉ]DITOS?)\s*:?\s*\$?\s*([\d,\.]+)',
    'TOTAL_CARGOS'  : r'TOTAL\s+(?:CARGOS|D[EÉ]BITOS?)\s*:?\s*\$?\s*([\d,\.]+)',
    'SALDO_ACTUAL'  : r'SALDO\s+(?:ACTUAL|FINAL|DISPONIBLE)\s*:?\s*\$?\s*([\d,\.]+)',
}


def _buscar_resumen_en_texto(texto, resumen):
    """Busca patrones de resumen en el texto; no sobreescribe valores ya encontrados."""
    for clave, pat in _PAT_RESUMEN.items():
        if clave in resumen:
            continue
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            v = m.group(1).replace(',', '')
            resumen[clave] = limpiar_num(v)


def _extraer_valor(nums):
    """
    Dado el bloque de números al final de una fila de extracto, devuelve el valor
    de la transaccion (positivo = abono, negativo = cargo).

    Casos:
      1 numero   → es el saldo, no hay valor individual (raro)
      2 numeros  → (valor, saldo)  — formato bancolombia clasico
      3 numeros  → (debito, credito, saldo) — bancos con columnas separadas
                   si debito >0 y credito==0 → cargo (negativo)
                   si credito>0 y debito==0  → abono (positivo)
      4+ numeros → usar los dos ultimos como (valor, saldo)
    """
    if len(nums) == 1:
        return None
    if len(nums) == 2:
        return nums[-2]
    if len(nums) == 3:
        debito_v  = nums[-3]
        credito_v = nums[-2]
        if debito_v and not credito_v:
            return -abs(debito_v)          # cargo
        if credito_v and not debito_v:
            return abs(credito_v)          # abono
        return (credito_v or 0) - (debito_v or 0)
    # 4+ numeros: tomar penultimo
    return nums[-2]


def parsear_banco_pdf(ruta, usar_ocr=False):
    registros = []
    resumen   = {}
    anio_extracto = datetime.now().year

    with pdfplumber.open(ruta) as pdf:
        for n_pag, pag in enumerate(pdf.pages):
            texto = pag.extract_text() or ''
            if len(texto.strip()) <= 30 and usar_ocr and OCR_AVAILABLE:
                texto = ocr_pdf_page(ruta, pag.page_number)


            # Detectar anio del periodo solo en pagina 0
            if n_pag == 0:
                anio_extracto = _detectar_anio_extracto(texto)
                log.debug("[banco_pdf] anio_extracto=%d", anio_extracto)

            # Buscar totales en TODAS las paginas (en Bancolombia aparecen en la ultima)
            _buscar_resumen_en_texto(texto, resumen)

            # --- Intentar extraccion por tabla ---
            tabla = pag.extract_table({
                'vertical_strategy'  : 'lines',
                'horizontal_strategy': 'lines',
            })
            if not tabla:
                tabla = pag.extract_table()

            if tabla:
                for fila in tabla:
                    if not fila:
                        continue
                    celdas = [str(c or '').strip() for c in fila]
                    fecha_raw = next((c for c in celdas if es_fecha_banco(c)), None)
                    if not fecha_raw:
                        continue
                    nums = []
                    for c in reversed(celdas):
                        v = limpiar_num(c)
                        if v is not None:
                            nums.insert(0, v)
                        elif nums:
                            break
                    if len(nums) < 1:
                        continue
                    saldo = nums[-1]
                    valor = _extraer_valor(nums)
                    idx_f = celdas.index(fecha_raw)
                    n_num = len(nums)
                    desc  = ' '.join(c for c in celdas[idx_f + 1:len(celdas) - n_num]
                                     if c and not es_fecha_banco(c))
                    desc  = re.sub(r'\s+', ' ', desc).strip()
                    registros.append({
                        'FECHA_RAW'  : fecha_raw,
                        'FECHA'      : _parse_fecha(fecha_raw, anio_extracto),
                        'DESCRIPCION': desc,
                        'VALOR'      : valor,
                        'SALDO'      : saldo,
                        'TIPO'       : 'ABONO' if (valor or 0) >= 0 else 'CARGO',
                        'PAGINA'     : n_pag + 1,
                    })
            else:
                # --- Fallback linea a linea ---
                for linea in texto.split('\n'):
                    partes = linea.strip().split()
                    if not partes or not es_fecha_banco(partes[0]):
                        continue
                    fecha_raw = partes[0]
                    nums = []; desc_p = []
                    for p in partes[1:]:
                        v = limpiar_num(p)
                        if v is not None:
                            nums.append(v)
                        elif not nums:
                            desc_p.append(p)
                    if not nums:
                        continue
                    saldo = nums[-1]
                    valor = _extraer_valor(nums)
                    registros.append({
                        'FECHA_RAW'  : fecha_raw,
                        'FECHA'      : _parse_fecha(fecha_raw, anio_extracto),
                        'DESCRIPCION': ' '.join(desc_p),
                        'VALOR'      : valor,
                        'SALDO'      : saldo,
                        'TIPO'       : 'ABONO' if (valor or 0) >= 0 else 'CARGO',
                        'PAGINA'     : n_pag + 1,
                    })

    df = pd.DataFrame(registros)
    if df.empty:
        return df, resumen

    df = df[df['VALOR'].notna()].copy()
    df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')
    df = df.drop_duplicates(subset=['FECHA_RAW', 'DESCRIPCION', 'VALOR', 'SALDO'])
    df = df.sort_values('FECHA', na_position='last').reset_index(drop=True)
    df.index += 1

    # Agregar SALDO_FINAL desde la ultima transaccion (mas confiable que regex en texto)
    saldos_validos = df['SALDO'].dropna()
    if not saldos_validos.empty:
        resumen['SALDO_FINAL'] = float(saldos_validos.iloc[-1])

    # Registrar para aprendizaje (nombre_formato lo sobreescribe despachar_ruta)
    try:
        registrar_formato_pdf(
            ruta, 'BANCO', list(df.columns), 'banco_pdf', [],
            banco_detectado='banco_pdf_generico',
        )
    except Exception:
        pass

    return df, resumen
