"""
parsers/formato_txt.py — Parsers de auxiliar contable en formato TXT
"""
import re
import logging
from datetime import datetime

import pandas as pd

logger = logging.getLogger(__name__)

# ── Patrones de fecha ──────────────────────────────────────────────────────────
# Formato con separadores: 01/06/2025  01-06-2025  01.06.2025
_RE_FECHA_SEP = re.compile(r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})')
# Formato texto: 01 jun 2025 / 01 junio 2025
_RE_FECHA_TXT = re.compile(
    r'(\d{1,2})\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic'
    r'|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre'
    r'|octubre|noviembre|diciembre)\s+(\d{2,4})', re.IGNORECASE)
# Formato ISO: 2025-06-01
_RE_FECHA_ISO = re.compile(r'(\d{4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})')
# FIX: Formato compacto legado PRN: 01062025 (DDMMYYYY) o 20250601 (YYYYMMDD)
_RE_FECHA_COMPACTA_DMY  = re.compile(r'\b(\d{2})(\d{2})(\d{4})\b')   # DDMMYYYY
_RE_FECHA_COMPACTA_YMD  = re.compile(r'\b(\d{4})(\d{2})(\d{2})\b')   # YYYYMMDD

_MESES = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12,
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5,
    'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9,
    'octubre': 10, 'noviembre': 11, 'diciembre': 12,
}

_RE_VALOR = re.compile(r'[-+]?\$?\s*(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{0,2})?)')


def _parsear_fecha(s):
    """Intenta parsear una fecha desde string, soportando múltiples formatos."""
    if not s:
        return None
    s = s.strip()

    # ISO: 2025-06-01
    m = _RE_FECHA_ISO.match(s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return datetime(y, mo, d)
        except ValueError:
            pass

    # Con separadores: 01/06/2025
    m = _RE_FECHA_SEP.match(s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if mo > 12:
            d, mo = mo, d
        try:
            return datetime(y, mo, d)
        except ValueError:
            pass

    # Texto: 01 jun 2025
    m = _RE_FECHA_TXT.match(s)
    if m:
        d = int(m.group(1))
        mo = _MESES.get(m.group(2).lower(), 0)
        y = int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime(y, mo, d)
        except ValueError:
            pass

    # FIX compacto DDMMYYYY: 01062025
    m = _RE_FECHA_COMPACTA_DMY.match(s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12 and 2000 <= y <= 2100:
            try:
                return datetime(y, mo, d)
            except ValueError:
                pass

    # FIX compacto YYYYMMDD: 20250601
    m = _RE_FECHA_COMPACTA_YMD.match(s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 2000 <= y <= 2100 and 1 <= mo <= 12 and 1 <= d <= 31:
            try:
                return datetime(y, mo, d)
            except ValueError:
                pass

    return None


def _buscar_fecha_en_linea(linea):
    """Busca fecha en una línea, probando todos los patrones en orden de prioridad."""
    # ISO primero (más específico)
    m = _RE_FECHA_ISO.search(linea)
    if m:
        f = _parsear_fecha(m.group(0))
        if f:
            return f, m.end()

    # Con separadores
    m = _RE_FECHA_SEP.search(linea)
    if m:
        f = _parsear_fecha(m.group(0))
        if f:
            return f, m.end()

    # Texto
    m = _RE_FECHA_TXT.search(linea)
    if m:
        f = _parsear_fecha(m.group(0))
        if f:
            return f, m.end()

    # Compacto YYYYMMDD
    m = _RE_FECHA_COMPACTA_YMD.search(linea)
    if m:
        f = _parsear_fecha(m.group(0))
        if f:
            return f, m.end()

    # Compacto DDMMYYYY (FIX PRN legado)
    m = _RE_FECHA_COMPACTA_DMY.search(linea)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12 and 2000 <= y <= 2100:
            try:
                f = datetime(y, mo, d)
                return f, m.end()
            except ValueError:
                pass

    return None, 0


def _limpiar_valor(s):
    """Convierte string de valor colombiano a float."""
    s = str(s).strip().replace('$', '').replace(' ', '')
    neg = s.startswith('-') or s.startswith('(')
    s = s.lstrip('-()+')
    if ',' in s and '.' in s:
        if s.rindex(',') > s.rindex('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except Exception:
        return 0.0


def _detectar_columnas(lineas):
    """Detecta el separador y estructura de columnas del archivo."""
    for sep in [';', '\t', '|', ',']:
        for linea in lineas[:10]:
            if sep in linea and linea.count(sep) >= 2:
                return sep
    return None  # columnas fijas


def parsear_txt(file_bytes, tipo='auxiliar', encoding='utf-8'):
    """
    Parser principal para archivos TXT/PRN de auxiliares contables y extractos.
    Soporta:
      - Columnas delimitadas (;  |  TAB  ,)
      - Columnas fijas (ancho fijo)
      - Formato compacto PRN legado (DDMMYYYY sin separador)
    Retorna DataFrame con columnas FECHA, DESCRIPCION, DEBITO, CREDITO, VALOR, ORIGEN
    """
    try:
        texto = file_bytes.decode(encoding, errors='replace')
    except Exception:
        texto = file_bytes.decode('latin-1', errors='replace')

    lineas = [l.rstrip() for l in texto.splitlines() if l.strip()]
    if not lineas:
        return pd.DataFrame(columns=['FECHA', 'DESCRIPCION', 'DEBITO', 'CREDITO', 'VALOR', 'ORIGEN'])

    sep = _detectar_columnas(lineas)
    movs = []

    if sep:
        # ── Formato delimitado ─────────────────────────────────────────────────
        encabezado = None
        for linea in lineas:
            cols = [c.strip() for c in linea.split(sep)]
            if encabezado is None:
                # Detectar si es línea de encabezado
                muestra = ' '.join(cols).upper()
                if any(k in muestra for k in ['FECHA', 'DATE', 'DATA', 'DIA']):
                    encabezado = [c.upper() for c in cols]
                    continue
                else:
                    # Primera línea sin encabezado: intentar parsear directamente
                    encabezado = []

            if not cols or len(cols) < 2:
                continue

            # Buscar fecha en columnas
            fecha = None
            desc = ''
            debito = credito = 0.0

            for i, col in enumerate(cols):
                if fecha is None:
                    f, _ = _buscar_fecha_en_linea(col)
                    if f:
                        fecha = f
                        continue
                vals = _RE_VALOR.findall(col)
                if vals and not desc:
                    # Primera columna no-fecha con texto largo → descripción
                    if len(col) > 5 and not vals:
                        desc = col[:120]

            if fecha is None:
                continue

            # Extraer descripción y valores numéricos
            nums = []
            for col in cols:
                vv = _RE_VALOR.findall(col)
                if vv:
                    nums.extend([_limpiar_valor(v) for v in vv if _limpiar_valor(v) != 0])
                elif len(col) > 4 and not _RE_FECHA_SEP.search(col) and not _RE_FECHA_ISO.search(col):
                    if not desc:
                        desc = col[:120]

            if len(nums) >= 2:
                debito, credito = nums[0], nums[1]
            elif len(nums) == 1:
                v = nums[0]
                if v > 0:
                    credito = v
                else:
                    debito = abs(v)

            valor = credito - debito if credito or debito else (nums[0] if nums else 0)

            movs.append({
                'FECHA': fecha,
                'DESCRIPCION': desc.strip()[:120],
                'DEBITO': debito,
                'CREDITO': credito,
                'VALOR': valor,
                'ORIGEN': 'txt'
            })

    else:
        # ── Columnas fijas / PRN legado ────────────────────────────────────────
        for linea in lineas:
            if len(linea) < 8:
                continue

            fecha, pos = _buscar_fecha_en_linea(linea)
            if not fecha:
                continue

            resto = linea[pos:].strip()
            nums = [_limpiar_valor(v) for v in _RE_VALOR.findall(resto) if _limpiar_valor(v) != 0]
            # Descripción: texto no numérico después de la fecha
            desc_match = re.sub(r'[\d.,\-\+\$\(\)]+', ' ', resto).strip()
            desc = ' '.join(desc_match.split())[:120]

            debito = credito = 0.0
            if len(nums) >= 2:
                debito, credito = abs(nums[0]), abs(nums[1])
            elif len(nums) == 1:
                v = nums[0]
                if v < 0:
                    debito = abs(v)
                else:
                    credito = v

            valor = credito - debito

            movs.append({
                'FECHA': fecha,
                'DESCRIPCION': desc,
                'DEBITO': debito,
                'CREDITO': credito,
                'VALOR': valor,
                'ORIGEN': 'txt_prn'
            })

    if not movs:
        return pd.DataFrame(columns=['FECHA', 'DESCRIPCION', 'DEBITO', 'CREDITO', 'VALOR', 'ORIGEN'])

    df = pd.DataFrame(movs)
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    for col in ['DEBITO', 'CREDITO', 'VALOR']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df.dropna(subset=['FECHA']).sort_values('FECHA').reset_index(drop=True)
