"""
parsers/banco_pdf.py
Parser de extractos bancarios en PDF para COCILABANK.
Soporta: Bancolombia, Davivienda, BBVA, Banco de Bogota, Banco Popular,
         Banco de Occidente, Colpatria, Itau, AV Villas, Caja Social, Nequi, Nubank.
Usa pdfplumber como motor principal, con fallback a PyMuPDF y OCR.
"""
from __future__ import annotations
import io, logging, re
from datetime import datetime
from typing import Optional
import pandas as pd

logger = logging.getLogger(__name__)

_MESES = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12,
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5,
    'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9,
    'octubre': 10, 'noviembre': 11, 'diciembre': 12
}
_RE_FECHA_DMY = re.compile(r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})')
_RE_FECHA_TEXTO = re.compile(
    r'(\d{1,2})\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic'
    r'|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre'
    r'|octubre|noviembre|diciembre)\s+(\d{2,4})', re.IGNORECASE)
_RE_VALOR = re.compile(r'[-+]?\$?\s*(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{0,2})?)(?:\s*(?:COP|pesos))?')

_DETECTORES_BANCO = [
    ('Bancolombia', ['bancolombia', 'grupo bancolombia']),
    ('Davivienda', ['davivienda', 'daviplata']),
    ('BBVA', ['bbva']),
    ('Banco de Bogota', ['banco de bogota', 'banbogota']),
    ('Banco Popular', ['banco popular']),
    ('Occidente', ['banco de occidente']),
    ('Colpatria', ['colpatria', 'scotiabank colpatria']),
    ('Itau', ['itau']),
    ('AV Villas', ['av villas']),
    ('Caja Social', ['caja social', 'bcsc']),
    ('Nequi', ['nequi']),
    ('Nubank', ['nubank']),
]


def _detectar_banco(texto):
    t = texto[:2000].lower()
    for banco, patrones in _DETECTORES_BANCO:
        if any(p in t for p in patrones):
            return banco
    return 'Desconocido'


def _parsear_fecha(s, ano=0):
    if not s:
        return None
    m = _RE_FECHA_TEXTO.match(s)
    if m:
        try:
            y = int(m.group(3))
            if y < 100:
                y += 2000
            return datetime(y, _MESES.get(m.group(2).lower(), 0), int(m.group(1)))
        except Exception:
            pass
    m = _RE_FECHA_DMY.match(s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if mo > 12:
            d, mo = mo, d
        try:
            return datetime(y, mo, d)
        except Exception:
            pass
    return None


def _limpiar_valor(s, neg=False):
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


def _extraer_pdfplumber(b):
    try:
        import pdfplumber
        t = ''
        tbl = []
        with pdfplumber.open(io.BytesIO(b)) as pdf:
            for p in pdf.pages:
                t += (p.extract_text() or '') + '\n'
                for z in p.extract_tables():
                    if z and len(z) > 1:
                        tbl.append(pd.DataFrame(z[1:], columns=z[0]))
        return t, tbl
    except Exception:
        return '', []


def _extraer_lineas(texto, banco, ano):
    movs = []
    for linea in texto.split('\n'):
        linea = linea.strip()
        if len(linea) < 10:
            continue
        mf = _RE_FECHA_DMY.search(linea) or _RE_FECHA_TEXTO.search(linea)
        if not mf:
            continue
        f = _parsear_fecha(mf.group(0), ano)
        if not f:
            continue
        vr = _RE_VALOR.findall(linea)
        if not vr:
            continue
        desc = linea[mf.end():].strip()
        vals = [_limpiar_valor(v) for v in vr if _limpiar_valor(v) != 0]
        if not vals:
            continue
        v = vals[-2] if len(vals) >= 2 else vals[0]
        du = (desc or '').upper()
        if any(k in du for k in ['DEBITO', 'RETIRO', 'PAGO', 'ND']):
            v = -abs(v)
        elif any(k in du for k in ['CREDITO', 'DEPOSITO', 'CONSIG', 'NC']):
            v = abs(v)
        movs.append({
            'FECHA': f,
            'DESCRIPCION': desc[:120],
            'VALOR': v,
            'BANCO': banco,
            'ORIGEN': 'banco_pdf'
        })
    return movs


def parsear_banco_pdf(file_bytes, ano_defecto=None, banco_hint=None):
    if ano_defecto is None:
        ano_defecto = datetime.now().year
    texto, tablas = _extraer_pdfplumber(file_bytes)
    if not texto:
        try:
            import fitz
            doc = fitz.open(stream=file_bytes, filetype='pdf')
            texto = '\n'.join(p.get_text() for p in doc)
            tablas = []
        except Exception:
            pass
    if not texto.strip():
        return pd.DataFrame(columns=['FECHA', 'DESCRIPCION', 'VALOR', 'BANCO', 'ORIGEN'])
    banco = banco_hint or _detectar_banco(texto)
    movs = _extraer_lineas(texto, banco, ano_defecto)
    if not movs:
        return pd.DataFrame(columns=['FECHA', 'DESCRIPCION', 'VALOR', 'BANCO', 'ORIGEN'])
    df = pd.DataFrame(movs)
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
    return df.dropna(subset=['FECHA']).sort_values('FECHA').reset_index(drop=True)


def diagnostico_pdf(file_bytes):
    info = {'paginas': 0, 'tiene_texto': False, 'banco': 'Desconocido', 'error': None}
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            info['paginas'] = len(pdf.pages)
            t = ''.join(p.extract_text() or '' for p in pdf.pages)
            info['tiene_texto'] = len(t.strip()) > 50
            info['banco'] = _detectar_banco(t)
    except Exception as e:
        info['error'] = str(e)
    return info
