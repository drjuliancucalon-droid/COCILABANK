"""
parsers/despachador.py — Registro extensible de formatos y despachador de parsers
Cómo añadir un nuevo formato:
  1. Escribir fn_detectar(ruta, muestra_texto) -> float  [0.0–1.0 confianza]
  2. Escribir fn_parsear(ruta, usar_ocr)       -> (DataFrame, dict_meta)
  3. Agregar una entrada al final de REGISTRO_FORMATOS con tipo/ext correctos.
"""
import re
import os
import io
import unicodedata as _ud
import tempfile
import logging
from pathlib import Path

import pandas as pd
import pdfplumber

from utils.pdf_diagnostico import diagnosticar_pdf, OCR_AVAILABLE
from storage.db import registrar_formato_pdf, buscar_formato_rapido
from parsers.banco_pdf import parsear_banco_pdf, limpiar_num
from parsers.auxiliar_pdf import parsear_auxiliar_pdf
from parsers.formatos_csv import parsear_banco_csv, parsear_auxiliar_csv, _col
from parsers.formato_txt import parsear_banco_txt, parsear_auxiliar_txt

log = logging.getLogger(__name__)

# ── Helpers internos ──────────────────────────────────────────────────────────
def _norm(s):
    return _ud.normalize('NFKD', s.lower()).encode('ascii', 'ignore').decode()

def muestra_texto(ruta, ext, n_lineas=50):
    """Texto de muestra para detección rápida (sin parsear el archivo completo)."""
    try:
        if ext == '.pdf':
            with pdfplumber.open(ruta) as pdf:
                return (pdf.pages[0].extract_text() or '') if pdf.pages else ''
        elif ext in ('.xlsx', '.xls'):
            # Excel binario: leer celdas de la primera hoja y convertir a texto
            df_head = pd.read_excel(ruta, header=None, nrows=n_lineas)
            return '\n'.join(
                '\t'.join(str(v) for v in row if pd.notna(v))
                for _, row in df_head.iterrows()
            )
        else:
            with open(ruta, 'r', encoding='latin1', errors='replace') as f:
                return ''.join(f.readlines()[:n_lineas])
    except Exception as e:
        log.error("[muestra_texto] %s: %s", ruta, e, exc_info=True)
        return ''

# Alias para compatibilidad interna
_muestra_texto = muestra_texto

def _header_row_csv(ruta, encoding='latin1'):
    """Fila donde empieza el encabezado real del CSV."""
    claves = {'documento', 'fecha', 'concepto', 'debito', 'credito',
              'valor', 'saldo', 'descripcion'}
    with open(ruta, 'r', encoding=encoding, errors='replace') as f:
        for i, linea in enumerate(f):
            hits = sum(1 for k in claves if k in _norm(linea))
            if hits >= 3:
                return i
    return 0

def _leer_csv_inteligente(ruta):
    """Lee CSV/Excel saltando metadatos y capturando saldo inicial si existe."""
    ext = Path(ruta).suffix.lower()

    # ── Excel binario: pd.read_csv no funciona, usar read_excel ──────────────
    if ext in ('.xlsx', '.xls'):
        try:
            df = pd.read_excel(ruta, header=0)
            df = df.dropna(how='all').reset_index(drop=True)
            # Intentar detectar saldo inicial en la primera hoja como número
            saldo_ini = None
            try:
                xls_raw = pd.read_excel(ruta, header=None, nrows=10)
                for _, row in xls_raw.iterrows():
                    txt = ' '.join(str(v) for v in row.values)
                    m = re.search(r'Saldo\s+Inicial[:\s]+([\d,\.]+)', txt, re.I)
                    if m:
                        saldo_ini = limpiar_num(m.group(1).replace(',', ''))
                        break
            except Exception:
                pass
            return df, saldo_ini
        except Exception as e:
            log.error("[_leer_csv_inteligente] Error leyendo Excel %s: %s", ruta, e)
            raise

    # ── CSV de texto plano ────────────────────────────────────────────────────
    skip = _header_row_csv(ruta)
    saldo_ini = None
    with open(ruta, 'r', encoding='latin1', errors='replace') as f:
        for linea in f.readlines()[:skip + 3]:
            m = re.search(r'Saldo\s+Inicial[:\s]+([\d,\.]+)', linea, re.I)
            if m:
                saldo_ini = limpiar_num(m.group(1).replace(',', ''))
                break
    df = pd.read_csv(ruta, encoding='latin1', sep=None, engine='python',
                     skiprows=skip, header=0)
    df = df.dropna(how='all').reset_index(drop=True)
    return df, saldo_ini

# ── Detectores ────────────────────────────────────────────────────────────────

def _det_bancolombia_pdf(ruta, m):
    hits = [
        bool(re.search(r'ESTADO\s+DE\s+CUENTA',      m, re.I)),
        bool(re.search(r'SALDO\s+ANTERIOR',           m, re.I)),
        bool(re.search(r'TOTAL\s+ABONOS',             m, re.I)),
        bool(re.search(r'TOTAL\s+CARGOS',             m, re.I)),
        bool(re.search(r'BANCOLOMBIA',                m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_siigo_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'Auxiliares\s*[-–]\s*Plan\s+de\s+Cuentas', m, re.I)),
        bool(re.search(r'Nit\.?\s*\d{6,}',            m, re.I)),
        bool(re.search(r'(?:CE|CON|NC|FA|RI|REC|CXP|EG)-\d+',          m)),
        bool(re.search(r'D[eé]bitos?',                m, re.I)),
        bool(re.search(r'Cr[eé]ditos?',               m, re.I)),
        bool(re.search(r'Saldo\s+Inicial',            m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_helisa_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'HELISA',                     m, re.I)),
        bool(re.search(r'Libro\s+Auxiliar',           m, re.I)),
        bool(re.search(r'(?:CE|CON|NC|RE|RG|FA|EG|AJ)-\d+',   m)),
        bool(re.search(r'D[eé]bito|Cr[eé]dito',      m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_world_office_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'World\s*Office|WO\s+\d',    m, re.I)),
        bool(re.search(r'Comprobante',                m, re.I)),
        bool(re.search(r'D[eé]bito|Cr[eé]dito',      m, re.I)),
        bool(re.search(r'\d{4}-\d{2}-\d{2}',         m)),
    ]
    return sum(hits) / len(hits)

def _det_siigo_aux_pdf(ruta, m):
    """Detector especifico para exportaciones PDF de SIIGO Auxiliares de Plan de Cuentas."""
    hits = [
        bool(re.search(r'(?:CON|CE|CG|NC|RE|RG|FA|RI|EG|AJ)-\d+',    m)),         # codigos de documento
        bool(re.search(r'D[e\xe9]bitos?',                  m, re.I)),   # columna Debitos
        bool(re.search(r'Cr[e\xe9]ditos?',                 m, re.I)),   # columna Creditos
        bool(re.search(r'Saldo\s+Inicial|Saldo\s+Final',  m, re.I)),   # saldos
        bool(re.search(r'Auxiliares|Plan\s+de\s+Cuentas', m, re.I)),   # encabezado SIIGO
        bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}',         m)),         # formato fecha DD/MM/YYYY
    ]
    return sum(hits) / len(hits)

def _det_aux_pdf_generico(ruta, m):
    """Detector generico para cualquier auxiliar en PDF con prefijos de documento."""
    hits = [
        bool(re.search(r'(?:CON|CE|CG|NC|RE|RG)-\d+',    m)),
        bool(re.search(r'D[e\xe9]bitos?',                  m, re.I)),
        bool(re.search(r'Cr[e\xe9]ditos?',                 m, re.I)),
        bool(re.search(r'Saldo',                           m, re.I)),
        bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}',         m)),
    ]
    return sum(hits) / len(hits)

def _det_banco_csv_generico(ruta, m):
    hits = [
        bool(re.search(r'\bfecha\b',                  m, re.I)),
        bool(re.search(r'\bvalor\b|\bmonto\b',        m, re.I)),
        bool(re.search(r'\bsaldo\b',                  m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_aux_csv_generico(ruta, m):
    hits = [
        bool(re.search(r'\bdocumento\b',              m, re.I)),
        bool(re.search(r'\bfecha\b',                  m, re.I)),
        bool(re.search(r'\bconcepto\b|\bdescripci',   m, re.I)),
        bool(re.search(r'\bd[eé]bito\b|\bhaber\b',   m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_banco_txt(ruta, m):
    lineas_fecha = sum(1 for l in m.split('\n')
                       if re.match(r'^\d{1,2}/\d{2}\s', l.strip()))
    return min(1.0, lineas_fecha / 5)

def _det_aux_txt(ruta, m):
    hits = [
        bool(re.search(r'(?:CON|CE|NC|FA|RI|REC|EG)-\d+',         m)),
        bool(re.search(r'\d{1,2}/\d{2}/\d{4}',       m)),
    ]
    return sum(hits) / len(hits)


# ── Detectores bancos colombianos adicionales ────────────────────────────────

def _det_davivienda_pdf(ruta, m):
    hits = [
        bool(re.search(r'DAVIVIENDA',          m, re.I)),
        bool(re.search(r'DAVIPLATA',            m, re.I)),
        bool(re.search(r'BANCO\s+DAVIVIENDA',  m, re.I)),
        bool(re.search(r'MOVIMIENTOS|EXTRACTO', m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_bbva_pdf(ruta, m):
    hits = [
        bool(re.search(r'\bBBVA\b',                m, re.I)),
        bool(re.search(r'BANCO\s+BILBAO',           m, re.I)),
        bool(re.search(r'SALDO|MOVIMIENTOS|EXTRACTO', m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_bogota_pdf(ruta, m):
    hits = [
        bool(re.search(r'BANCO\s+DE\s+BOGOT',      m, re.I)),
        bool(re.search(r'GRUPO\s+AVAL',             m, re.I)),
        bool(re.search(r'SALDO|MOVIMIENTO',         m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_nequi_pdf(ruta, m):
    hits = [
        bool(re.search(r'\bNEQUI\b',                m, re.I)),
        bool(re.search(r'MOVIMIENTOS|HISTORIAL',    m, re.I)),
        bool(re.search(r'SALDO',                    m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_colpatria_pdf(ruta, m):
    hits = [
        bool(re.search(r'SCOTIABANK|COLPATRIA',     m, re.I)),
        bool(re.search(r'SALDO|MOVIMIENTO',         m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_occidente_pdf(ruta, m):
    hits = [
        bool(re.search(r'BANCO\s+DE\s+OCCIDENTE',  m, re.I)),
        bool(re.search(r'\bOCCIDENTE\b',            m, re.I)),
        bool(re.search(r'SALDO|MOVIMIENTO',         m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_popular_pdf(ruta, m):
    hits = [
        bool(re.search(r'BANCO\s+POPULAR',          m, re.I)),
        bool(re.search(r'SALDO|MOVIMIENTO',         m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_banco_pdf_generico(ruta, m):
    """Detector de ultimo recurso para cualquier PDF bancario colombiano."""
    hits = [
        bool(re.search(r'\d{2}[-/]\d{2}[-/]\d{4}',               m)),
        bool(re.search(r'SALDO',                                    m, re.I)),
        bool(re.search(r'D.BITO|CR.DITO|VALOR|MONTO',              m, re.I)),
        bool(re.search(r'ESTADO\s+DE\s+CUENTA|EXTRACTO|MOVIMIENTO', m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_contapyme_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'ContaPyme|CONTAPYME',         m, re.I)),
        bool(re.search(r'Tipo\s+Comprobante|TipoDoc',  m, re.I)),
        bool(re.search(r'D.bito|Cr.dito',              m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_alegra_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'\bAlegra\b',                        m, re.I)),
        bool(re.search(r'Tipo\s+(?:movimiento|transacci)',    m, re.I)),
        bool(re.search(r'Ingreso|Egreso|Gasto',               m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_loggro_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'\bLoggro\b',           m, re.I)),
        bool(re.search(r'D.bito|Cr.dito',       m, re.I)),
    ]
    return sum(hits) / len(hits)

def _det_monica_aux_csv(ruta, m):
    hits = [
        bool(re.search(r'\bMONICA\b',           m, re.I)),
        bool(re.search(r'Debe|Haber',            m, re.I)),
    ]
    return sum(hits) / len(hits)


# ── Parsers de formato ─────────────────────────────────────────────────────────

def _par_bancolombia_pdf(ruta, usar_ocr):
    return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)

def _par_siigo_aux_csv(ruta, usar_ocr):
    df_raw, saldo_ini = _leer_csv_inteligente(ruta)
    df, meta = parsear_auxiliar_csv(df_raw)
    if saldo_ini is not None:
        meta['SALDO_INICIAL'] = saldo_ini
        meta['SALDO_FINAL'] = saldo_ini + meta.get('TOTAL_DEBITOS',0) - meta.get('TOTAL_CREDITOS',0)
    return df, meta

def _par_helisa_aux_csv(ruta, usar_ocr):
    # Helisa exporta similar a SIIGO; reutilizamos misma lógica
    return _par_siigo_aux_csv(ruta, usar_ocr)

def _par_world_office_aux_csv(ruta, usar_ocr):
    # World Office: columnas pueden llamarse "Debe"/"Haber" en lugar de Débito/Crédito
    df_raw, saldo_ini = _leer_csv_inteligente(ruta)
    df, meta = parsear_auxiliar_csv(df_raw)
    if saldo_ini is not None:
        meta['SALDO_INICIAL'] = saldo_ini
        meta['SALDO_FINAL'] = saldo_ini + meta.get('TOTAL_DEBITOS',0) - meta.get('TOTAL_CREDITOS',0)
    return df, meta

def _par_aux_pdf_generico(ruta, usar_ocr):
    return parsear_auxiliar_pdf(ruta, usar_ocr=usar_ocr)

def _par_banco_csv_generico(ruta, usar_ocr):
    df_raw, _ = _leer_csv_inteligente(ruta)
    return parsear_banco_csv(df_raw)

def _par_aux_csv_generico(ruta, usar_ocr):
    df_raw, saldo_ini = _leer_csv_inteligente(ruta)
    df, meta = parsear_auxiliar_csv(df_raw)
    if saldo_ini is not None:
        meta['SALDO_INICIAL'] = saldo_ini
        meta['SALDO_FINAL'] = saldo_ini + meta.get('TOTAL_DEBITOS',0) - meta.get('TOTAL_CREDITOS',0)
    return df, meta

def _par_banco_txt(ruta, usar_ocr):
    with open(ruta, 'r', encoding='utf-8', errors='replace') as f:
        return parsear_banco_txt(f.read())

def _par_aux_txt(ruta, usar_ocr):
    with open(ruta, 'r', encoding='utf-8', errors='replace') as f:
        return parsear_auxiliar_txt(f.read())


# ── Parsers bancos colombianos adicionales ────────────────────────────────────

def _par_davivienda_pdf(ruta, usar_ocr):    return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_bbva_pdf(ruta, usar_ocr):          return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_bogota_pdf(ruta, usar_ocr):        return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_nequi_pdf(ruta, usar_ocr):         return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_colpatria_pdf(ruta, usar_ocr):     return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_occidente_pdf(ruta, usar_ocr):     return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_popular_pdf(ruta, usar_ocr):       return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)
def _par_banco_pdf_generico(ruta, usar_ocr): return parsear_banco_pdf(ruta, usar_ocr=usar_ocr)

def _par_contapyme_aux_csv(ruta, usar_ocr):
    df_raw, saldo_ini = _leer_csv_inteligente(ruta)
    rename = {c: 'Debito' for c in df_raw.columns if c.lower().strip() == 'debe'}
    rename.update({c: 'Credito' for c in df_raw.columns if c.lower().strip() == 'haber'})
    if rename: df_raw = df_raw.rename(columns=rename)
    df, meta = parsear_auxiliar_csv(df_raw)
    if saldo_ini is not None: meta['SALDO_INICIAL'] = saldo_ini
    return df, meta

def _par_alegra_aux_csv(ruta, usar_ocr):
    df_raw, saldo_ini = _leer_csv_inteligente(ruta)
    rename = {}
    for col in df_raw.columns:
        n = col.lower().strip()
        if 'ingreso' in n:                 rename[col] = 'Credito'
        elif 'egreso' in n or 'gasto' in n: rename[col] = 'Debito'
        elif 'detail' in n:                rename[col] = 'Concepto'
    if rename: df_raw = df_raw.rename(columns=rename)
    df, meta = parsear_auxiliar_csv(df_raw)
    if saldo_ini is not None: meta['SALDO_INICIAL'] = saldo_ini
    return df, meta

def _par_loggro_aux_csv(ruta, usar_ocr):  return _par_siigo_aux_csv(ruta, usar_ocr)

def _par_monica_aux_csv(ruta, usar_ocr):
    df_raw, saldo_ini = _leer_csv_inteligente(ruta)
    rename = {c: 'Debito' for c in df_raw.columns if c.lower().strip() == 'debe'}
    rename.update({c: 'Credito' for c in df_raw.columns if c.lower().strip() == 'haber'})
    if rename: df_raw = df_raw.rename(columns=rename)
    df, meta = parsear_auxiliar_csv(df_raw)
    if saldo_ini is not None: meta['SALDO_INICIAL'] = saldo_ini
    return df, meta


# ── Registro principal ────────────────────────────────────────────────────────
# Orden importa: formatos más específicos primero dentro del mismo tipo/ext.
# El despachador elige el de mayor confianza, no el primero.

REGISTRO_FORMATOS = [
    {
        'nombre'  : 'Bancolombia — PDF Estado de Cuenta',
        'tipo'    : 'BANCO',
        'ext'     : ['.pdf'],
        'detectar': _det_bancolombia_pdf,
        'parsear' : _par_bancolombia_pdf,
    },
    {
        'nombre'  : 'Banco — CSV/Excel con encabezados estándar',
        'tipo'    : 'BANCO',
        'ext'     : ['.csv', '.xlsx', '.xls'],
        'detectar': _det_banco_csv_generico,
        'parsear' : _par_banco_csv_generico,
    },
    {
        'nombre'  : 'Banco — TXT líneas fecha/descripción/valor',
        'tipo'    : 'BANCO',
        'ext'     : ['.txt'],
        'detectar': _det_banco_txt,
        'parsear' : _par_banco_txt,
    },
    {
        'nombre'  : 'SIIGO — CSV Auxiliares Plan de Cuentas',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.csv'],
        'detectar': _det_siigo_aux_csv,
        'parsear' : _par_siigo_aux_csv,
    },
    {
        'nombre'  : 'Helisa — CSV Libro Auxiliar',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.csv', '.xlsx', '.xls'],
        'detectar': _det_helisa_aux_csv,
        'parsear' : _par_helisa_aux_csv,
    },
    {
        'nombre'  : 'World Office — CSV Auxiliar',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.csv', '.xlsx', '.xls'],
        'detectar': _det_world_office_aux_csv,
        'parsear' : _par_world_office_aux_csv,
    },
    {
        'nombre'  : 'SIIGO — PDF Auxiliares Plan de Cuentas',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.pdf'],
        'detectar': _det_siigo_aux_pdf,
        'parsear' : _par_aux_pdf_generico,
    },
    {
        'nombre'  : 'Auxiliar Contable — PDF (CON/CE/NC)',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.pdf'],
        'detectar': _det_aux_pdf_generico,
        'parsear' : _par_aux_pdf_generico,
    },
    {
        'nombre'  : 'Auxiliar Contable — CSV/Excel genérico',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.csv', '.xlsx', '.xls'],
        'detectar': _det_aux_csv_generico,
        'parsear' : _par_aux_csv_generico,
    },
    {'nombre':'Davivienda — PDF Estado de Cuenta',      'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_davivienda_pdf,    'parsear':_par_davivienda_pdf},
    {'nombre':'BBVA Colombia — PDF Estado de Cuenta',     'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_bbva_pdf,           'parsear':_par_bbva_pdf},
    {'nombre':'Banco de Bogota — PDF Estado de Cuenta',   'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_bogota_pdf,         'parsear':_par_bogota_pdf},
    {'nombre':'Nequi — PDF Movimientos',                  'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_nequi_pdf,          'parsear':_par_nequi_pdf},
    {'nombre':'Scotiabank Colpatria — PDF',               'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_colpatria_pdf,      'parsear':_par_colpatria_pdf},
    {'nombre':'Banco de Occidente — PDF',                 'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_occidente_pdf,      'parsear':_par_occidente_pdf},
    {'nombre':'Banco Popular — PDF',                      'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_popular_pdf,        'parsear':_par_popular_pdf},
    {'nombre':'Banco Colombiano — PDF (generico)',         'tipo':'BANCO',    'ext':['.pdf'],                    'detectar':_det_banco_pdf_generico, 'parsear':_par_banco_pdf_generico},
    {'nombre':'ContaPyme — CSV/Excel Auxiliar',           'tipo':'AUXILIAR', 'ext':['.csv','.xlsx','.xls'],    'detectar':_det_contapyme_aux_csv,  'parsear':_par_contapyme_aux_csv},
    {'nombre':'Alegra — CSV Movimientos',                 'tipo':'AUXILIAR', 'ext':['.csv','.xlsx','.xls'],    'detectar':_det_alegra_aux_csv,     'parsear':_par_alegra_aux_csv},
    {'nombre':'Loggro — CSV/Excel Auxiliar',              'tipo':'AUXILIAR', 'ext':['.csv','.xlsx','.xls'],    'detectar':_det_loggro_aux_csv,     'parsear':_par_loggro_aux_csv},
    {'nombre':'Monica — CSV/TXT Auxiliar',                'tipo':'AUXILIAR', 'ext':['.csv','.txt'],            'detectar':_det_monica_aux_csv,     'parsear':_par_monica_aux_csv},
    {
        'nombre'  : 'Auxiliar Contable — TXT',
        'tipo'    : 'AUXILIAR',
        'ext'     : ['.txt'],
        'detectar': _det_aux_txt,
        'parsear' : _par_aux_txt,
    },
]

# ── Despachador ───────────────────────────────────────────────────────────────
def despachar_ruta(ruta, tipo, ext, usar_ocr):
    """Elige el mejor formato registrado y lo parsea.
    Primero intenta un lookup rapido por nombre de archivo aprendido previamente.
    Si el archivo fue procesado >= 2 veces con el mismo formato, lo reutiliza
    directamente sin correr todos los detectores.
    """
    nombre_arch = Path(ruta).name
    candidatos  = [f for f in REGISTRO_FORMATOS if f['tipo'] == tipo and ext in f['ext']]
    if not candidatos:
        raise ValueError(f"Formato no soportado para tipo={tipo}, extension={ext}")

    # ── Lookup rapido: ya conocemos este archivo? ─────────────────────────────
    mejor = None
    conf  = 0
    fmt_aprendido = buscar_formato_rapido(nombre_arch, tipo)
    if fmt_aprendido:
        match = next((f for f in candidatos if f['nombre'] == fmt_aprendido), None)
        if match:
            mejor = match
            conf  = 100
            log.debug("[despachar] cache hit: %s -> %s", nombre_arch, fmt_aprendido)

    # ── Si no hay cache, correr todos los detectores ──────────────────────────
    if mejor is None:
        muestra      = _muestra_texto(ruta, ext)
        puntuaciones = [(f, f['detectar'](ruta, muestra)) for f in candidatos]
        mejor, conf_f = max(puntuaciones, key=lambda x: x[1])
        conf = round(conf_f * 100)

    df, meta = mejor['parsear'](ruta, usar_ocr)

    # ── Registrar el formato ganador para aprendizaje futuro ──────────────────
    try:
        cols = list(df.columns) if df is not None and not df.empty else []
        registrar_formato_pdf(
            nombre_arch, tipo, cols, '', [],
            banco_detectado = mejor['nombre'],
            nombre_formato  = mejor['nombre'],
        )
    except Exception as _reg_err:
        log.debug("[despachar] registrar_formato_pdf: %s", _reg_err)

    return df, meta, mejor['nombre'], conf

# Alias interno original
_despachar = despachar_ruta


# ── Función unificada de carga ────────────────────────────────────────────────
def cargar_y_parsear_uploaded_file(uploaded_file, tipo, usar_ocr=False):
    """Función pública: recibe UploadedFile de Streamlit, detecta y parsea."""
    nombre = uploaded_file.name
    ext    = Path(nombre).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(uploaded_file.getvalue())
        ruta = tmp.name
    try:
        df, meta, nombre_formato, confianza = despachar_ruta(ruta, tipo, ext, usar_ocr)
        if ext == '.pdf':
            leg = diagnosticar_pdf(ruta, tipo)
        else:
            leg = {
                'pct_estimado_datos': 100.0,
                'calidad': '\U0001f7e2 EXCELENTE',
                'advertencias': [],
                'ocr_usado': False,
            }
        return df, meta, (
            leg['pct_estimado_datos'],
            leg['calidad'],
            leg.get('advertencias', []),
            nombre_formato,
            confianza,
        )
    finally:
        try:
            os.unlink(ruta)
        except Exception:
            pass

cargar_y_parsear = cargar_y_parsear_uploaded_file
