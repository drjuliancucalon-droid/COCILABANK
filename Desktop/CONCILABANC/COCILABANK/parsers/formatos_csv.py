"""
parsers/formatos_csv.py — Parsers CSV/Excel para banco y auxiliar contable
Soporta: genérico, SIIGO, Helisa, WorldOffice
"""
import re
import unicodedata
import logging

import pandas as pd

from engine.columna import determinar_columna
from storage.db import registrar_formato_pdf

log = logging.getLogger(__name__)


def _col(df, *palabras):
    """Busca columna por palabras clave (normaliza acentos)."""
    def norm(s):
        return unicodedata.normalize('NFKD', s.lower()).encode('ascii', 'ignore').decode()
    palabras_norm = [norm(p) for p in palabras]
    return next((c for c in df.columns if any(p in norm(c) for p in palabras_norm)), None)


def limpiar_num(t):
    """Convierte texto numérico a float. Importado aquí para evitar dependencia circular."""
    from parsers.banco_pdf import limpiar_num as _ln
    return _ln(t)

def parsear_banco_csv(df):
    registros = []
    resumen = {}
    # Detectar columnas por nombre
    col_fecha = next((c for c in df.columns if 'fecha' in c.lower()), None)
    col_desc  = next((c for c in df.columns if 'descrip' in c.lower() or 'concepto' in c.lower()), None)
    col_valor = next((c for c in df.columns if 'valor' in c.lower() or 'monto' in c.lower()), None)
    col_saldo = next((c for c in df.columns if 'saldo' in c.lower()), None)

    if not col_fecha: raise ValueError("No se encontró columna de fecha en el archivo del banco")
    for _, row in df.iterrows():
        fecha = str(row[col_fecha])
        try:
            fecha_dt = pd.to_datetime(fecha, dayfirst=True, errors='coerce')
        except:
            fecha_dt = pd.NaT
        desc = str(row[col_desc]) if col_desc else ''
        valor = limpiar_num(row[col_valor]) if col_valor else 0
        saldo = limpiar_num(row[col_saldo]) if col_saldo else None
        registros.append({
            'FECHA_RAW': fecha, 'FECHA': fecha_dt,
            'DESCRIPCION': desc, 'VALOR': valor, 'SALDO': saldo,
            'TIPO': 'ABONO' if (valor or 0) >= 0 else 'CARGO'
        })
    df_out = pd.DataFrame(registros)
    df_out = df_out[df_out['VALOR'].notna()]
    df_out['VALOR'] = pd.to_numeric(df_out['VALOR'], errors='coerce')
    df_out = df_out.drop_duplicates()
    df_out = df_out.sort_values('FECHA', na_position='last').reset_index(drop=True)
    df_out.index += 1
    # Calcular totales
    resumen['TOTAL_ABONOS'] = df_out[df_out['VALOR'] > 0]['VALOR'].sum()
    resumen['TOTAL_CARGOS'] = df_out[df_out['VALOR'] < 0]['VALOR'].sum()
    if col_saldo:
        resumen['SALDO_INICIAL'] = df_out.iloc[0]['SALDO'] if not df_out.empty else 0
        resumen['SALDO_FINAL']   = df_out.iloc[-1]['SALDO'] if not df_out.empty else 0
    else:
        resumen['SALDO_INICIAL'] = resumen['SALDO_FINAL'] = 0
    return df_out, resumen

def parsear_auxiliar_csv(df):
    registros = []
    meta = {}
    col_doc = _col(df, 'documento', 'doc')
    col_fec = _col(df, 'fecha')
    col_con = _col(df, 'concepto', 'descrip')
    col_deb = _col(df, 'debito', 'débito', 'debe', 'debitos', 'débitos')
    col_cre = _col(df, 'credito', 'crédito', 'haber', 'creditos', 'créditos')

    if not col_fec: raise ValueError("No se encontró columna de fecha en el auxiliar")

    # Filtrar filas que no sean movimientos reales (sin doc o sin fecha válida)
    for _, row in df.iterrows():
        doc = str(row[col_doc]).strip() if col_doc else ''
        fecha = str(row[col_fec]).strip()
        if not fecha or fecha in ('nan', 'NaT', 'Fecha'): continue
        try: fecha_dt = pd.to_datetime(fecha, dayfirst=True, errors='coerce')
        except: fecha_dt = pd.NaT
        if pd.isna(fecha_dt): continue
        concepto = str(row[col_con]).strip() if col_con else ''
        debito  = limpiar_num(row[col_deb])  if col_deb else None
        credito = limpiar_num(row[col_cre]) if col_cre else None
        # Ignorar filas de subtotales/encabezados sin datos reales.
        # Un doc vacío o 'nan' junto a un concepto vacío indica fila de totales.
        # Documentos numéricos (ej. 250201) o con prefijo (CE-123) son válidos.
        _doc_vacio = not doc or doc.lower() in ('nan', 'none', 'nat', '0')
        _con_vacio = not concepto or concepto.lower() in ('nan', 'none', 'nat')
        if _doc_vacio and _con_vacio:
            continue
        col_asiento = determinar_columna(concepto, doc)
        registros.append({
            'DOCUMENTO': doc, 'FECHA_RAW': fecha, 'FECHA': fecha_dt,
            'CONCEPTO': concepto, 'DEBITO': debito, 'CREDITO': credito,
            'COLUMNA': col_asiento, 'VALOR_NETO': (debito or 0) - (credito or 0)
        })
    df_out = pd.DataFrame(registros)
    if not df_out.empty:
        # CSV es exportación autoritativa del sistema contable; no deduplicar
        # (misma cuenta puede tener múltiples líneas idénticas en un comprobante)
        df_out = df_out.sort_values('FECHA', na_position='last').reset_index(drop=True)
        df_out.index += 1
    meta['TOTAL_DEBITOS']  = df_out['DEBITO'].sum()  if not df_out.empty else 0
    meta['TOTAL_CREDITOS'] = df_out['CREDITO'].sum() if not df_out.empty else 0
    # Buscar saldo inicial en el CSV (fila con 'Saldo Inicial:')
    if 'SALDO_INICIAL' not in meta:
        meta['SALDO_INICIAL'] = 0
    if 'SALDO_FINAL' not in meta:
        meta['SALDO_FINAL'] = meta['SALDO_INICIAL'] + meta['TOTAL_DEBITOS'] - meta['TOTAL_CREDITOS']
    try:
        registrar_formato_pdf('', 'AUXILIAR', list(df_out.columns), 'csv', [],
                              banco_detectado='auxiliar_csv')
    except Exception:
        pass
    return df_out, meta

def parsear_banco_txt(texto):
    registros = []
    resumen = {}
    # Detectar año en el texto (evita hardcoding)
    _m_anio_t = re.search(r'\b(20\d{2})\b', texto or '')
    anio_extracto = int(_m_anio_t.group(1)) if _m_anio_t else datetime.now().year
    for linea in texto.split('\n'):
        partes = linea.strip().split()
        if not partes or not es_fecha_banco(partes[0]): continue
        fecha_raw = partes[0]
        nums = []; desc_p = []
        for p in partes[1:]:
            v = limpiar_num(p)
            if v is not None: nums.append(v)
            elif not nums: desc_p.append(p)
        if not nums: continue
        saldo = nums[-1]
        valor = nums[-2] if len(nums) >= 2 else nums[0]
        registros.append({
            'FECHA_RAW': fecha_raw,
            'FECHA': pd.to_datetime(f'{anio_extracto}/' + fecha_raw, format='%Y/%d/%m', errors='coerce'),
            'DESCRIPCION': ' '.join(desc_p), 'VALOR': valor, 'SALDO': saldo,
            'TIPO': 'ABONO' if (valor or 0) >= 0 else 'CARGO'
        })
    df = pd.DataFrame(registros)
    if not df.empty:
        df = df[df['VALOR'].notna()]
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')
        df = df.drop_duplicates()
        df = df.sort_values('FECHA', na_position='last').reset_index(drop=True)
        df.index += 1
    resumen['TOTAL_ABONOS'] = df[df['VALOR'] > 0]['VALOR'].sum()
    resumen['TOTAL_CARGOS'] = df[df['VALOR'] < 0]['VALOR'].sum()
    resumen['SALDO_INICIAL'] = df.iloc[0]['SALDO'] if not df.empty else 0
    resumen['SALDO_FINAL']   = df.iloc[-1]['SALDO'] if not df.empty else 0
    return df, resumen

