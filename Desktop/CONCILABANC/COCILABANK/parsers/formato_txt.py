"""
parsers/formato_txt.py — Parsers de auxiliar contable en formato TXT
"""
import re
import logging
from datetime import datetime

import pandas as pd

from engine.columna import determinar_columna
from storage.db import registrar_formato_pdf
from parsers.banco_pdf import es_fecha_banco

log = logging.getLogger(__name__)


def limpiar_num(t):
    from parsers.banco_pdf import limpiar_num as _ln
    return _ln(t)

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

def parsear_auxiliar_txt(texto_completo):
    meta = {}
    registros = []
    lineas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    pending_doc = None
    PAT_DOC = re.compile(r'^((?:CON|CE|CG|NC|RE|RG|FA|RI|REC|CXP|CXC|OC|NI|AJ|TF|EG|EI|NE|CT|ND|JR|CV|CM|TE|TC)-\d+)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(.*)')
    PAT_MONTO = re.compile(r'^([\d]{1,3}(?:,[\d]{3})*(?:\.[\d]{1,2})?)$')
    PAT_MPFX  = re.compile(r'^([\d]{1,3}(?:,[\d]{3})*(?:\.[\d]{1,2})?)\s+((?:CON|CE|CG|NC|RE|RG|FA|RI|REC|CXP|CXC|OC|NI|AJ|TF|EG|EI|NE|CT|ND|JR|CV|CM|TE|TC)-\d+.*)$')
    PAT_MSFX  = re.compile(r'\s([\d]{1,3}(?:,[\d]{3})*(?:\.[\d]{1,2})?)$')
    def guardar(doc, fecha_s, concepto, monto_str):
        monto = limpiar_num(monto_str.replace(',',''))
        if not monto or monto <= 0: return
        col = determinar_columna(concepto, doc)
        debito = monto if col=='DEBITO' else None
        credito = monto if col=='CREDITO' else None
        try: fdt = pd.to_datetime(fecha_s, format='%d/%m/%Y', errors='coerce')
        except: fdt = pd.NaT
        registros.append({
            'DOCUMENTO': doc, 'FECHA_RAW': fecha_s, 'FECHA': fdt,
            'CONCEPTO': concepto, 'DEBITO': debito, 'CREDITO': credito,
            'COLUMNA': col, 'VALOR_NETO': (debito or 0) - (credito or 0)
        })
    for linea in lineas:
        m_pfx = PAT_MPFX.match(linea)
        if m_pfx:
            if pending_doc:
                guardar(pending_doc['doc'], pending_doc['date'], pending_doc['concept'], m_pfx.group(1))
                pending_doc = None
            m_doc = PAT_DOC.match(m_pfx.group(2))
            if m_doc:
                doc_c, fecha_s, concepto_raw = m_doc.group(1), m_doc.group(2), m_doc.group(3)
                m_end = PAT_MSFX.search(concepto_raw)
                if m_end:
                    guardar(doc_c, fecha_s, concepto_raw[:m_end.start()].strip(), m_end.group(1))
                else:
                    pending_doc = {'doc': doc_c, 'date': fecha_s, 'concept': concepto_raw}
            continue
        m_doc = PAT_DOC.match(linea)
        if m_doc:
            doc_c, fecha_s, concepto_raw = m_doc.group(1), m_doc.group(2), m_doc.group(3)
            m_end = PAT_MSFX.search(concepto_raw)
            if m_end:
                guardar(doc_c, fecha_s, concepto_raw[:m_end.start()].strip(), m_end.group(1))
            else:
                pending_doc = {'doc': doc_c, 'date': fecha_s, 'concept': concepto_raw}
            continue
        if PAT_MONTO.match(linea) and pending_doc:
            guardar(pending_doc['doc'], pending_doc['date'], pending_doc['concept'], linea)
            pending_doc = None
            continue
        if pending_doc:
            m_end = PAT_MSFX.search(linea)
            if m_end:
                try:
                    if float(m_end.group(1).replace(',','')) > 100:
                        guardar(pending_doc['doc'], pending_doc['date'], pending_doc['concept'], m_end.group(1))
                        pending_doc = None
                except: pass
    df = pd.DataFrame(registros)
    df = df.drop_duplicates()
    df = df.sort_values('FECHA', na_position='last').reset_index(drop=True)
    df.index += 1
    meta['TOTAL_DEBITOS'] = df['DEBITO'].sum() if not df.empty else 0
    meta['TOTAL_CREDITOS'] = df['CREDITO'].sum() if not df.empty else 0
    try:
        registrar_formato_pdf('', 'AUXILIAR', list(df.columns), 'txt', [],
                              banco_detectado='auxiliar_txt')
    except Exception:
        pass
    # Saldos no disponibles en TXT generalment