"""
parsers/despachador.py
Despachador/router de parsers para COCILABANK.
Detecta el tipo de archivo (banco/auxiliar) y formato (PDF/CSV/XLS/TXT)
y llama al parser correcto.
"""
from __future__ import annotations
import logging, os
from typing import Optional
import pandas as pd

logger = logging.getLogger(__name__)
_EXT_PDF = {'.pdf'}
_EXT_CSV = {'.csv'}
_EXT_XLS = {'.xls', '.xlsx', '.xlsm', '.xlsb'}
_EXT_TXT = {'.txt', '.prn', '.dat', '.asc'}


def _ext(nombre):
    return os.path.splitext(nombre)[1].lower()


def _detectar_tipo(nombre, contenido, hint=None):
    if hint:
        return hint.lower()
    n = nombre.upper()
    banco = ['BANCO', 'EXTRACTO', 'BANCOLOMBIA', 'DAVIVIENDA', 'BBVA', 'BOGOTA', 'NEQUI', 'NUBANK', 'AVVILLAS']
    aux = ['AUXILIAR', 'CONTAB', 'SIIGO', 'HELISA', 'WORLD', 'CONTAPYME', 'ALEGRA', 'COMPROBANTE']
    for p in banco:
        if p in n:
            return 'banco'
    for p in aux:
        if p in n:
            return 'auxiliar'
    if contenido:
        pr = contenido[:500].decode('utf-8', errors='ignore').upper()
        for p in banco:
            if p in pr:
                return 'banco'
        for p in aux:
            if p in pr:
                return 'auxiliar'
    return 'banco'


def despachar(file_bytes, nombre_archivo, tipo_hint=None, banco_hint=None,
              encoding='utf-8', ano_defecto=None):
    if not file_bytes:
        raise ValueError(f"Archivo vacio: {nombre_archivo}")
    ext = _ext(nombre_archivo)
    tipo = _detectar_tipo(nombre_archivo, file_bytes, tipo_hint)
    logger.info("Despachando '%s' ext=%s tipo=%s", nombre_archivo, ext, tipo)
    if ext in _EXT_PDF:
        if tipo == 'banco':
            from parsers.banco_pdf import parsear_banco_pdf
            return parsear_banco_pdf(file_bytes, ano_defecto=ano_defecto, banco_hint=banco_hint)
        else:
            from parsers.auxiliar_pdf import parsear_auxiliar_pdf
            return parsear_auxiliar_pdf(file_bytes, ano_defecto=ano_defecto)
    elif ext in _EXT_CSV:
        from parsers.formatos_csv import parsear_csv
        return parsear_csv(file_bytes, tipo=tipo, encoding=encoding)
    elif ext in _EXT_XLS:
        from parsers.formatos_csv import parsear_excel
        return parsear_excel(file_bytes, tipo=tipo)
    elif ext in _EXT_TXT:
        from parsers.formato_txt import parsear_txt
        return parsear_txt(file_bytes, tipo=tipo, encoding=encoding)
    else:
        try:
            from parsers.formato_txt import parsear_txt
            return parsear_txt(file_bytes, tipo=tipo, encoding=encoding)
        except Exception as e:
            raise ValueError(f"Formato no soportado: {ext}: {e}")


def despachar_multiples(archivos, tipo_hint=None, banco_hint=None,
                        encoding='utf-8', ano_defecto=None):
    dfs = []
    errores = []
    for fb, nm in archivos:
        try:
            df = despachar(fb, nm, tipo_hint, banco_hint, encoding, ano_defecto)
            if not df.empty:
                df['_archivo'] = nm
                dfs.append(df)
        except Exception as e:
            errores.append(f"{nm}: {e}")
            logger.error("%s: %s", nm, e)
    if not dfs:
        return pd.DataFrame(columns=['FECHA', 'DESCRIPCION', 'VALOR'])
    r = pd.concat(dfs, ignore_index=True)
    if 'FECHA' in r.columns:
        r['FECHA'] = pd.to_datetime(r['FECHA'], errors='coerce')
        r = r.sort_values('FECHA').reset_index(drop=True)
    return r


def formatos_soportados():
    return {
        'banco': {'pdf': 'Extractos PDF', 'csv': 'CSV/Excel', 'txt': 'TXT/PRN'},
        'auxiliar': {'pdf': 'Auxiliares PDF', 'csv': 'CSV/Excel (SIIGO, Helisa)'}
    }
