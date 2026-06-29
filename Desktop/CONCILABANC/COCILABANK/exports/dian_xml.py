"""
Exportación DIAN — Medios Magnéticos (Resolución DIAN 000055 y ss.)
Genera archivos XML para reporte de terceros ante la DIAN colombiana.
"""
import hashlib
import logging
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List, Optional
import io

import pandas as pd

log = logging.getLogger(__name__)


def _ahora() -> str:
    return datetime.now().isoformat(timespec='seconds')


def _leer_config(clave: str, default: str = '') -> str:
    try:
        from storage.db import _init_db
        conn = _init_db()
        row = conn.execute(
            "SELECT valor FROM configuracion_empresa WHERE clave=?", (clave,)
        ).fetchone()
        conn.close()
        return row[0] if row else default
    except Exception:
        return default


# ── Generadores XML ───────────────────────────────────────────────────────────
def generar_xml_medios_magneticos(
    df_transacciones: pd.DataFrame,
    periodo: str,
    tipo_reporte: str = '1001',
    usuario: str = 'admin',
) -> tuple:
    """
    Genera XML de medios magnéticos DIAN.

    tipo_reporte:
        '1001' = Socios y accionistas
        '1007' = Retenciones en la fuente practicadas
        '1008' = Pagos o abonos en cuenta
        '1009' = Pagos al exterior
        '2276' = Información de IVA generado

    Returns: (xml_bytes, nombre_archivo, hash_sha256)
    """
    empresa = _leer_config('empresa_nombre', 'CREDIEXPRESS POPAYÁN SAS')
    nit     = _leer_config('empresa_nit', '900000000-0').replace('-', '').strip()
    ciudad  = _leer_config('empresa_ciudad', 'Popayán')
    year    = periodo[:4] if len(periodo) >= 4 else str(datetime.now().year)

    # Raíz XML
    root = ET.Element('InformacionExogena')
    root.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
    root.set('version', '1.0')

    # Encabezado
    enc = ET.SubElement(root, 'Encabezado')
    ET.SubElement(enc, 'TipoReporte').text   = tipo_reporte
    ET.SubElement(enc, 'AnioGravable').text  = year
    ET.SubElement(enc, 'NitInformante').text = nit
    ET.SubElement(enc, 'RazonSocial').text   = empresa
    ET.SubElement(enc, 'FechaGeneracion').text = _ahora()
    ET.SubElement(enc, 'MunicipioActividad').text = ciudad
    ET.SubElement(enc, 'TotalRegistros').text = str(len(df_transacciones))

    # Detalle
    detalle = ET.SubElement(root, 'Detalle')
    total_pagos = 0.0

    desc_col = next((c for c in ['DESCRIPCION', 'CONCEPTO'] if c in df_transacciones.columns), None)
    val_col  = next((c for c in ['VALOR', 'valor'] if c in df_transacciones.columns), None)
    fecha_col = next((c for c in ['FECHA_RAW', 'FECHA'] if c in df_transacciones.columns), None)

    for i, (_, row) in enumerate(df_transacciones.iterrows(), 1):
        reg = ET.SubElement(detalle, 'Registro')
        reg.set('secuencia', str(i))

        valor = float(row.get(val_col, 0) or 0) if val_col else 0.0
        total_pagos += abs(valor)

        ET.SubElement(reg, 'NumeroIdentificacion').text = '0000000000'  # NIT tercero
        ET.SubElement(reg, 'TipoDocumento').text        = '31'  # NIT
        ET.SubElement(reg, 'RazonSocial').text          = str(row.get(desc_col, ''))[:60] if desc_col else ''
        ET.SubElement(reg, 'ValorPagos').text           = f"{abs(valor):.2f}"
        ET.SubElement(reg, 'ValorRetencion').text       = '0.00'
        ET.SubElement(reg, 'FechaTransaccion').text     = str(row.get(fecha_col, ''))[:10] if fecha_col else ''
        ET.SubElement(reg, 'Concepto').text             = tipo_reporte

    # Totales
    tot = ET.SubElement(root, 'Totales')
    ET.SubElement(tot, 'TotalPagos').text     = f"{total_pagos:.2f}"
    ET.SubElement(tot, 'TotalRetencion').text = '0.00'
    ET.SubElement(tot, 'NumeroRegistros').text = str(len(df_transacciones))

    # Serializar
    tree = ET.ElementTree(root)
    buf  = io.BytesIO()
    tree.write(buf, encoding='utf-8', xml_declaration=True)
    xml_bytes   = buf.getvalue()
    hash_sha256 = hashlib.sha256(xml_bytes).hexdigest()

    nombre = f"DIAN_{tipo_reporte}_{nit}_{year}_{_ahora()[:10].replace('-','')}.xml"

    # Registrar en SQLite
    _registrar_exportacion_dian(periodo, tipo_reporte, nombre, hash_sha256, usuario)

    log.info("[dian_xml] Generado %s | %d registros | SHA-256: %s…",
             nombre, len(df_transacciones), hash_sha256[:16])
    return xml_bytes, nombre, hash_sha256


def generar_formato_1007_retenciones(df: pd.DataFrame, periodo: str,
                                      usuario: str = 'admin') -> tuple:
    """Formato 1007: Retenciones en la fuente practicadas."""
    # Filtrar solo movimientos con retención
    import re
    if df is None or df.empty:
        return b'', '', ''
    desc_col = next((c for c in ['DESCRIPCION', 'CONCEPTO'] if c in df.columns), None)
    if desc_col:
        mascara = df[desc_col].str.contains(
            r'RETENCI[ÓO]N|RETEFUENTE|RTEFTE', case=False, na=False, regex=True
        )
        df_ret = df[mascara].copy()
    else:
        df_ret = pd.DataFrame()
    return generar_xml_medios_magneticos(df_ret, periodo, '1007', usuario)


def generar_formato_1008_pagos(df: pd.DataFrame, periodo: str,
                                usuario: str = 'admin') -> tuple:
    """Formato 1008: Pagos o abonos en cuenta (débitos del período)."""
    if df is None or df.empty:
        return b'', '', ''
    tipo_col = 'TIPO' if 'TIPO' in df.columns else None
    if tipo_col:
        df_pagos = df[df[tipo_col] == 'DEBITO'].copy()
    else:
        df_pagos = df.copy()
    return generar_xml_medios_magneticos(df_pagos, periodo, '1008', usuario)


def _registrar_exportacion_dian(periodo: str, tipo: str, archivo: str,
                                  hash_sha256: str, usuario: str):
    try:
        from storage.db import _init_db
        conn = _init_db()
        conn.execute(
            """INSERT INTO exportaciones_dian
               (periodo, tipo, archivo_xml, hash_sha256, estado, usuario, fecha_generacion)
               VALUES (?,?,?,?,?,?,?)""",
            (periodo, f'DIAN_{tipo}', archivo, hash_sha256, 'GENERADO', usuario, _ahora())
        )
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[dian_xml] Error registrando: %s", e)


def listar_exportaciones_dian(limite: int = 50) -> list:
    """Lista exportaciones DIAN generadas."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        rows = conn.execute(
            "SELECT periodo, tipo, archivo_xml, hash_sha256, estado, "
            "usuario, fecha_generacion FROM exportaciones_dian "
            "ORDER BY id DESC LIMIT ?", (limite,)
        ).fetchall()
        conn.close()
        return [{'periodo': r[0], 'tipo': r[1], 'archivo': r[2],
                 'hash': r[3], 'estado': r[4], 'usuario': r[5], 'fecha': r[6]}
                for r in rows]
    except Exception as e:
        log.error("[dian_xml] Error listando: %s", e)
        return []
