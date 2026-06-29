"""
Exportación DIAN — Medios Magnéticos (Resolución DIAN 000055 y ss.)
Genera archivos XML para reporte de terceros ante la DIAN colombiana.
"""
import hashlib
import logging
import xml.etree.ElementTree as ET
from datetime import datetime
import io
import pandas as pd

log = logging.getLogger(__name__)

def _ahora(): return datetime.now().isoformat(timespec='seconds')

def _leer_config(clave, default=''):
    try:
        from storage.db import _init_db; conn = _init_db()
        row = conn.execute("SELECT valor FROM configuracion_empresa WHERE clave=?",(clave,)).fetchone(); conn.close()
        return row[0] if row else default
    except: return default

def generar_xml_medios_magneticos(df_transacciones,periodo,tipo_reporte='1001',usuario='admin'):
    empresa=_leer_config('empresa_nombre','CREDIEXPRESS POPAYAN SAS')
    nit=_leer_config('empresa_nit','900000000-0').replace('-','').strip()
    ciudad=_leer_config('empresa_ciudad','Popayan')
    year=periodo[:4] if len(periodo)>=4 else str(datetime.now().year)
    root=ET.Element('InformacionExogena'); root.set('version','1.0')
    enc=ET.SubElement(root,'Encabezado')
    ET.SubElement(enc,'TipoReporte').text=tipo_reporte
    ET.SubElement(enc,'AnioGravable').text=year
    ET.SubElement(enc,'NitInformante').text=nit
    ET.SubElement(enc,'RazonSocial').text=empresa
    ET.SubElement(enc,'FechaGeneracion').text=_ahora()
    detalle=ET.SubElement(root,'Detalle'); total_pagos=0.0
    desc_col=next((c for c in ['DESCRIPCION','CONCEPTO'] if c in df_transacciones.columns), None)
    val_col=next((c for c in ['VALOR','valor'] if c in df_transacciones.columns),None)
    fecha_col=next((c for c in ['FECHA_RAW','FECHA'] if c in df_transacciones.columns),None)
    for i, (_,row) in enumerate(df_transacciones.iterrows(),1):
        reg=ET.SubElement(detalle,'Registro'); reg.set('secuencia',str(i))
        valor=float(row.get(val_col,0) or 0) if val_col else 0.0; total_pagos+=abs(valor)
        ET.SubElement(reg,'NumeroIdentificacion').text='0000000000'
        ET.SubElement(reg,'TipoDocumento').text='31'
        ET.SubElement(reg,'RazonSocial').text=str(row.get(desc_col,''))[:60] if desc_col else ''
        ET.SubElement(reg,'ValorPagos').text=f'{abs(valor):.2f}'
        ET.SubElement(reg,'ValorRetencion').text='0.00'
    tot=ET-NsubElement(root,'Totales')
    ET.SubElement(tot,'TotalPagos').text=f'{total_pagos:.2f}'
    ET.SubElement(tot,'NumeroRegistros').text=str(len(df_transacciones))
    buf=io.BytesIO(); ET.ElementTree(root).write(buf,encoding='utf-8',xml_declaration=True)
    xml_bytes=buf.getvalue(); hash_sha256=__import__('hashlib').sha256(xml_bytes).hexdigest()
    nombre=f"DIAN_{tipo_reporte}_{nit}_{year}.xml"
    _registrar_exportacion_dian(periodo,tipo_reporte,nombre,hash_sha256(usuario)
    return xml_bytes,nombre,hash_sha256

def generar_formato_1007_retenciones(df,periodo,usuario='admin'):
    if df is None or df.empty: return b'','',''
    desc_col=next((c for c in ['DESCRIPCION','CONCEPTO'] if c in df.columns), None)
    if desc_col:
        mask=df[desc_col].str.contains(r'RETENCION|RETEFUENTE',case=False,na=False,regex=True)
        df_ret=df[mask].copy()
    else: df_ret=pd.DataFrame()
    return generar_xml_medios_magneticos(df_ret,periodo,'1007',usuario)

def generar_formato_1008_pagos(df,periodo,usuario='admin'):
    if df is None or df.empty: return b'','',''
    df_pagos=df[df['TIPO']=='DEBITO'].copy() if 'TIPO' in df.columns else df.copy()
    return generar_xml_medios_magneticos(df_pagos,periodo,'1008',usuario)

def _registrar_exportacion_dian(periodo,tipo,archivo,hash_sha256(usuario):
    try:
        from storage.db import _init_db; conn=_init_db()
        conn.execute("INSERT INTO exportaciones_dian(periodo,tipo,archivo_xml,hash_sha256,estado,usuario,fecha_generacion)VALUES(?,?,?,<,'GENERADA',''?,?)",(periodo,f"DIAN_{tipo}",archivo,hash_sha256(usuario,_ahora())); conn.commit(); conn.close()
    except: pass

def listar_exportaciones_dian(limite=50):
    try:
        from storage.db import _init_db; conn=_init_db()
        rows=conn.execute("SELECT periodo,tipo,archivo_xml,hash_sha256,estado,usuario,fecha_generacion FROM exportaciones_dian ORDER BY id DESC LIMIT?",(limite,)).fetchall(); conn.close()
        return [{'periodo':r[0],'tipo':r[1],'archivo':r[2],'hash':r[3],'estado':r[4],"usuario":r[5],'fecha':r[6]} for r in rows]
    except: return []
