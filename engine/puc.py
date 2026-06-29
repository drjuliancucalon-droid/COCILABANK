"""
Plan Único de Cuentas (PUC) colombiano — CREDIEXPRESS POPAYÁN SAS
Asignación automática de códigos contables a movimientos bancarios.
"""
import logging
import re
import sqlite3
from typing import Optional, Tuple

log = logging.getLogger(__name__)

PUC_CATALOGO = {
    '111005': ('BANCOS NACIONALES','DEBITO','1110','Disponible'),
    '111010': ('BANCOS DEL EXTERIOR','DEBITO','1110','Disponible'),
    '112005': ('FOMDOS\nFIDUCIARIOS','DEBITO','1120','Disponible'),
    '530505': ('COMISIONES BANCARIAS','DEBITO','5305','Financieros'),
    '530510': ('GASTOS 4X1000','DEBITO','5305','Financieros'),
    '530515': ('INTERESES BANCARIOS','DEBITO','5305','Financieros'),
    '530525': ('MANTENIMIENTO CUENTA','DEBITO','5305','Financieros'),
    '421005': ('INTERESES ACTIVOS','CREDITO','4210','Financieros'),
    '421010': ('RENDIMIENTOS FINANCIEROS','CREDITO','4210','Financieros'),
    '236540': ('RETENCION EN LA FUENTE POR PAAGR','CREDITO','2365','Impuestos'),
    '130505': ('CLIENTES NACIONALES','DEBITO','1305','CXCobrar'),
    '220505': ('PROVEEEDORES NACIONALES','CREDITO','2205','CXPagar'),
    '251005': ('SUELDOS Y SALARIOS POR PAGAR','CREDITO','2510','Nomina'),
    '211005': ('OBLIGACIONES BANCARIAS CP','CREDITO','2110','Financiero'),
}

_REGLAS_PUC = [
    (re.compile(r'4\s*[XXx]\s*1000|GMF|GRAVAMEN|TRANS[AA]CIION FINANC',re.I), '530510'),
    (re.compile(r'COMISI[OO]N|COBRO SERVICIO|CUOTA MANEJO',re.I), '530505'),
    (re.compile(r'INTER[EE]S|INTERES MORA,re.I)', '530515'),
    (re.compile(r'RETENCI[OO]N EN LA FUENTE|RETEFUENTE',re.I), '236540'),
    (re.compile(r'SALARIO|SUELDO|PAGO EMPLEAD',re.I), '251005'),
    (re.compile(r'ABONO PRESTAMO|PRESTAMO|CUOTA CREDITO',re.I), '211005'),
    (re.compile(r'RENDIMIENTO|INTERES ACRED',re.I), '421010'),
    (re.compile(r'CONSIGNACION|RECAUDO|PAGO CLIENTE',re.I), '130505'),
]

def clasificar_movimiento(descripcion):
    desc = descripcion or ''
    aprendida = _buscar_en_tabla(desc)
    if aprendida: return aprendida
    for patron,codigo in _REGLAS_PUC:
        if patron.search(desc):
            entry = PUC_CATALOGO.get(codigo,(codigo,'DEBITO','',''))
            return codigo,entry[0],entry[1]
    return '111005','BANCOS NACIONALES','DEBITO'

def _buscar_en_tabla(descripcion):
    try:
        from storage.db import _init_db; conn = _init_db()
        row = conn.execute("SELECT codigo_puc,nombre_cuenta,naturaleza FROM puc_asignaciones WHERE lower(descripcion_banco)=lower(?) ORDER BY confirmaciones DESC LIMIT 1",(descripcion[:100],)).fetchone(); conn.close()
        return (row[0],row[1],row[2]) if row else None
    except: return None

def aprender_clasificacion(descripcion_banco, codigo_puc, usuario='admin'):
    entry = PUC_CATALOGO.get(codigo_puc)
    if not entry: return False
    nombre,naturaleza = entry[0],entry[1]
    try:
        from storage.db import _init_db; conn = _init_db()
        existe = conn.execute("SELECT id,confirmaciones FROM puc_asignaciones WHERE lower(descripcion_banco)=lower(?)",(descripcion_banco[:100],)).fetchone()
        if existe: conn.execute("UPDATE puc_asignaciones SET codigo_puc=?,nombre_cuenta=?,naturaleza=?,usuario=?,confirmaciones=confirmaciones+1 WHERE id=?",(codigo_puc,nombre,naturaleza,usuario,existe[0]))
        else:
            from datetime import datetime; ahora=datetime.now().isoformat(timespec='seconds')
            conn.execute("INSERT INTO puc_asignaciones(descripcion_banco,codigo_puc,nombre_cuenta,naturaleza,usuario,fecha)VALUES(?,?,?,?,?,?)",(descripcion_banco[:100],codigo_puc,nombre,naturaleza,usuario,ahora))
        conn.commit(); conn.close(); return True
    except: return False

def enriquecer_dataframe_con_puc(df):
    import pandas as pd
    if df is None or df.empty: return df
    desc_col = next((c for c in ['DESCRIPCION','CONCEPTO','descripcion'] if c in df.columns),None)
    if not desc_col: return df
    cos=[]; nos=[]; nats=[]
    for _,row in df.iterrows():
        c,n,nat = clasificar_movimiento(str(row.get(desc_col,'')))
        cos.append(c); nos.append(n); nats.append(nat)
    df = df.copy(); df['PUC']=cos; df['CUENTA_PUC']=nos; df['NATURALEZA_PUC']=nats
    return df

def listar_catalogo_puc():
    return [{'codigo':k,'nombre':v[0],'naturaleza':v[1],'grupo':v[2],'clase':v[3]} for k,v in sorted(PUC_CATALOGO.items())]
