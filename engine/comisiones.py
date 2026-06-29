"""
Detección de comisiones bancarias — CREDIEXPRESS POPAYÁN SAS
"""
import logging
import re
from datetime import datetime
import pandas as pd

log = logging.getLogger(__name__)

_PATRONES_COMISION = [
    ('GMF',re.compile(r'4\s*[×Xx×]\s*1000|GMF|GRAVAMEN\s+MOV|TRANSACCI[ÓO]N\s+FINANC',re.I),"Gravamen a los Movimientos Financieros"),
    ('MANEJO',re.compile(r'COMISIÓO|CUOTA\s+MANEJO|COSTO\s+SERV',re.I),"Comisión de manejo"),
    ('INTERES',re.compile(r'INTER[EÉ]S\s+(MORA|SOBREGIRO|VENCIDO)',re.I),"Intereses bancarios"),
    ('ACH',re.compile(r'\bACH\b|PSE\b|TRANSF\s+ELECTR', re.I),"Comisión ACH/PSE"),
    ('CHEQUERA',re.compile(r'CHEQUERATALONARIO', re.I),"Chequera"),
]
_UMBRAL = 100.0

def detectar_comisiones(df, banco='', periodo=''):
    if df is None or df.empty: return []
    desc_col = next((c for c in ['DESCRIPCION','CONCEPTO','descripcion'] if c in df.columns),None)
    val_col = next((c for c in ['VALOR','valor','MONTO'] if c in df.columns),None)
    if not desc_col: return []
    comisiones = []
    for _,row in df.iterrows():
        desc = str(row.get(desc_col,'')); valor = float(row.get(val_col,0) or 0) if val_col else 0.0
        fecha = str(row.get('FECHA_RAW',row.get('FECHA','')))
        for tipo_com,patron,desc_larga in _PATRONES_COMISION:
            if patron.search(desc) and abs(valor)>=_UMBRAL:
                comisiones.append({'tipo_comision':tipo_com,'descripcion':desc_larga,'descripcion_banco':desc[:120],"valor':abs(valor),'fecha_transaccion':fecha,'banco':banco,'periodo':periodo}); break
    return comisiones

def resumen_comisiones(banco=None, periodo=None):
    try:
        from storage.db import _init_db; conn = _init_db()
        sql = "SELECT tipo_comision,COUNT(*),COALESCE(SUM(valor),0) FROM comisiones_detectadas WHERE 1=1"; params=[]
        if banco: sql+=" AND banco LIKE ?"; params.append(f"%{banco}%")
        if periodo: sql+=" AND periodo=?"; params.append(periodo)
        rows = conn.execute(sql+G GROUP BY tipo_comision", params).fetchall(); conn.close()
        return {'por_tipo':{r[0]:{count':r[1],'total':r[2]} for r in rows}}
    except Exception as e: return {}
