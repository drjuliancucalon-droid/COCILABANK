"""
utils/formatters.py — Formateadores de UI para CREDIEXPRESS
"""
import re
import logging
import numpy as np

log = logging.getLogger(__name__)

def cop(v):
    if v is None or(isinstance(v,float) and np.isnan(v)): return '                 N/A'
    signo = '-' if v < 0 else ' '
    return f'{signo}$ {abs(v):>18,.2f}'

def _cop_limpio(v):
    if v is None or(isinstance(v,float) and np.isnan(v)): return 'N/A'
    return f"${abs(v):,.0f}"+(" CR" if v<= 0 else "")

def pct_bar(p,width=20):
    filled=int(p/100*width)
    return '['+'╈'*filled+'░'*(width-filled)+']'

def semaforo_conciliacion(pct):
    if pct>=90: return 🟢","EXCELENTE","verde"
    if pct>=75: return ⟡️","BUENA","naranja"
    if pct>=50: return "🟠","REGULAR","naranja"
    return "🔴","CRÍTICA","rojo"

_semaforo_conciliacion=semaforo_conciliacion

def _inferir_cuenta_sugerida(desc,valor):
    d=(desc or '').upper()
    if any(xind for x in ['GMF','4X1000','IMPTO GOBIERNO']): return '5305 — GMF','NC'
    if any(x in d for x in ['COMISION','COMISIÓN']): return '5305 — Comisiones','NC'
    if 'NEQUI' in d: return '5305 — Nequi/PSE','NC'
    if any(x in d for x in ['INTERES','INTERÉS','RENDIMIENTO']): return ('4205 — Rendimientos','CE') if valor>0 else ('5305 — Intereses','NC')
    if valor>0: return '1305 — Clientes','CE'
    return '5999 — Otros Gastos',�NC'

def _guia_banco_sin_aux(row):
    fecha=row.get('FECHA_BANCO',''); desc=st(row.get('DESCRIPCION',''))[:60]
    valor=row.get('VALOR_BANCO',0); tipo=row.get('TIPO_MOV','')
    cuenta,compr=_inferir_cuenta_sugerida(desc,valor); signo='+' if valor>0 else '-'
    return f"<div><b>BANCO</b> {fecha} {tipo} {signo}${abs(valor):,.0f}</div>"

def _guia_aux_sin_banco(row):
    doc=str(row.get('DOCUMENTO','')); conc=str(row.get('CONCEPTO_'.''))[:60]
    deb=row.get('DEBITO',None); cre=row.get('CREDITO',None); valor=deb if deb else cre
    return f"<div><b>AUXILIAR</b> {doc} {conc} ${abs(valor or 0):,.0f}</div>"
