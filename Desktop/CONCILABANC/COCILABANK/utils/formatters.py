"""
utils/formatters.py — Formateadores de UI para CREDIEXPRESS
"""
import re
import logging
import numpy as np

log = logging.getLogger(__name__)

# ── Formato monetario COP ─────────────────────────────────────────────────────
def cop(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return '                 N/A'
    signo = '-' if v < 0 else ' '
    return f'{signo}$ {abs(v):>18,.2f}'

def _cop_limpio(v):
    """Versión limpia del cop() para HTML."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 'N/A'
    return f"${abs(v):,.0f}" + (" CR" if v < 0 else "")

def pct_bar(p, width=20):
    filled = int(p / 100 * width)
    return '[' + '█' * filled + '░' * (width - filled) + ']'

# ── Semáforo de conciliación ──────────────────────────────────────────────────
def semaforo_conciliacion(pct):
    if pct >= 90: return "🟢", "EXCELENTE", "verde"
    if pct >= 75: return "🟡", "BUENA",     "naranja"
    if pct >= 50: return "🟠", "REGULAR",   "naranja"
    return "🔴", "CRÍTICA", "rojo"

# Alias interno
_semaforo_conciliacion = semaforo_conciliacion

# ── Sugerencia contable ───────────────────────────────────────────────────────
def _inferir_cuenta_sugerida(desc, valor):
    """Sugiere cuenta contable para movimientos bancarios sin asiento."""
    d = (desc or '').upper()
    if any(x in d for x in ['GMF','4X1000','IMPTO GOBIERNO']):
        return '5305 — Impuesto GMF 4×1000', 'NC'
    if any(x in d for x in ['COMISION','COMISIÓN']):
        return '5305 — Comisiones Bancarias', 'NC'
    if 'NEQUI' in d:
        return '5305 — Comisiones Nequi/PSE', 'NC'
    if 'PSE' in d and valor < 0:
        return '5305 — Comisiones PSE', 'NC'
    if any(x in d for x in ['INTERES','INTERÉS','RENDIMIENTO']):
        return ('4205 — Rendimientos Financieros', 'CE') if valor > 0 else ('5305 — Intereses Débito', 'NC')
    if any(x in d for x in ['PAGO A PROV','PAGO A PROVE','PAGO PROVE']):
        return '2205 — Proveedores', 'CE'
    if any(x in d for x in ['NOMINA','NÓMINA','SALARIO']):
        return '2335 — Nómina por Pagar', 'CE'
    if any(x in d for x in ['TRANSFERENCIA','TRASLADO']):
        return '1110 — Bancos (verificar destino)', 'CE' if valor > 0 else 'NC'
    if valor > 0:
        return '1305 — Clientes / Recaudo (verificar)', 'CE'
    return '5999 — Otros Gastos (verificar)', 'NC'

def _guia_banco_sin_aux(row):
    """Genera instrucción específica para movimiento bancario sin asiento."""
    fecha  = row.get('FECHA_BANCO', '')
    desc   = str(row.get('DESCRIPCION', ''))[:60]
    valor  = row.get('VALOR_BANCO', 0)
    tipo   = row.get('TIPO_MOV', '')
    pagina = row.get('PAGINA_PDF', '')
    ref_pagina = f"Página <b>{pagina}</b> del PDF" if pagina else "Ver extracto bancario"
    cuenta, comprobante = _inferir_cuenta_sugerida(desc, valor)
    signo  = "+" if valor > 0 else "-"

    cand_doc  = str(row.get('CANDIDATO_DOC', '') or '')
    cand_conc = str(row.get('CANDIDATO_CONCEPTO', '') or '')[:60]
    cand_mon  = row.get('CANDIDATO_MONTO', None)
    diff_pct  = row.get('CANDIDATO_DIFF_PCT', None)
    razon     = str(row.get('RAZON_NO_MATCH', '') or '')

    if cand_doc and str(cand_mon) not in ('nan', '', 'None'):
        try:
            _dpct_str = f"{float(diff_pct):.1f}%" if diff_pct is not None else '—'
            _cmon_str = f"${float(cand_mon):,.0f}"
        except Exception:
            _dpct_str, _cmon_str = '—', '—'
        if float(diff_pct or 99) <= 0.5:
            _cand_color = '#f57f17'; _cand_label = '🟡 Monto coincide — revisar concepto/doc'
        elif float(diff_pct or 99) <= 5:
            _cand_color = '#e65100'; _cand_label = '🟠 Diferencia leve de monto'
        else:
            _cand_color = '#c62828'; _cand_label = '🔴 Sin asiento equivalente — crear NC'
        cand_html = f"""<br><b>🔍 CANDIDATO MÁS CERCANO EN AUXILIAR</b>
&nbsp;&nbsp;Doc: <b>{cand_doc}</b> &nbsp;|&nbsp; Monto: <b>{_cmon_str}</b>
&nbsp;&nbsp;Diff: <b style='color:{_cand_color}'>{_dpct_str}</b> &nbsp;|&nbsp;
<span style='color:{_cand_color}'>{_cand_label}</span><br>
&nbsp;&nbsp;Concepto: <i>"{cand_conc}"</i><br>
&nbsp;&nbsp;<small style='color:#666'>Razón: {razon}</small>"""
    else:
        cand_html = "<br><b>🔍 CANDIDATO:</b> <span style='color:#c62828'>Sin asiento similar encontrado — crear NC</span>"

    return f"""
<div class='guia-row'>
<b>📍 UBICAR EN EXTRACTO BANCARIO</b><br>
&nbsp;&nbsp;Fecha: <b>{fecha}</b> &nbsp;|&nbsp; Tipo: <b>{tipo}</b> &nbsp;|&nbsp; Valor: <b>{signo}${abs(valor):,.0f}</b> &nbsp;|&nbsp; {ref_pagina}<br>
&nbsp;&nbsp;Descripción: <i>"{desc}"</i>
{cand_html}<br>
<b>✏️ ACCIÓN EN SISTEMA CONTABLE</b><br>
&nbsp;&nbsp;Crear comprobante: <b>{comprobante}-XXXXXX</b><br>
&nbsp;&nbsp;Fecha: <b>{fecha}</b> &nbsp;|&nbsp; Cuenta sugerida: <b>{cuenta}</b><br>
&nbsp;&nbsp;Valor: <b>${abs(valor):,.0f}</b>
</div>"""

def _guia_aux_sin_banco(row):
    """Genera instrucción específica para asiento contable sin transacción bancaria."""
    doc     = str(row.get('DOCUMENTO', ''))
    fecha   = str(row.get('FECHA_RAW', ''))
    concepto= str(row.get('CONCEPTO', ''))[:60]
    deb     = row.get('DEBITO',  None)
    cre     = row.get('CREDITO', None)
    valor   = deb if deb else cre
    col     = row.get('COLUMNA', '')
    return f"""
<div class='guia-row'>
<b>📋 DOCUMENTO EN AUXILIAR CONTABLE</b><br>
&nbsp;&nbsp;Documento: <b>{doc}</b> &nbsp;|&nbsp; Fecha: <b>{fecha}</b> &nbsp;|&nbsp; Tipo: <b>{col}</b><br>
&nbsp;&nbsp;Concepto: <i>"{concepto}"</i> &nbsp;|&nbsp; Valor: <b>${abs(valor or 0):,.0f}</b><br><br>
<b>🔍 BUSCAR EN EXTRACTO BANCARIO</b><br>
&nbsp;&nbsp;Buscar movimiento de <b>${abs(valor or 0):,.0f} COP</b> cerca del <b>{fecha}</b><br>
&nbsp;&nbsp;Si no aparece: verificar si fue anulado, está en otro período o es asiento de ajuste interno.
</div>"""
