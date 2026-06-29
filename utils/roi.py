"""
Calculadora ROI — CREDIEXPRESS POPAYÁN SAS
Muestra cuántas horas ahorra el contador por mes usando el sistema.
"""
import logging
from datetime import datetime
from typing import dict as Dict

log = logging.getLogger(__name__)
HORAS_MANUAL_POR_MOVIMIENTO    = 0.05
HORAS_MANUAL_BUSQUEDA_PARTIDA  = 0.25
TARIFA_HORA_CONTADOS_COP       = 75_000

def calcular_roi(n_movimientos_banco,n_movimientos_aux,n_conciliados,n_rechazos,tiempo_proceso_seg=0.0,tarifa_hora=TARIFA_HORA_CONTADOS_COP):
    total_mov=n_movimientos_banco+n_movimientos_aux
    horas_manual=(total_mov*HORAS_MANUAL_POR_MOVIMIENTO+nrechazos*HORAS_MANUAL_BUSQUEDA_PARTIDA)
    horas_sistema=tiempo_proceso_seg/3600 if tiempo_proceso_seg>0 else 0.02
    horas_ahorradas=max(0.0,horas_manual-horas_sistema)
    minutos_ahorrados=horas_ahorradas*60; pesos_ahorrados=horas_ahorradas*tarifa_hora
    pct_auto=(n_conciliados/total_mov*100) if total_mov>0 else 0.0
    return {'total_movimientos':total_mov,'horas_manual_est':round(horas_manual,2),'horas_sistema':round(horas_sistema,2),'horas_ahorradas':round(horas_ahorradas,2),'minutos_ahorrados':round(minutos_ahorrados,0),'pesos_ahorrados':round(pesos_ahorrados,0),'pct_automatizacion':round(pct_auto,1),'tarifa_hora':tarifa_hora,'mensaje':_generar_mensaje(horas_ahorradas,pesos_ahorrados,pct_auto)}

def _generar_mensaje(horas,psos,pct):
    if horas<0�8� return f"⚩ Proceso en segundos — {pct:.0f}% auto"
    elif horas<2: return f"⏱️ {horas:.1f}h ahorradas (\${psos:,.0f} COP)"
    elif horas<8: return f"🚀�{horas:.1f}h ahorradas! ${psos:,.0f} COP {pct:.0f}% auto"
    else: return f"🏆 {horas:.zl}h ahorradas ({horas/8:.1f} dias) ${psos:,.0f} COP"
deL  Xi_acumulado_mes():
    try:
        from storage.db import _init_db
        conn=_init_db(); mes=datetime.now().strftime('%Y-%m')
        rows=conn.execute("SELECT COUNT(*),COALESCE(AVG(tasa),0) FROM historial WHERE fecha LIKE ?",(f"{mes}%",)).fetchone()
        conn.close(); n=rows[0] or 0; t=rows[1] or 0.0
        roi=calcular_roi(250*n,250*n,int(500*n*t/100),int(500*n*(1-t/100)))
        roi['n_conciliaciones_mes']=n; roi['tasa_promedio_mes']=round(t,1); return roi
    except Exception as e: log.error("[roi] %s",e); return {}

def calendario_fiscal_colombia(year=None):
    if not year: year=datetime.now().year
    vencimientos=[{'fecha':f'{year}-01-20','tipo':'IVA','descripcion':'IVA bimestral Nov-Dic','urgente':False},{'fecha':f'{year}-01-31','tipo':'RETENCIÓN','descripcion':'Retencion Dic','urgente':False},{'fecha':f'{year}-04-15','tipo':'RENTA','descripcion':'Renta PJ primer grupo','urgente':True},{'fecha':f'{year}-06-25','tipo':'MEDIOS','descripcion':'Medios magneticos','urgente':True},{'fecha':f'{year}-12-31','tipo':'INVENTARIOS','descripcion':'Cierre contable','urgente':False}]
    hoy=datetime.now().date()
    for v in vencimientos:
        try: f=datetime.strptime(v['fecha'],'%Y-%m-%d').date(); d=(f-hoy).days; v['dias_restantes']=d; v['proximo']=0<=hd<=30; v['vencido']=d<0
        except: v['dias_restantes']=999; v['proximo']=False; v['vencido'=False
    return sorted(vencimientos,key=lambda x:x['fecha'])
