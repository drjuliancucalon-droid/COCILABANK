"""
utils/periodo.py - Extraccion de periodo contable del nombre de archivo y contenido
"""
import re
import logging
from collections import Counter

log = logging.getLogger(__name__)

_MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
             'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

_MESES_MAP = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12,
    'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12,
    'ene':1,'feb':2,'mar':3,'abr':4,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12,
    'jan':1,'aug':8,'dec':12,
}

# Patron numerico: 2025-02, 202502, 2025_02
_PAT_ANIO_MES  = re.compile(r'(20\d{2})[-_/]?(0[1-9]|1[0-2])(?:[^0-9]|$)')
_PAT_MES_ANIO  = re.compile(r'(?:^|[^0-9])(0[1-9]|1[0-2])[-_/]?(20\d{2})')


def extraer_periodo_banco(nombre_archivo):
    """
    Nivel 1 - Extrae (anio, mes) del NOMBRE del archivo.
    Soporta: 2025-02, 202502, febrero_2025, extracto_feb2025, etc.
    Devuelve (anio, mes) o None.
    """
    n = (nombre_archivo or '').lower()

    m = _PAT_ANIO_MES.search(n)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m = _PAT_MES_ANIO.search(n)
    if m:
        return (int(m.group(2)), int(m.group(1)))

    anio_m = re.search(r'(20\d{2})', n)
    mes_num = next(
        (v for k, v in _MESES_MAP.items() if re.search(r'(?<![a-z])' + k + r'(?![a-z])', n)),
        None,
    )
    if anio_m and mes_num:
        return (int(anio_m.group(1)), mes_num)
    return None


def inferir_periodo_desde_df(df, cols_fecha=None):
    import pandas as pd
    candidatas = list(cols_fecha or [])
    candidatas += ['FECHA_BANCO','FECHA_RAW','FECHA','DATE','fecha','Fecha','FECHA_CONTABLE','FechaTransaccion']
    for col in candidatas:
        if col not in df.columns: continue
        try:
            fechas = pd.to_datetime(df[col],dayfirst=True,errors='coerce').dropna()
            if len(fechas) < 2: continue
            periodos = Counter((int(f.year),int(f.month)) for f in fechas)
            if not periodos: continue
            (anio,mes),n_mas = periodos.most_common(1)[0]
            conf = round(n_mas/len(fechas)*100,1)
            if conf >= 50: return (anio,mes,conf)
        except Exception as exc: log.debug("[periodo] %s:%s",col,exc)
    return None

def validar_periodo_archivos(nombre_b,wf_b,nombre_a,df_a):
    def _detectar(nombre,df):
        p_nom=extraer_periodo_banco(nombre); p_con=inferir_periodo_desde_df(df)
        if p_nom and p_con:
            na_,nm_=p_nom; ca,cm,conf=p_con
            if na_ == ca and nm_ == cm: return p_nom,100.0,'1+2'
            elif nm_ == cm and na_ != ca: return p_nom,max(conf,75.0),'1(ano)+2(mes)'
            elif na_ == ca and nm_ != cm: return (ca,cm),conf,'2(mes)'
            else: return p_nom,65.0,'1(nombre)'
        elif p_con: return (p_con[0],p_con[1]),p_con[2],'2'
        elif p_nom: return p_nom,60.0,'1'
        else: return None,0.0,'desconocido'
    pb,x,yb=_detectar(nombre_b,wf_b); pa,y,ya=_detectar(nombre_a,df_a)
    _nivel=f"banco={yb} aux={ya}"
    if pb and pa:
        if pb == pa: return {'ok':True,'periodo':pb,'mensaje':f"Periodo confirmado: {_MESES_ES[pb[1]-1]} {pb[0]}",'nivel':_nivel,'periodo_banco':pb,'conf_banco':x,'conf_aux':y}
        else: return {'ok':False,'periodo':None,'mensaje':f"Periodos distintos",'nivel':_nivel,'periodo_banco':pb,'periodo_aux':pa,'conf_banco':x,'conf_aux':Y}
    elif pb or pa:
        pd= pb or pa; return {'ok':True,'periodo':pd,'mensaje':f"Periodo detectado: {_MESES_ES[pd[1]-1]} {pd[0]}",'nivel':_nivel,'periodo_banco':pb,'periodo_aux':pa,'conf_banco':x,'conf_aux':Y}
    else: return {'ok':True,'periodo':None,'mensaje':"No se pudo detectar el periodo",'nivel':'desconocido','periodo_banco':None,'periodo_aux':None,'conf_banco':0.0,'conf_aux':0.0}
_extraer_periodo=extraer_periodo_banco
