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
    """
    Nivel 2 - Infiere (anio, mes, confianza_pct) analizando las FECHAS
    del contenido del DataFrame.
    Retorna (anio, mes, confianza) o None.
    """
    import pandas as pd

    candidatas = list(cols_fecha or [])
    candidatas += [
        'FECHA_BANCO', 'FECHA_RAW', 'FECHA', 'DATE', 'fecha',
        'Fecha', 'FECHA_CONTABLE', 'FechaTransaccion',
    ]

    for col in candidatas:
        if col not in df.columns:
            continue
        try:
            fechas = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dropna()
            if len(fechas) < 2:
                continue
            periodos = Counter((int(f.year), int(f.month)) for f in fechas)
            if not periodos:
                continue
            (anio, mes), n_mas = periodos.most_common(1)[0]
            confianza = round(n_mas / len(fechas) * 100, 1)
            if confianza >= 50:
                return (anio, mes, confianza)
        except Exception as exc:
            log.debug("[periodo] Error leyendo columna %s: %s", col, exc)
    return None


def validar_periodo_archivos(nombre_b, df_b, nombre_a, df_a):
    """
    Nivel 3 - Cruce Nivel 1 (nombre) + Nivel 2 (contenido) para ambos archivos.
    Retorna dict con claves:
      ok, periodo, mensaje, nivel, periodo_banco, periodo_aux, conf_banco, conf_aux
    """
    def _detectar(nombre, df):
        p_nom = extraer_periodo_banco(nombre)   # (anio, mes) o None
        p_con = inferir_periodo_desde_df(df)    # (anio, mes, conf) o None

        if p_nom and p_con:
            nom_anio, nom_mes = p_nom
            con_anio, con_mes, conf = p_con

            if nom_anio == con_anio and nom_mes == con_mes:
                # Acuerdo perfecto nombre + contenido
                return p_nom, 100.0, '1+2'
            elif nom_mes == con_mes and nom_anio != con_anio:
                # Mismo mes, año distinto → el NOMBRE gana.
                # Caso típico: el PDF bancario incluye la fecha del saldo anterior
                # del mes previo (ej. DESDE: 2025/12/31) y eso sesga la detección
                # de contenido hacia el año anterior. El nombre del archivo
                # (ENERO_2026) es la fuente de verdad del usuario.
                log.debug(
                    "[periodo] Nombre=%s vs Contenido=%s — mismo mes, año distinto. "
                    "Se usa año del nombre.", p_nom, (con_anio, con_mes)
                )
                return p_nom, max(conf, 75.0), '1(año)+2(mes)'
            elif nom_anio == con_anio and nom_mes != con_mes:
                # Mismo año, mes distinto → el contenido tiene más datos reales
                return (con_anio, con_mes), conf, '2(mes)'
            else:
                # Desacuerdo total → nombre del archivo gana (más confiable que
                # el contenido si el parser incluye fechas de encabezado/pie)
                log.debug(
                    "[periodo] Nombre=%s vs Contenido=%s — desacuerdo total. "
                    "Se usa nombre.", p_nom, (con_anio, con_mes)
                )
                return p_nom, 65.0, '1(nombre)'
        elif p_con:
            return (p_con[0], p_con[1]), p_con[2], '2'
        elif p_nom:
            return p_nom, 60.0, '1'
        else:
            return None, 0.0, 'desconocido'

    periodo_b, conf_b, nivel_b = _detectar(nombre_b, df_b)
    periodo_a, conf_a, nivel_a = _detectar(nombre_a, df_a)
    _nivel = f"banco={nivel_b} aux={nivel_a}"

    if periodo_b and periodo_a:
        if periodo_b == periodo_a:
            mes_lbl = _MESES_ES[periodo_b[1] - 1]
            return {
                'ok': True,
                'periodo': periodo_b,
                'mensaje': f"Periodo confirmado: {mes_lbl} {periodo_b[0]}",
                'nivel': _nivel,
                'periodo_banco': periodo_b,                 'conf_banco': conf_b, 'conf_aux': conf_a,
            }
        else:
            mb = _MESES_ES[periodo_b[1] - 1]
            ma = _MESES_ES[periodo_a[1] - 1]
            return {
                'ok': False,
                'periodo': None,
                'mensaje': (f"Periodos distintos - Banco: {mb} {periodo_b[0]} "
                            f"vs Auxiliar: {ma} {periodo_a[0]}"),
                'nivel': _nivel,
                'periodo_banco': periodo_b, 'periodo_aux': periodo_a,
                'conf_banco': conf_b, 'conf_aux': conf_a,
            }
    elif periodo_b or periodo_a:
        p_det = periodo_b or periodo_a
        mes_lbl = _MESES_ES[p_det[1] - 1]
        return {
            'ok': True,
            'periodo': p_det,
            'mensaje': f"Periodo detectado en un solo archivo: {mes_lbl} {p_det[0]}",
            'nivel': _nivel,
            'periodo_banco': periodo_b, 'periodo_aux': periodo_a,
            'conf_banco': conf_b, 'conf_aux': conf_a,
        }
    else:
        return {
            'ok': True,
            'periodo': None,
            'mensaje': "No se pudo detectar el periodo en ninguno de los archivos",
            'nivel': 'desconocido',
            'periodo_banco': None, 'periodo_aux': None,
            'conf_banco': 0.0, 'conf_aux': 0.0,
        }


# Alias compatible con codigo original
_extraer_periodo = extraer_periodo_banco
o = extraer_periodo_banco
