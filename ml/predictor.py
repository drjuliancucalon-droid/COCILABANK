"""
ML Predictor — CREDIEXPRESS POPAYÁN SAS
Predice partidas conciliatorias del próximo mes usando historial SQLite.
Modelo: frecuencia + Jaccard sobre descripciones (sin dependencias ML externas).
"""
import hashlib
import logging
import re
from collections import Counter
from datetime import datetime
from typing import List

log = logging.getLogger(__name__)

_STOP = {'de','la','el','los','las','en','con','por','para','del','al',
         'un','una','y','a','se','que','no','es','su','le','lo'}


def _tokenizar(texto: str) -> frozenset:
    tokens = re.findall(r'[A-Za-záéíóúÁÉÍÓÚñÑ0-9]+', texto.lower())
    return frozenset(t for t in tokens if len(t) >= 3 and t not in _STOP)


def _jaccard(a: frozenset, b: frozenset) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _get_db():
    from storage.db import _init_db
    return _init_db()


def _ahora() -> str:
    return datetime.now().isoformat(timespec='seconds')


def predecir_partidas_proximas(periodo_base: str,
                                periodo_predicho: str,
                                min_confianza: float = 0.4) -> List[dict]:
    """
    Analiza el historial de partidas conciliatorias y predice cuáles
    es probable que aparezcan en el siguiente período.

    Lógica:
    1. Toma las partidas que se repitieron en >= 2 períodos anteriores.
    2. Calcula similitud de descripción entre períodos (Jaccard).
    3. Retorna predicciones con confianza estimada.
    """
    try:
        conn = _get_db()
        # Historial de partidas conciliadas y pendientes
        rows = conn.execute(
            """SELECT tipo, descripcion, valor, periodo_origen, estado
               FROM partidas_conciliatorias
               WHERE periodo_origen != ? AND fecha_registro IS NOT NULL
               ORDER BY fecha_registro DESC LIMIT 500""",
            (periodo_predicho,)
        ).fetchall()
        conn.close()
    except Exception as e:
        log.error("[ml] Error leyendo historial: %s", e)
        return []

    if not rows:
        log.info("[ml] Sin historial suficiente para predecir")
        return []

    # Agrupar por tipo + descripción similar
    grupos: dict = {}
    for tipo, desc, valor, periodo, estado in rows:
        tok = _tokenizar(desc or '')
        key = (tipo, frozenset(list(tok)[:5]))  # clave aproximada
        if key not in grupos:
            grupos[key] = []
        grupos[key].append({
            'desc': desc, 'valor': valor,
            'periodo': periodo, 'estado': estado, 'tok': tok
        })

    predicciones = []
    for (tipo, _), ocurrencias in grupos.items():
        if len(ocurrencias) < 2:
            continue  # necesita al menos 2 apariciones

        periodos_unicos = len({o['periodo'] for o in ocurrencias})
        if periodos_unicos < 2:
            continue

        # Valor promedio
        valores = [abs(o['valor']) for o in ocurrencias if o['valor']]
        valor_est = sum(valores) / len(valores) if valores else 0.0

        # Confianza: basada en frecuencia y consistencia de valores
        cv = (max(valores) - min(valores)) / (valor_est + 1) if valor_est else 0
        confianza = min(0.95, 0.4 + (periodos_unicos * 0.1) - (cv * 0.1))

        if confianza < min_confianza:
            continue

        # Descripción más representativa (la más frecuente)
        desc_counter = Counter(o['desc'] for o in ocurrencias)
        desc_rep = desc_counter.most_common(1)[0][0]

        pred = {
            'tipo':             tipo,
            'descripcion':      desc_rep,
            'valor_estimado':   valor_est,
            'confianza':        round(confianza, 3),
            'apariciones':      len(ocurrencias),
            'periodos_vistos':  periodos_unicos,
            'periodo_predicho': periodo_predicho,
            'confirmado':       False,
        }
        predicciones.append(pred)

    # Ordenar por confianza desc
    predicciones.sort(key=lambda x: x['confianza'], reverse=True)

    # Guardar en SQLite
    _guardar_predicciones(predicciones, periodo_base)

    log.info("[ml] %d predicciones para %s (base: %s)",
             len(predicciones), periodo_predicho, periodo_base)
    return predicciones


def _guardar_predicciones(predicciones: List[dict], periodo_base: str):
    """Persiste predicciones en la tabla ml_predicciones."""
    try:
        conn = _get_db()
        conn.execute(
            "DELETE FROM ml_predicciones WHERE periodo_predicho=? AND confirmado=0",
            (predicciones[0]['periodo_predicho'],)
        ) if predicciones else None
        for p in predicciones:
            conn.execute(
                """INSERT OR IGNORE INTO ml_predicciones
                   (periodo_base, periodo_predicho, descripcion,
                    valor_estimado, confianza, fecha_prediccion)
                   VALUES (?,?,?,?,?,?)""",
                (periodo_base, p['periodo_predicho'], p['descripcion'],
                 p['valor_estimado'], p['confianza'], _ahora())
            )
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[ml] Error guardando predicciones: %s", e)


def confirmar_prediccion(prediccion_id: int) -> bool:
    """Marca una predicción como confirmada (ocurrió realmente)."""
    try:
        conn = _get_db()
        conn.execute(
            "UPDATE ml_predicciones SET confirmado=1 WHERE id=?", (prediccion_id,)
        )
        conn.commit(); conn.close()
        return True
    except Exception as e:
        log.error("[ml] Error confirmando predicción: %s", e)
        return False


def listar_predicciones(periodo: str = None) -> List[dict]:
    """Lista predicciones ML guardadas."""
    try:
        conn = _get_db()
        sql = ("SELECT id, periodo_predicho, descripcion, valor_estimado, "
               "confianza, confirmado, fecha_prediccion FROM ml_predicciones")
        params = []
        if periodo:
            sql += " WHERE periodo_predicho=?"; params.append(periodo)
        sql += " ORDER BY confianza DESC LIMIT 50"
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return [{'id': r[0], 'periodo': r[1], 'descripcion': r[2],
                 'valor': r[3], 'confianza': r[4],
                 'confirmado': bool(r[5]), 'fecha': r[6]}
                for r in rows]
    except Exception as e:
        log.error("[ml] Error listando predicciones: %s", e)
        return []


def accuracy_modelo() -> dict:
    """Calcula precisión del modelo sobre predicciones pasadas."""
    try:
        conn = _get_db()
        total = conn.execute("SELECT COUNT(*) FROM ml_predicciones").fetchone()[0]
        confirmadas = conn.execute(
            "SELECT COUNT(*) FROM ml_predicciones WHERE confirmado=1"
        ).fetchone()[0]
        conn.close()
        if not total:
            return {'total': 0, 'confirmadas': 0, 'accuracy': 0.0}
        return {
            'total': total,
            'confirmadas': confirmadas,
            'accuracy': round(confirmadas / total * 100, 1)
        }
    except Exception as e:
        log.error("[ml] Error accuracy: %s", e)
        return {'total': 0, 'confirmadas': 0, 'accuracy': 0.0}
