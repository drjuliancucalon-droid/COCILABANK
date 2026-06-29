"""
Calculadora ROI — CREDIEXPRESS POPAYÁN SAS
Muestra cuántas horas ahorra el contador por mes usando el sistema.
"""
import logging
from datetime import datetime
from typing import dict as Dict

log = logging.getLogger(__name__)

# Parámetros de referencia (ajustables en config)
HORAS_MANUAL_POR_MOVIMIENTO    = 0.05   # 3 min promedio manual por movimiento
HORAS_MANUAL_BUSQUEDA_PARTIDA  = 0.25   # 15 min por partida difícil
TARIFA_HORA_CONTADOR_COP       = 75_000  # $75.000/hora contador colombiano


def calcular_roi(
    n_movimientos_banco: int,
    n_movimientos_aux: int,
    n_conciliados: int,
    n_rechazos: int,
    tiempo_proceso_seg: float = 0.0,
    tarifa_hora: float = TARIFA_HORA_CONTADOR_COP,
) -> dict:
    """
    Calcula el ROI del uso del sistema vs proceso manual.

    Returns dict con métricas de ahorro.
    """
    total_mov = n_movimientos_banco + n_movimientos_aux

    # Tiempo manual estimado (horas)
    horas_manual = (
        total_mov * HORAS_MANUAL_POR_MOVIMIENTO
        + n_rechazos * HORAS_MANUAL_BUSQUEDA_PARTIDA
    )

    # Tiempo real del sistema
    horas_sistema = tiempo_proceso_seg / 3600 if tiempo_proceso_seg > 0 else 0.02

    # Ahorro
    horas_ahorradas = max(0.0, horas_manual - horas_sistema)
    minutos_ahorrados = horas_ahorradas * 60
    pesos_ahorrados   = horas_ahorradas * tarifa_hora

    # Porcentaje de automatización
    pct_auto = (n_conciliados / total_mov * 100) if total_mov > 0 else 0.0

    return {
        'total_movimientos':   total_mov,
        'horas_manual_est':    round(horas_manual, 2),
        'horas_sistema':       round(horas_sistema, 2),
        'horas_ahorradas':     round(horas_ahorradas, 2),
        'minutos_ahorrados':   round(minutos_ahorrados, 0),
        'pesos_ahorrados':     round(pesos_ahorrados, 0),
        'pct_automatizacion':  round(pct_auto, 1),
        'tarifa_hora':         tarifa_hora,
        'mensaje': _generar_mensaje(horas_ahorradas, pesos_ahorrados, pct_auto),
    }


def _generar_mensaje(horas: float, pesos: float, pct: float) -> str:
    if horas < 0.5:
        return f"⚡ Proceso completado en segundos — {pct:.0f}% automatizado"
    elif horas < 2:
        return (f"⏱️ Ahorraste {horas:.1f} horas de trabajo manual "
                f"(≈ ${pesos:,.0f} COP en honorarios)")
    elif horas < 8:
        return (f"🚀 ¡{horas:.1f} horas ahorradas! Equivale a "
                f"${pesos:,.0f} COP — {pct:.0f}% automatizado")
    else:
        return (f"🏆 {horas:.0f} horas ahorradas ({horas/8:.1f} días hábiles) — "
                f"ahorro de ${pesos:,.0f} COP este período")


def roi_acumulado_mes() -> dict:
    """Calcula ROI acumulado del mes actual desde el historial."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        mes_actual = datetime.now().strftime('%Y-%m')
        rows = conn.execute(
            "SELECT COUNT(*), COALESCE(AVG(tasa),0) FROM historial "
            "WHERE fecha LIKE ?", (f"{mes_actual}%",)
        ).fetchone()
        conn.close()
        n_conciliaciones = rows[0] or 0
        tasa_promedio    = rows[1] or 0.0
        # Estimación conservadora: 500 movimientos promedio por conciliación
        roi = calcular_roi(
            n_movimientos_banco=250 * n_conciliaciones,
            n_movimientos_aux=250 * n_conciliaciones,
            n_conciliados=int(500 * n_conciliaciones * tasa_promedio / 100),
            n_rechazos=int(500 * n_conciliaciones * (1 - tasa_promedio / 100)),
        )
        roi['n_conciliaciones_mes'] = n_conciliaciones
        roi['tasa_promedio_mes']    = round(tasa_promedio, 1)
        return roi
    except Exception as e:
        log.error("[roi] Error acumulado mes: %s", e)
        return {}


def calendario_fiscal_colombia(year: int = None) -> list:
    """
    Retorna vencimientos DIAN importantes para el año indicado.
    Basado en el calendario tributario colombiano.
    """
    if not year:
        year = datetime.now().year

    vencimientos = [
        # Enero
        {'fecha': f'{year}-01-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Noviembre-Diciembre', 'urgente': False},
        {'fecha': f'{year}-01-31', 'tipo': 'RETENCIÓN', 'descripcion': 'Declaración Retención en la Fuente Diciembre', 'urgente': False},
        # Febrero
        {'fecha': f'{year}-02-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Enero-Febrero (grandes contrib.)', 'urgente': False},
        # Marzo
        {'fecha': f'{year}-03-20', 'tipo': 'ICA', 'descripcion': 'ICA trimestral (según municipio)', 'urgente': False},
        {'fecha': f'{year}-03-31', 'tipo': 'RETENCIÓN', 'descripcion': 'Retención en la Fuente Febrero', 'urgente': False},
        # Abril
        {'fecha': f'{year}-04-15', 'tipo': 'RENTA', 'descripcion': 'Impuesto de Renta Personas Jurídicas (primer grupo)', 'urgente': True},
        {'fecha': f'{year}-04-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Marzo-Abril', 'urgente': False},
        # Mayo
        {'fecha': f'{year}-05-12', 'tipo': 'RENTA', 'descripcion': 'Declaración Renta — segundo grupo NIT', 'urgente': True},
        # Junio
        {'fecha': f'{year}-06-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Mayo-Junio', 'urgente': False},
        {'fecha': f'{year}-06-25', 'tipo': 'MEDIOS', 'descripcion': 'Información exógena — medios magnéticos', 'urgente': True},
        # Agosto
        {'fecha': f'{year}-08-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Julio-Agosto', 'urgente': False},
        # Septiembre
        {'fecha': f'{year}-09-30', 'tipo': 'RETENCIÓN', 'descripcion': 'Declaración Retención en la Fuente Agosto', 'urgente': False},
        # Octubre
        {'fecha': f'{year}-10-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Septiembre-Octubre', 'urgente': False},
        # Noviembre
        {'fecha': f'{year}-11-10', 'tipo': 'RENTA', 'descripcion': 'Segunda cuota Renta Personas Jurídicas', 'urgente': True},
        # Diciembre
        {'fecha': f'{year}-12-20', 'tipo': 'IVA', 'descripcion': 'IVA bimestral Noviembre-Diciembre', 'urgente': False},
        {'fecha': f'{year}-12-31', 'tipo': 'INVENTARIOS', 'descripcion': 'Cierre contable — Inventario físico', 'urgente': False},
    ]

    # Marcar próximos 30 días
    hoy = datetime.now().date()
    for v in vencimientos:
        try:
            fecha = datetime.strptime(v['fecha'], '%Y-%m-%d').date()
            dias = (fecha - hoy).days
            v['dias_restantes'] = dias
            v['proximo'] = 0 <= dias <= 30
            v['vencido']  = dias < 0
        except Exception:
            v['dias_restantes'] = 999
            v['proximo'] = False
            v['vencido']  = False

    return sorted(vencimientos, key=lambda x: x['fecha'])
