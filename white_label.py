"""
Configuración White Label — CREDIEXPRESS POPAYÁN SAS
Permite a firmas contables personalizar la aplicación con su propia marca.
"""
import logging
import os
from typing import Optional

log = logging.getLogger(__name__)

# ── Valores por defecto (se sobreescriben con config de SQLite) ───────────────
_DEFAULTS = {
    'empresa_nombre':          'CREDIEXPRESS POPAYÁN SAS',
    'empresa_nit':             '900000000-0',
    'empresa_ciudad':          'Popayán, Cauca',
    'empresa_color_primario':  '#1F4E79',
    'empresa_color_secundario': '#C9A227',
    'empresa_color_fondo':     '#0D1B2A',
    'empresa_color_texto':     '#E8F4FD',
    'plan_activo':             'profesional',
    'tema_default':            'oscuro',
    'logo_path':               '',
    'app_titulo':              'Conciliación Bancaria',
    'app_subtitulo':           'Sistema Profesional para Contadores Colombianos',
}

_cache: dict = {}


def _leer_config() -> dict:
    """Lee configuración desde SQLite (con fallback a defaults)."""
    global _cache
    if _cache:
        return _cache
    try:
        from storage.db import _init_db
        conn = _init_db()
        rows = conn.execute(
            "SELECT clave, valor FROM configuracion_empresa"
        ).fetchall()
        conn.close()
        _cache = dict(_DEFAULTS)
        _cache.update({r[0]: r[1] for r in rows if r[1] is not None})
    except Exception as e:
        log.warning("[white_label] Usando defaults: %s", e)
        _cache = dict(_DEFAULTS)
    return _cache


def invalidar_cache():
    """Fuerza recarga de config en el siguiente acceso."""
    global _cache
    _cache = {}


def get(clave: str, default: str = '') -> str:
    """Obtiene un valor de configuración de marca."""
    return _leer_config().get(clave, _DEFAULTS.get(clave, default))


def set_config(clave: str, valor: str, usuario: str = 'admin') -> bool:
    """Actualiza un valor de configuración en SQLite."""
    try:
        from storage.db import _init_db
        from datetime import datetime
        conn = _init_db()
        conn.execute(
            """INSERT INTO configuracion_empresa (clave, valor, modificado_por, fecha_modificacion)
               VALUES (?,?,?,?)
               ON CONFLICT(clave) DO UPDATE SET
               valor=excluded.valor,
               modificado_por=excluded.modificado_por,
               fecha_modificacion=excluded.fecha_modificacion""",
            (clave, valor, usuario, datetime.now().isoformat(timespec='seconds'))
        )
        conn.commit(); conn.close()
        invalidar_cache()
        return True
    except Exception as e:
        log.error("[white_label] Error guardando config: %s", e)
        return False


def get_css_variables() -> str:
    """
    Genera variables CSS para el tema de la empresa.
    Se inyectan en el CSS de Streamlit.
    """
    cfg = _leer_config()
    c1  = cfg.get('empresa_color_primario',   '#1F4E79')
    c2  = cfg.get('empresa_color_secundario', '#C9A227')
    bg  = cfg.get('empresa_color_fondo',      '#0D1B2A')
    txt = cfg.get('empresa_color_texto',       '#E8F4FD')
    tema = cfg.get('tema_default', 'oscuro')

    if tema == 'claro':
        bg  = '#F4F6F9'
        txt = '#1A1A2E'
        bg2 = '#FFFFFF'
        bg3 = '#EBF5FB'
    else:
        bg2 = '#162032'
        bg3 = '#1E3A5F'

    return f"""
    :root {{
        --color-primario:    {c1};
        --color-secundario:  {c2};
        --color-fondo:       {bg};
        --color-fondo-2:     {bg2};
        --color-fondo-3:     {bg3};
        --color-texto:       {txt};
        --color-borde:       {c1}44;
    }}
    """


def es_plan(plan_requerido: str) -> bool:
    """Verifica si el plan activo tiene acceso a una funcionalidad."""
    jerarquia = {'starter': 0, 'profesional': 1, 'empresarial': 2}
    plan_actual = get('plan_activo', 'profesional')
    return jerarquia.get(plan_actual, 0) >= jerarquia.get(plan_requerido, 0)


def get_all() -> dict:
    """Retorna toda la configuración."""
    return dict(_leer_config())
