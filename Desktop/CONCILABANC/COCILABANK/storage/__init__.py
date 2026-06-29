"""
storage/__init__.py — Punto de entrada unificado para persistencia
Enruta a SQLite (offline) o Google Sheets (cloud) según OFFLINE_MODE.
"""
from config import OFFLINE_MODE
from storage.db import (
    _init_db,
    _auto_guardar_archivo,
    _auto_guardar_excel,
    _guardar_historial_sqlite,
    leer_historial_sqlite,
    registrar_formato_pdf,
    buscar_formato_pdf,
    listar_formatos_aprendidos,
)
from storage.sheets import (
    _guardar_historial_sheets,
    sincronizar_catalogo_nc,
    _aprender_match_nc_cloud,
)

__all__ = [
    "_init_db",
    "_auto_guardar_archivo",
    "_auto_guardar_excel",
    "guardar_historial",
    "leer_historial_sqlite",
    "registrar_formato_pdf",
    "buscar_formato_pdf",
    "listar_formatos_aprendidos",
    "registrar_aprendizaje_nc",
    "sincronizar_catalogo_nc",
]

def guardar_historial(d):
    """Punto de entrada: SQLite si es offline, Google Sheets si es cloud."""
    if OFFLINE_MODE:
        _guardar_historial_sqlite(d)
    else:
        _guardar_historial_sheets(d)

def registrar_aprendizaje_nc(banco_desc, aux_doc, aux_concepto, metodo,
                             valor_banco=None, valor_aux=None):
    """Punto de entrada unificado: SQLite (offline) o Sheets (cloud)."""
    if OFFLINE_MODE:
        from engine.nc_learning import _aprender_match_nc
        _aprender_match_nc(banco_desc, aux_doc, aux_concepto, metodo,
                           valor_banco, valor_aux)
    else:
        _aprender_match_nc_cloud(banco_desc, aux_doc, aux_concepto, metodo)
