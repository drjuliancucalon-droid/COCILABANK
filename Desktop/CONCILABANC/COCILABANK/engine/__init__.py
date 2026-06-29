"""engine — Motor de conciliación, columnas y aprendizaje NC."""
from engine.reconciliador import comparar_documentos, validar_diferencia_neta
from engine.columna import determinar_columna
from engine.nc_learning import (
    buscar_en_catalogo_nc,
    listar_catalogo_nc,
    _aprender_match_nc,
    _extraer_tokens_nc,
    _uuid_par_nc,
)

__all__ = [
    "comparar_documentos",
    "validar_diferencia_neta",
    "determinar_columna",
    "buscar_en_catalogo_nc",
    "listar_catalogo_nc",
]
