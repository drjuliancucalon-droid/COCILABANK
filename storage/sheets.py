"""
storage/sheets.py
Integración con Google Sheets para exportar/importar datos de conciliación.
Usa gspread si está disponible, con fallback a exportación CSV.
"""
from __future__ import annotations
import logging
import os
from typing import Optional
import pandas as pd

logger = logging.getLogger(__name__)

_GSPREAD_AVAILABLE = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    _GSPREAD_AVAILABLE = True
except ImportError:
    pass


SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


def _get_client(credentials_path: Optional[str] = None):
    """Crea cliente gspread autenticado."""
    if not _GSPREAD_AVAILABLE:
        raise ImportError(
            "gspread y google-auth no están instalados. "
            "Agrega 'gspread' y 'google-auth' a requirements.txt."
        )
    creds_path = credentials_path or os.environ.get(
        "GOOGLE_SHEETS_CREDENTIALS", "google_credentials.json"
    )
    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"No se encontró el archivo de credenciales: {creds_path}. "
            "Descarga el JSON de la cuenta de servicio desde Google Cloud Console."
        )
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


def exportar_a_sheets(
    df: pd.DataFrame,
    spreadsheet_id: str,
    hoja: str = "Conciliacion",
    credentials_path: Optional[str] = None,
    limpiar: bool = True,
) -> str:
    """
    Exporta un DataFrame a Google Sheets.

    Args:
        df: DataFrame a exportar.
        spreadsheet_id: ID del spreadsheet (en la URL de Google Sheets).
        hoja: Nombre de la hoja destino.
        credentials_path: Ruta al JSON de credenciales de cuenta de servicio.
        limpiar: Si True, limpia la hoja antes de escribir.

    Returns:
        URL del spreadsheet actualizado.
    """
    if not _GSPREAD_AVAILABLE:
        raise ImportError("gspread no disponible. Ver storage/sheets.py.")

    client = _get_client(credentials_path)
    sh = client.open_by_key(spreadsheet_id)

    try:
        ws = sh.worksheet(hoja)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=hoja, rows=str(max(len(df) + 10, 100)), cols="26")

    if limpiar:
        ws.clear()

    # Escribir encabezados + datos
    data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    ws.update(data, value_input_option="USER_ENTERED")

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    logger.info("Exportado a Sheets: %s (hoja=%s, filas=%s)", url, hoja, len(df))
    return url


def importar_desde_sheets(
    spreadsheet_id: str,
    hoja: str = "Conciliacion",
    credentials_path: Optional[str] = None,
) -> pd.DataFrame:
    """
    Importa datos desde Google Sheets a un DataFrame.

    Args:
        spreadsheet_id: ID del spreadsheet.
        hoja: Nombre de la hoja a importar.
        credentials_path: Ruta al JSON de credenciales.

    Returns:
        DataFrame con los datos importados.
    """
    if not _GSPREAD_AVAILABLE:
        raise ImportError("gspread no disponible. Ver storage/sheets.py.")

    client = _get_client(credentials_path)
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(hoja)
    records = ws.get_all_records()

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    logger.info("Importado desde Sheets: hoja=%s, filas=%s", hoja, len(df))
    return df


def exportar_multiples_hojas(
    hojas: dict[str, pd.DataFrame],
    spreadsheet_id: str,
    credentials_path: Optional[str] = None,
) -> str:
    """
    Exporta múltiples DataFrames a hojas diferentes del mismo spreadsheet.

    Args:
        hojas: dict {nombre_hoja: DataFrame}.
        spreadsheet_id: ID del spreadsheet.
        credentials_path: Ruta al JSON de credenciales.

    Returns:
        URL del spreadsheet.
    """
    for nombre, df in hojas.items():
        exportar_a_sheets(df, spreadsheet_id, hoja=nombre,
                          credentials_path=credentials_path)
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"


def sheets_disponible() -> bool:
    """Retorna True si gspread y google-auth están instalados."""
    return _GSPREAD_AVAILABLE


def exportar_csv_fallback(
    df: pd.DataFrame,
    ruta: str,
    encoding: str = "utf-8-sig",
) -> str:
    """
    Exporta un DataFrame a CSV cuando Sheets no está disponible.
    utf-8-sig garantiza compatibilidad con Excel en Windows.
    """
    df.to_csv(ruta, index=False, encoding=encoding)
    logger.info("Exportado a CSV: %s", ruta)
    return ruta
