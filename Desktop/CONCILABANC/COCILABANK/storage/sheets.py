"""
storage/sheets.py — Sincronización con Google Sheets para CREDIEXPRESS
Modo cloud (OFFLINE_MODE=False): historial y catálogo NC en Google Sheets.
"""
import json
import logging
from datetime import datetime

import streamlit as st

from config import OFFLINE_MODE

log = logging.getLogger(__name__)

# ── Historial en Google Sheets ────────────────────────────────────────────────
def _guardar_historial_sheets(d):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        creds_json = st.secrets.get("GOOGLE_SHEETS_CREDS", None)
        sheet_id   = st.secrets.get("GOOGLE_SHEET_ID",    None)
        if not creds_json or not sheet_id:
            return
        creds = Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"])
        ws = gspread.authorize(creds).open_by_key(sheet_id).sheet1
        if not ws.get_all_values():
            ws.append_row(["Fecha","Banco","Auxiliar","Periodo",
                           "Mov.Banco","Asientos","Exactas","Aprox",
                           "Solo Banco","Solo Aux","Tasa %",
                           "Saldo Banco","Saldo Aux","Diferencia Neta"])
        ws.append_row([
            d["fecha_hora"],d["archivo_banco"],d["archivo_auxiliar"],d["periodo"],
            d["n_banco"],d["n_aux"],d["n_exactas"],d["n_aprox"],
            d["n_solo_banco"],d["n_solo_aux"],round(d["tasa"],1),
            round(d["saldo_banco"] or 0,2),round(d["saldo_aux"] or 0,2),
            round(d["diferencia_neta"] or 0,2)])
    except Exception as e:
        log.error("[_guardar_historial_sheets] %s", e, exc_info=True)

# ── Catálogo NC en Google Sheets ──────────────────────────────────────────────
def _push_catalogo_to_sheets():
    """Sube reglas PENDIENTE_SYNC del SQLite a Google Sheets."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        from storage.db import _init_db
        creds_json = st.secrets.get("GOOGLE_SHEETS_CREDS", None)
        sheet_id   = st.secrets.get("GOOGLE_SHEET_ID",    None)
        if not creds_json or not sheet_id:
            return 0
        creds = Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"])
        gc = gspread.authorize(creds)
        wb = gc.open_by_key(sheet_id)
        try:
            ws = wb.worksheet("nc_catalogo")
        except Exception:
            ws = wb.add_worksheet("nc_catalogo", rows=2000, cols=9)
            ws.append_row(["uuid","banco_tokens","aux_tokens","confirmaciones",
                           "nivel","aprobado_por","fecha_primera","fecha_ultima"])
        conn = _init_db()
        pendientes = conn.execute(
            "SELECT uuid, banco_tokens, aux_tokens, confirmaciones, nivel, "
            "aprobado_por, fecha_primera, fecha_ultima "
            "FROM nc_catalogo WHERE sync_status='PENDIENTE_SYNC'"
        ).fetchall()
        existentes = {r[0] for r in ws.get_all_values()[1:] if r}
        n = 0
        for row in pendientes:
            if row[0] not in existentes:
                ws.append_row(list(row))
                n += 1
            conn.execute(
                "UPDATE nc_catalogo SET sync_status='SINCRONIZADO' WHERE uuid=?",
                (row[0],))
        conn.commit(); conn.close()
        return n
    except Exception as e:
        log.error("[_push_catalogo_to_sheets] %s", e, exc_info=True)
        return 0

def _pull_catalogo_from_sheets():
    """Descarga reglas nuevas de Google Sheets al SQLite local."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        from storage.db import _init_db
        creds_json = st.secrets.get("GOOGLE_SHEETS_CREDS", None)
        sheet_id   = st.secrets.get("GOOGLE_SHEET_ID",    None)
        if not creds_json or not sheet_id:
            return 0
        creds = Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"])
        gc = gspread.authorize(creds)
        wb = gc.open_by_key(sheet_id)
        try:
            ws = wb.worksheet("nc_catalogo")
        except Exception:
            return 0
        rows = ws.get_all_values()
        if len(rows) <= 1:
            return 0
        conn = _init_db()
        n = 0
        for row in rows[1:]:
            if len(row) < 7 or not row[0]:
                continue
            uuid  = row[0]
            conf  = int(row[3]) if str(row[3]).isdigit() else 1
            nivel = row[4] or 'MEDIA'
            existe = conn.execute(
                "SELECT id, confirmaciones FROM nc_catalogo WHERE uuid=?", (uuid,)).fetchone()
            if not existe:
                conn.execute("""INSERT INTO nc_catalogo
                    (uuid, banco_tokens, aux_tokens, confirmaciones, nivel,
                     aprobado_por, fecha_primera, fecha_ultima, sync_status)
                    VALUES (?,?,?,?,?,?,?,?,'SINCRONIZADO')""",
                    (uuid, row[1], row[2], conf, nivel,
                     row[5] if len(row)>5 else 'AUTO',
                     row[6] if len(row)>6 else '',
                     row[7] if len(row)>7 else ''))
                n += 1
            elif conf > existe[1]:
                nuevo_nivel = 'ALTA' if conf >= 5 else 'MEDIA'
                conn.execute(
                    "UPDATE nc_catalogo SET confirmaciones=?, nivel=?, "
                    "sync_status='SINCRONIZADO' WHERE uuid=?",
                    (conf, nuevo_nivel, uuid))
        conn.commit(); conn.close()
        return n
    except Exception as e:
        log.error("[_pull_catalogo_from_sheets] %s", e, exc_info=True)
        return 0

def sincronizar_catalogo_nc():
    """Sincronización bidireccional: push local -> Sheets, pull Sheets -> local."""
    if not OFFLINE_MODE:
        return 0, 0
    n_up   = _push_catalogo_to_sheets()
    n_down = _pull_catalogo_from_sheets()
    return n_up, n_down

# ── Aprendizaje NC en modo cloud ──────────────────────────────────────────────
def _aprender_match_nc_cloud(banco_desc, aux_doc, aux_concepto, metodo):
    """
    Para Streamlit Cloud: registra el aprendizaje NC directamente en Google Sheets
    (tabla nc_aprendizaje del spreadsheet).
    """
    if not aux_doc or not str(aux_doc).upper().startswith('NC-'):
        return
    try:
        from engine.nc_learning import _extraer_tokens_nc, _uuid_par_nc
        banco_tok = _extraer_tokens_nc(banco_desc)
        aux_tok   = _extraer_tokens_nc(aux_concepto)
        if len(banco_tok) < 2 or len(aux_tok) < 2:
            return
        uuid  = _uuid_par_nc(banco_tok, aux_tok)
        ahora = datetime.now().isoformat(timespec='seconds')
        import gspread
        from google.oauth2.service_account import Credentials
        creds_json = st.secrets.get("GOOGLE_SHEETS_CREDS", None)
        sheet_id   = st.secrets.get("GOOGLE_SHEET_ID",    None)
        if not creds_json or not sheet_id:
            return
        creds = Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"])
        gc = gspread.authorize(creds)
        wb = gc.open_by_key(sheet_id)
        try:
            ws = wb.worksheet("nc_aprendizaje")
        except Exception:
            ws = wb.add_worksheet("nc_aprendizaje", rows=2000, cols=8)
            ws.append_row(["uuid","banco_desc_raw","aux_concepto_raw",
                           "banco_tokens","aux_tokens","veces_visto",
                           "fecha_primera","fecha_ultima"])
        rows = ws.get_all_values()
        uuid_map = {r[0]: i+2 for i,r in enumerate(rows[1:]) if r and r[0]}
        if uuid in uuid_map:
            row_n = uuid_map[uuid]
            try:
                vv = int(ws.cell(row_n, 6).value or 1)
            except Exception:
                vv = 1
            ws.update_cell(row_n, 6, vv + 1)
            ws.update_cell(row_n, 8, ahora)
            if vv + 1 >= 3:
                _promover_cloud_to_catalogo(ws, row_n, uuid, banco_tok,
                                            aux_tok, vv+1, wb)
        else:
            ws.append_row([uuid, (banco_desc or '')[:150],
                           (aux_concepto or '')[:150],
                           json.dumps(banco_tok), json.dumps(aux_tok),
                           1, ahora, ahora])
    except Exception as e:
        log.error("[_aprender_match_nc_cloud] %s", e, exc_info=True)

def _promover_cloud_to_catalogo(ws_aprendizaje, row_n, uuid, banco_tok,
                                aux_tok, vv, wb):
    """Promueve un candidato al nc_catalogo en Google Sheets."""
    try:
        try:
            ws_cat = wb.worksheet("nc_catalogo")
        except Exception:
            ws_cat = wb.add_worksheet("nc_catalogo", rows=2000, cols=9)
            ws_cat.append_row(["uuid","banco_tokens","aux_tokens","confirmaciones",
                               "nivel","aprobado_por","fecha_primera","fecha_ultima"])
        existentes = {r[0] for r in ws_cat.get_all_values()[1:] if r}
        if uuid not in existentes:
            ahora = datetime.now().isoformat(timespec='seconds')
            nivel = 'ALTA' if vv >= 5 else 'MEDIA'
            ws_cat.append_row([uuid, json.dumps(banco_tok), json.dumps(aux_tok),
                               vv, nivel, 'AUTO', ahora, ahora])
    except Exception as e:
        log.error("[_promover_cloud_to_catalogo] %s", e, exc_info=True)
