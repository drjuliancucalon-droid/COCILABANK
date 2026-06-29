"""
storage/db.py — Persistencia SQLite para CREDIEXPRESS Conciliación Bancaria
Tablas: historial, pdf_formatos, nc_catalogo, nc_aprendizaje, nc_historial_match
"""
import sqlite3
import json
import hashlib
import logging
import os
from datetime import datetime
from pathlib import Path

from config import BASE_DIR, DB_PATH, OFFLINE_MODE
from storage.migrations import MigrationManager

log = logging.getLogger(__name__)

# ── Inicialización ─────────────────────────────────────────────────────────────
def _init_db():
    conn = sqlite3.connect(DB_PATH)
    MigrationManager(conn).apply_pending()
    return conn

# ── Auto-guardado de archivos originales ──────────────────────────────────────
def _auto_guardar_archivo(uploaded_file, subfolder="datos_entrada"):
    if not OFFLINE_MODE or uploaded_file is None:
        return None, False
    dest_dir = os.path.join(BASE_DIR, subfolder, datetime.now().strftime("%Y-%m"))
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, uploaded_file.name)
    if not os.path.exists(dest):
        with open(dest, "wb") as f:
            f.write(uploaded_file.getvalue())
        return dest, True
    return dest, False

def _auto_guardar_excel(excel_bytes, nombre):
    if not OFFLINE_MODE:
        return None
    ts      = datetime.now().strftime("%Y-%m-%d_%H-%M")
    destdir = os.path.join(BASE_DIR, "reportes_excel", ts)
    os.makedirs(destdir, exist_ok=True)
    dest = os.path.join(destdir, nombre)
    with open(dest, "wb") as f:
        f.write(excel_bytes)
    return dest

# ── Historial de conciliaciones ───────────────────────────────────────────────
def _guardar_historial_sqlite(d):
    try:
        conn = _init_db()
        conn.execute("""INSERT INTO historial
            (fecha_hora,archivo_banco,archivo_auxiliar,periodo,
             n_banco,n_aux,n_exactas,n_aprox,n_solo_banco,n_solo_aux,
             tasa,saldo_banco,saldo_aux,diferencia_neta,excel_path)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (d["fecha_hora"],d["archivo_banco"],d["archivo_auxiliar"],d["periodo"],
             d["n_banco"],d["n_aux"],d["n_exactas"],d["n_aprox"],
             d["n_solo_banco"],d["n_solo_aux"],d["tasa"],
             d["saldo_banco"],d["saldo_aux"],d["diferencia_neta"],d.get("excel_path","")))
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[_guardar_historial_sqlite] %s", e, exc_info=True)

def leer_historial_sqlite(limite=8):
    try:
        if not os.path.exists(DB_PATH): return []
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            """SELECT fecha_hora,archivo_banco,archivo_auxiliar,periodo,
                      tasa,n_exactas,n_banco,diferencia_neta
               FROM historial ORDER BY id DESC LIMIT ?""", (limite,)).fetchall()
        conn.close(); return rows
    except Exception as e:
        log.error("[leer_historial_sqlite] %s", e, exc_info=True)
        return []

# ── Catálogo de formatos PDF aprendidos ───────────────────────────────────────
def _firma_pdf(nombre_archivo, n_columnas, banco_detectado=""):
    """Genera firma MD5 única por nombre normalizado + nro. columnas + banco."""
    raw = f"{nombre_archivo}_{n_columnas}_{banco_detectado}".lower()
    return hashlib.md5(raw.encode()).hexdigest()[:16]

def registrar_formato_pdf(nombre_archivo, tipo_doc, columnas, fmt_fecha,
                          prefijos_doc, banco_detectado="", nombre_formato=""):
    """Guarda o actualiza el patron de un archivo procesado exitosamente.
    nombre_formato: nombre del entry ganador en REGISTRO_FORMATOS (ej. 'Bancolombia — PDF').
    """
    if not OFFLINE_MODE:
        return
    try:
        nombre_base = Path(nombre_archivo).name.lower()
        firma = _firma_pdf(nombre_archivo, len(columnas) if columnas else 0, banco_detectado)
        ahora = datetime.now().isoformat(timespec='seconds')
        conn  = _init_db()
        existe = conn.execute(
            "SELECT id, usos FROM pdf_formatos WHERE firma=?", (firma,)).fetchone()
        if existe:
            conn.execute(
                """UPDATE pdf_formatos
                   SET usos=?, ultima_vez=?, nombre_base=?, nombre_formato=?
                   WHERE firma=?""",
                (existe[1] + 1, ahora, nombre_base, nombre_formato or banco_detectado, firma))
        else:
            conn.execute("""INSERT INTO pdf_formatos
                (firma, tipo_doc, columnas, fmt_fecha, prefijos_doc,
                 banco_detectado, usos, ultima_vez, nombre_base, nombre_formato)
                VALUES (?,?,?,?,?,?,1,?,?,?)""",
                (firma, tipo_doc,
                 json.dumps(columnas or [], ensure_ascii=False),
                 fmt_fecha or '',
                 json.dumps(prefijos_doc or [], ensure_ascii=False),
                 banco_detectado or '', ahora,
                 nombre_base, nombre_formato or banco_detectado))
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[registrar_formato_pdf] %s", e, exc_info=True)


def buscar_formato_rapido(nombre_archivo, tipo_doc):
    """Lookup rapido por nombre base de archivo + tipo — sin necesitar n_columnas.
    Devuelve nombre_formato (str) del parser que funcionó antes, o None.
    Solo retorna si fue usado >= 2 veces (evita falsos positivos en primer uso).
    """
    if not OFFLINE_MODE:
        return None
    try:
        nombre_base = Path(nombre_archivo).name.lower()
        conn = _init_db()
        row = conn.execute(
            """SELECT nombre_formato, usos FROM pdf_formatos
               WHERE nombre_base = ? AND tipo_doc = ? AND nombre_formato != ''
               ORDER BY usos DESC LIMIT 1""",
            (nombre_base, tipo_doc)
        ).fetchone()
        conn.close()
        if row and row[1] >= 2:
            return row[0]   # nombre del formato ganador (ej. 'Davivienda — PDF Estado de Cuenta')
        return None
    except Exception as e:
        log.error("[buscar_formato_rapido] %s", e, exc_info=True)
        return None

def buscar_formato_pdf(nombre_archivo, n_columnas, banco_detectado=""):
    """Devuelve dict con info del formato guardado, o None si no existe."""
    if not OFFLINE_MODE:
        return None
    try:
        firma = _firma_pdf(nombre_archivo, n_columnas, banco_detectado)
        conn  = _init_db()
        row   = conn.execute(
            """SELECT tipo_doc, columnas, fmt_fecha, prefijos_doc,
                      banco_detectado, usos, ultima_vez
               FROM pdf_formatos WHERE firma=?""", (firma,)).fetchone()
        conn.close()
        if not row:
            return None
        return {
            'tipo_doc'       : row[0],
            'columnas'       : json.loads(row[1] or '[]'),
            'fmt_fecha'      : row[2],
            'prefijos_doc'   : json.loads(row[3] or '[]'),
            'banco_detectado': row[4],
            'usos'           : row[5],
            'ultima_vez'     : row[6],
        }
    except Exception as e:
        log.error("[buscar_formato_pdf] %s", e, exc_info=True)
        return None

def listar_formatos_aprendidos():
    """Devuelve todos los formatos guardados en el catalogo."""
    try:
        if not os.path.exists(DB_PATH): return []
        conn = _init_db()
        rows = conn.execute(
            """SELECT firma, tipo_doc, banco_detectado, usos, ultima_vez
               FROM pdf_formatos ORDER BY usos DESC""").fetchall()
        conn.close()
        return rows
    except Exception as e:
        log.error("[listar_formatos_aprendidos] %s", e, exc_info=True)
        return []
