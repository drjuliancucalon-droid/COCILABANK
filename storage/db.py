"""
storage/db.py
Gestor de base de datos SQLite para COCILABANK / CREDIEXPRESS POPAYÁN SAS.
Maneja conciliaciones, partidas, usuarios, auditoría y configuración.
"""
from __future__ import annotations
import sqlite3
import os
import json
import hashlib
import logging
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Optional

logger = logging.getLogger(__name__)

DB_PATH = os.environ.get("COCILABANK_DB", "conciliaciones.db")


def _get_db_path() -> str:
    """Retorna la ruta de la DB, buscando en el directorio del script primero."""
    if os.path.isabs(DB_PATH):
        return DB_PATH
    # Buscar junto al script principal
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "..", DB_PATH)


@contextmanager
def get_conn(path: Optional[str] = None):
    """Context manager para conexión SQLite con WAL y foreign keys."""
    db = path or _get_db_path()
    conn = sqlite3.connect(db, timeout=30, detect_types=sqlite3.PARSE_DECLTYPES)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(path: Optional[str] = None) -> None:
    """Inicializa todas las tablas si no existen."""
    from storage.migrations import run_migrations
    run_migrations(path)


# ── CONCILIACIONES ─────────────────────────────────────────────────────────────

def guardar_conciliacion(
    periodo: str,
    empresa: str,
    banco: str,
    resumen: dict,
    usuario: str = "sistema",
    path: Optional[str] = None,
) -> int:
    """
    Guarda el resultado de una conciliación.
    Retorna el ID insertado.
    """
    with get_conn(path) as conn:
        cur = conn.execute(
            """
            INSERT INTO conciliaciones
                (periodo, empresa, banco, resumen_json, usuario, fecha_hora)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                periodo,
                empresa,
                banco,
                json.dumps(resumen, ensure_ascii=False),
                usuario,
                datetime.now().isoformat(),
            ),
        )
        rid = cur.lastrowid
        logger.info("Conciliación guardada id=%s periodo=%s", rid, periodo)
        return rid


def listar_conciliaciones(
    empresa: Optional[str] = None,
    limit: int = 100,
    path: Optional[str] = None,
) -> list[dict]:
    """Lista conciliaciones ordenadas por fecha desc."""
    with get_conn(path) as conn:
        if empresa:
            rows = conn.execute(
                "SELECT * FROM conciliaciones WHERE empresa=? ORDER BY fecha_hora DESC LIMIT ?",
                (empresa, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM conciliaciones ORDER BY fecha_hora DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]


def obtener_conciliacion(rid: int, path: Optional[str] = None) -> Optional[dict]:
    """Obtiene una conciliación por ID."""
    with get_conn(path) as conn:
        row = conn.execute(
            "SELECT * FROM conciliaciones WHERE id=?", (rid,)
        ).fetchone()
        if row:
            d = dict(row)
            d["resumen"] = json.loads(d.get("resumen_json") or "{}")
            return d
        return None


def eliminar_conciliacion(rid: int, path: Optional[str] = None) -> bool:
    """Elimina una conciliación por ID."""
    with get_conn(path) as conn:
        cur = conn.execute("DELETE FROM conciliaciones WHERE id=?", (rid,))
        return cur.rowcount > 0


# ── PARTIDAS ───────────────────────────────────────────────────────────────────

def guardar_partidas(
    conciliacion_id: int,
    partidas: list[dict],
    tipo: str = "abierta",
    path: Optional[str] = None,
) -> int:
    """Guarda lista de partidas asociadas a una conciliación. Retorna filas insertadas."""
    with get_conn(path) as conn:
        conn.execute(
            "DELETE FROM partidas WHERE conciliacion_id=? AND tipo=?",
            (conciliacion_id, tipo),
        )
        data = [
            (
                conciliacion_id,
                tipo,
                p.get("fecha"),
                p.get("descripcion"),
                p.get("valor", 0),
                p.get("referencia"),
                p.get("origen"),
                p.get("estado", "pendiente"),
                json.dumps(p, ensure_ascii=False),
            )
            for p in partidas
        ]
        conn.executemany(
            """
            INSERT INTO partidas
                (conciliacion_id, tipo, fecha, descripcion, valor,
                 referencia, origen, estado, data_json)
            VALUES (?,?,?,?,?,?,?,?,?)
            """,
            data,
        )
        return len(data)


def listar_partidas(
    conciliacion_id: int,
    tipo: Optional[str] = None,
    path: Optional[str] = None,
) -> list[dict]:
    """Lista partidas de una conciliación."""
    with get_conn(path) as conn:
        if tipo:
            rows = conn.execute(
                "SELECT * FROM partidas WHERE conciliacion_id=? AND tipo=?",
                (conciliacion_id, tipo),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM partidas WHERE conciliacion_id=?",
                (conciliacion_id,),
            ).fetchall()
        result = []
        for r in rows:
            d = dict(r)
            d["data"] = json.loads(d.get("data_json") or "{}")
            result.append(d)
        return result


# ── CONFIGURACIÓN ──────────────────────────────────────────────────────────────

def set_config(clave: str, valor: Any, path: Optional[str] = None) -> None:
    """Guarda o actualiza un valor de configuración."""
    with get_conn(path) as conn:
        conn.execute(
            """
            INSERT INTO configuracion (clave, valor, actualizado)
            VALUES (?, ?, ?)
            ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor, actualizado=excluded.actualizado
            """,
            (clave, json.dumps(valor, ensure_ascii=False), datetime.now().isoformat()),
        )


def get_config(clave: str, default: Any = None, path: Optional[str] = None) -> Any:
    """Obtiene un valor de configuración."""
    with get_conn(path) as conn:
        row = conn.execute(
            "SELECT valor FROM configuracion WHERE clave=?", (clave,)
        ).fetchone()
        if row:
            return json.loads(row["valor"])
        return default


def all_config(path: Optional[str] = None) -> dict:
    """Retorna toda la configuración como dict."""
    with get_conn(path) as conn:
        rows = conn.execute("SELECT clave, valor FROM configuracion").fetchall()
        return {r["clave"]: json.loads(r["valor"]) for r in rows}


# ── AUDITORÍA ──────────────────────────────────────────────────────────────────

def log_auditoria(
    accion: str,
    usuario: str,
    detalle: str = "",
    path: Optional[str] = None,
) -> None:
    """Registra una acción en el log de auditoría."""
    with get_conn(path) as conn:
        conn.execute(
            "INSERT INTO auditoria (accion, usuario, detalle, fecha_hora) VALUES (?,?,?,?)",
            (accion, usuario, detalle, datetime.now().isoformat()),
        )


def listar_auditoria(
    limit: int = 200,
    usuario: Optional[str] = None,
    path: Optional[str] = None,
) -> list[dict]:
    """Lista entradas de auditoría."""
    with get_conn(path) as conn:
        if usuario:
            rows = conn.execute(
                "SELECT * FROM auditoria WHERE usuario=? ORDER BY fecha_hora DESC LIMIT ?",
                (usuario, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM auditoria ORDER BY fecha_hora DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]


# ── REGLAS ML ─────────────────────────────────────────────────────────────────

def guardar_regla_ml(
    patron: str,
    categoria: str,
    confianza: float = 1.0,
    path: Optional[str] = None,
) -> None:
    """Guarda o actualiza una regla de clasificación ML."""
    key = hashlib.md5(patron.encode()).hexdigest()
    with get_conn(path) as conn:
        conn.execute(
            """
            INSERT INTO reglas_ml (hash_patron, patron, categoria, confianza, fecha_hora)
            VALUES (?,?,?,?,?)
            ON CONFLICT(hash_patron) DO UPDATE SET
                categoria=excluded.categoria,
                confianza=excluded.confianza,
                fecha_hora=excluded.fecha_hora
            """,
            (key, patron, categoria, confianza, datetime.now().isoformat()),
        )


def listar_reglas_ml(path: Optional[str] = None) -> list[dict]:
    """Lista todas las reglas ML."""
    with get_conn(path) as conn:
        rows = conn.execute(
            "SELECT * FROM reglas_ml ORDER BY confianza DESC"
        ).fetchall()
        return [dict(r) for r in rows]


# ── ESTADÍSTICAS ───────────────────────────────────────────────────────────────

def estadisticas_globales(path: Optional[str] = None) -> dict:
    """Retorna estadísticas generales del sistema."""
    with get_conn(path) as conn:
        total_conc = conn.execute(
            "SELECT COUNT(*) AS n FROM conciliaciones"
        ).fetchone()["n"]
        total_part = conn.execute(
            "SELECT COUNT(*) AS n FROM partidas"
        ).fetchone()["n"]
        last_conc = conn.execute(
            "SELECT fecha_hora FROM conciliaciones ORDER BY fecha_hora DESC LIMIT 1"
        ).fetchone()
        return {
            "total_conciliaciones": total_conc,
            "total_partidas": total_part,
            "ultima_conciliacion": last_conc["fecha_hora"] if last_conc else None,
        }
