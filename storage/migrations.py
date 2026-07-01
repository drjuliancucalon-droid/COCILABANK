"""
storage/migrations.py
Migraciones de base de datos para COCILABANK.
Ejecuta DDL de forma idempotente con versiones.
"""
from __future__ import annotations
import sqlite3
import logging
from typing import Optional

logger = logging.getLogger(__name__)

# Cada migración tiene un número de versión y el SQL correspondiente
MIGRATIONS: list[tuple[int, str]] = [
    (1, """
    CREATE TABLE IF NOT EXISTS conciliaciones (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        periodo     TEXT NOT NULL,
        empresa     TEXT NOT NULL DEFAULT 'CREDIEXPRESS',
        banco       TEXT NOT NULL DEFAULT '',
        resumen_json TEXT NOT NULL DEFAULT '{}',
        usuario     TEXT NOT NULL DEFAULT 'sistema',
        fecha_hora  TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_conc_periodo ON conciliaciones(periodo);
    CREATE INDEX IF NOT EXISTS idx_conc_empresa ON conciliaciones(empresa);
    """),
    (2, """
    CREATE TABLE IF NOT EXISTS partidas (
        id                INTEGER PRIMARY KEY AUTOINCREMENT,
        conciliacion_id   INTEGER NOT NULL REFERENCES conciliaciones(id) ON DELETE CASCADE,
        tipo              TEXT NOT NULL DEFAULT 'abierta',
        fecha             TEXT,
        descripcion       TEXT,
        valor             REAL DEFAULT 0,
        referencia        TEXT,
        origen            TEXT,
        estado            TEXT DEFAULT 'pendiente',
        data_json         TEXT DEFAULT '{}'
    );
    CREATE INDEX IF NOT EXISTS idx_part_conc ON partidas(conciliacion_id);
    CREATE INDEX IF NOT EXISTS idx_part_tipo ON partidas(tipo);
    """),
    (3, """
    CREATE TABLE IF NOT EXISTS configuracion (
        clave       TEXT PRIMARY KEY,
        valor       TEXT NOT NULL DEFAULT 'null',
        actualizado TEXT NOT NULL DEFAULT ''
    );
    """),
    (4, """
    CREATE TABLE IF NOT EXISTS auditoria (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        accion      TEXT NOT NULL,
        usuario     TEXT NOT NULL DEFAULT 'sistema',
        detalle     TEXT DEFAULT '',
        fecha_hora  TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_aud_usuario ON auditoria(usuario);
    CREATE INDEX IF NOT EXISTS idx_aud_fecha ON auditoria(fecha_hora);
    """),
    (5, """
    CREATE TABLE IF NOT EXISTS reglas_ml (
        hash_patron TEXT PRIMARY KEY,
        patron      TEXT NOT NULL,
        categoria   TEXT NOT NULL,
        confianza   REAL DEFAULT 1.0,
        fecha_hora  TEXT NOT NULL
    );
    """),
    (6, """
    CREATE TABLE IF NOT EXISTS notification_queue (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        channel     TEXT NOT NULL,
        recipient   TEXT NOT NULL,
        subject     TEXT DEFAULT '',
        body        TEXT NOT NULL,
        status      TEXT DEFAULT 'pending',
        attempts    INTEGER DEFAULT 0,
        max_attempts INTEGER DEFAULT 3,
        created_at  TEXT NOT NULL,
        next_try_at TEXT NOT NULL,
        sent_at     TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_nq_status ON notification_queue(status);
    CREATE INDEX IF NOT EXISTS idx_nq_next ON notification_queue(next_try_at);
    """),
    (7, """
    CREATE TABLE IF NOT EXISTS usuarios (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        username    TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        rol         TEXT NOT NULL DEFAULT 'viewer',
        activo      INTEGER NOT NULL DEFAULT 1,
        empresa     TEXT DEFAULT 'CREDIEXPRESS',
        creado_en   TEXT NOT NULL,
        ultimo_login TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_usr_username ON usuarios(username);
    """),
]


def _get_version(conn: sqlite3.Connection) -> int:
    """Obtiene la versión actual del schema."""
    conn.execute(
        "CREATE TABLE IF NOT EXISTS schema_version (version INTEGER NOT NULL DEFAULT 0)"
    )
    row = conn.execute("SELECT MAX(version) FROM schema_version").fetchone()
    return row[0] or 0


def _set_version(conn: sqlite3.Connection, version: int) -> None:
    conn.execute("DELETE FROM schema_version")
    conn.execute("INSERT INTO schema_version (version) VALUES (?)", (version,))


def run_migrations(db_path: Optional[str] = None) -> int:
    """
    Ejecuta todas las migraciones pendientes.
    Retorna el número de migraciones aplicadas.
    """
    if db_path is None:
        import os
        base = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(base, "..", "conciliaciones.db")

    conn = sqlite3.connect(db_path, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")

    applied = 0
    try:
        current = _get_version(conn)
        for version, sql in MIGRATIONS:
            if version > current:
                logger.info("Aplicando migración v%s", version)
                for stmt in sql.strip().split(";"):
                    stmt = stmt.strip()
                    if stmt:
                        conn.execute(stmt)
                current = version
                applied += 1
        _set_version(conn, current)
        conn.commit()
        logger.info("DB actualizada a v%s (%s migraciones aplicadas)", current, applied)
    except Exception as e:
        conn.rollback()
        logger.error("Error en migración: %s", e)
        raise
    finally:
        conn.close()

    return applied


def reset_db(db_path: Optional[str] = None) -> None:
    """
    PELIGRO: Elimina todas las tablas y recrea desde cero.
    Solo para desarrollo/testing.
    """
    import os
    if db_path is None:
        base = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(base, "..", "conciliaciones.db")

    if os.path.exists(db_path):
        os.remove(db_path)
        logger.warning("DB eliminada: %s", db_path)

    run_migrations(db_path)
    logger.info("DB recreada desde cero: %s", db_path)
