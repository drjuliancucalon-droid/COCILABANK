"""
Gestión de versiones del esquema SQLite — CREDIEXPRESS POPAYÁN SAS
Incluye todas las tablas del plan de producto comercial.
"""
import logging
import sqlite3

log = logging.getLogger(__name__)

_MIGRATIONS = [
    # ── v1: Tablas iniciales ───────────────────────────────────────────────
    (1, "Tablas iniciales", [
        """CREATE TABLE IF NOT EXISTS historial (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha         TEXT,
            banco         TEXT,
            auxiliar      TEXT,
            periodo       TEXT,
            tasa          REAL,
            saldo_banco   REAL,
            saldo_aux     REAL,
            diferencia_neta REAL DEFAULT 0,
            excel_path    TEXT,
            usuario       TEXT DEFAULT 'admin'
        )""",
        """CREATE TABLE IF NOT EXISTS pdf_formatos (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            firma         TEXT UNIQUE,
            banco         TEXT,
            n_columnas    INTEGER,
            mapa_columnas TEXT,
            fecha_registro TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS nc_catalogo (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid          TEXT UNIQUE,
            banco_tokens  TEXT,
            aux_tokens    TEXT,
            confirmaciones INTEGER DEFAULT 1,
            nivel         TEXT DEFAULT 'BAJA',
            aprobado_por  TEXT DEFAULT 'AUTO',
            fecha_primera TEXT,
            fecha_ultima  TEXT,
            sync_status   TEXT DEFAULT 'PENDIENTE_SYNC'
        )""",
        """CREATE TABLE IF NOT EXISTS nc_aprendizaje (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid          TEXT UNIQUE,
            banco_desc_raw TEXT,
            aux_concepto_raw TEXT,
            banco_tokens  TEXT,
            aux_tokens    TEXT,
            veces_visto   INTEGER DEFAULT 1,
            fecha_primera TEXT,
            fecha_ultima  TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS nc_historial_match (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid_par      TEXT,
            banco_desc    TEXT,
            aux_doc       TEXT,
            aux_concepto  TEXT,
            metodo        TEXT,
            valor_banco   REAL,
            valor_aux     REAL,
            fecha_match   TEXT
        )""",
    ]),

    # ── v2: Multi-usuario + Auditoría ─────────────────────────────────────
    (2, "Multi-usuario y auditoría", [
        """CREATE TABLE IF NOT EXISTS usuarios (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            rol           TEXT NOT NULL DEFAULT 'auxiliar',
            nombre_completo TEXT,
            email         TEXT,
            activo        INTEGER DEFAULT 1,
            fecha_creacion TEXT,
            ultimo_acceso TEXT,
            creado_por    TEXT DEFAULT 'admin'
        )""",
        """CREATE TABLE IF NOT EXISTS sesiones (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT NOT NULL,
            token         TEXT UNIQUE NOT NULL,
            ip            TEXT,
            fecha_inicio  TEXT,
            fecha_expira  TEXT,
            activa        INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS auditoria (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha         TEXT NOT NULL,
            usuario       TEXT NOT NULL,
            accion        TEXT NOT NULL,
            modulo        TEXT,
            detalle       TEXT,
            ip            TEXT,
            resultado     TEXT DEFAULT 'OK'
        )""",
        # Insertar usuario admin por defecto (password: Admin2024*)
        # Hash bcrypt generado offline; se reemplaza en primer login
        """INSERT OR IGNORE INTO usuarios
           (username, password_hash, rol, nombre_completo, fecha_creacion)
           VALUES ('admin',
                   '$2b$12$placeholder_change_on_first_login_admin2024',
                   'admin', 'Administrador', datetime('now'))""",
    ]),

    # ── v3: Partidas conciliatorias ────────────────────────────────────────
    (3, "Partidas conciliatorias", [
        """CREATE TABLE IF NOT EXISTS partidas_conciliatorias (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid          TEXT UNIQUE NOT NULL,
            periodo_origen TEXT NOT NULL,
            tipo          TEXT NOT NULL,
            descripcion   TEXT,
            valor         REAL NOT NULL,
            banco         TEXT,
            doc_referencia TEXT,
            estado        TEXT DEFAULT 'PENDIENTE',
            periodo_cierre TEXT,
            usuario_registro TEXT,
            fecha_registro TEXT,
            fecha_cierre  TEXT,
            observaciones TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS partidas_historial (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            partida_uuid  TEXT NOT NULL,
            periodo       TEXT NOT NULL,
            accion        TEXT NOT NULL,
            usuario       TEXT,
            detalle       TEXT,
            fecha         TEXT
        )""",
    ]),

    # ── v4: Notificaciones + Backup queue ─────────────────────────────────
    (4, "Cola notificaciones y backup", [
        """CREATE TABLE IF NOT EXISTS notificaciones_queue (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo          TEXT NOT NULL,
            canal         TEXT NOT NULL,
            destinatario  TEXT NOT NULL,
            asunto        TEXT,
            cuerpo        TEXT,
            estado        TEXT DEFAULT 'PENDIENTE',
            intentos      INTEGER DEFAULT 0,
            fecha_creacion TEXT,
            fecha_envio   TEXT,
            error_msg     TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS backup_queue (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            archivo_local TEXT NOT NULL,
            destino       TEXT NOT NULL,
            tipo          TEXT DEFAULT 'GOOGLE_DRIVE',
            estado        TEXT DEFAULT 'PENDIENTE',
            intentos      INTEGER DEFAULT 0,
            fecha_creacion TEXT,
            fecha_subida  TEXT,
            error_msg     TEXT
        )""",
    ]),

    # ── v5: DIAN + Comisiones + PUC ───────────────────────────────────────
    (5, "DIAN, comisiones bancarias y PUC", [
        """CREATE TABLE IF NOT EXISTS exportaciones_dian (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo       TEXT NOT NULL,
            tipo          TEXT NOT NULL,
            archivo_xml   TEXT,
            hash_sha256   TEXT,
            estado        TEXT DEFAULT 'BORRADOR',
            usuario       TEXT,
            fecha_generacion TEXT,
            fecha_envio   TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS comisiones_detectadas (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo       TEXT,
            banco         TEXT,
            tipo_comision TEXT,
            descripcion   TEXT,
            valor         REAL,
            fecha_transaccion TEXT,
            revisado      INTEGER DEFAULT 0,
            fecha_deteccion TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS puc_asignaciones (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            descripcion_banco TEXT,
            codigo_puc    TEXT,
            nombre_cuenta TEXT,
            naturaleza    TEXT,
            usuario       TEXT,
            fecha         TEXT,
            confirmaciones INTEGER DEFAULT 1
        )""",
    ]),

    # ── v6: ML predictor + White label ────────────────────────────────────
    (6, "ML predictor y white label", [
        """CREATE TABLE IF NOT EXISTS ml_predicciones (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo_base  TEXT,
            periodo_predicho TEXT,
            descripcion   TEXT,
            valor_estimado REAL,
            confianza     REAL,
            confirmado    INTEGER DEFAULT 0,
            fecha_prediccion TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS configuracion_empresa (
            clave         TEXT PRIMARY KEY,
            valor         TEXT,
            tipo          TEXT DEFAULT 'texto',
            descripcion   TEXT,
            modificado_por TEXT,
            fecha_modificacion TEXT
        )""",
        # Configuración inicial empresa
        """INSERT OR IGNORE INTO configuracion_empresa (clave, valor, tipo, descripcion)
           VALUES
           ('empresa_nombre',         'CREDIEXPRESS POPAYÁN SAS', 'texto',  'Nombre empresa'),
           ('empresa_nit',            '900000000-0',              'texto',  'NIT empresa'),
           ('empresa_ciudad',         'Popayán',                  'texto',  'Ciudad'),
           ('empresa_color_primario', '#1F4E79',                  'color',  'Color corporativo primario'),
           ('empresa_color_secundario', '#C9A227',                'color',  'Color acento dorado'),
           ('plan_activo',            'profesional',              'texto',  'Plan: starter/profesional/empresarial'),
           ('tema_default',           'oscuro',                   'texto',  'Tema UI: claro/oscuro'),
           ('notif_email_activo',     '0',                        'bool',   'Notificaciones email activas'),
           ('notif_whatsapp_activo',  '0',                        'bool',   'Notificaciones WhatsApp activas'),
           ('backup_automatico',      '1',                        'bool',   'Backup automático activado'),
           ('backup_carpeta',         './backups',                'ruta',   'Carpeta backup local'),
           ('max_bancos',             '10',                       'numero', 'Máximo bancos simultáneos'),
           ('empresa_rep_legal',      '',                         'texto',  'Nombre Representante Legal'),
           ('empresa_rep_legal_cc',   '',                         'texto',  'C.C. Representante Legal'),
           ('empresa_contador',       '',                         'texto',  'Nombre Contador Público'),
           ('empresa_tp_contador',    '',                         'texto',  'Tarjeta Profesional Contador'),
           ('empresa_cuenta_bancaria','',                         'texto',  'Cuenta bancaria principal (ej: 1105-01 Caja)'),
           ('empresa_banco_principal','',                         'texto',  'Banco principal (ej: Bancolombia, Davivienda)')
        """,
    ]),
    # ── v7: Aprendizaje de formatos — lookup rapido por nombre de archivo ─────
    (7, "Aprendizaje formato rapido por nombre base", [
        # Agregar columna nombre_base (nombre de archivo sin ruta) y nombre_formato
        # (el 'nombre' del entry ganador en REGISTRO_FORMATOS) a pdf_formatos.
        # ALTER TABLE en SQLite no permite ADD COLUMN IF NOT EXISTS, se usa try/except
        # en el MigrationManager. Aqui va el SQL; el manager lo ejecuta.
        "ALTER TABLE pdf_formatos ADD COLUMN nombre_base    TEXT DEFAULT ''",
        "ALTER TABLE pdf_formatos ADD COLUMN nombre_formato TEXT DEFAULT ''",
        "CREATE INDEX IF NOT EXISTS idx_pdf_formatos_nombre ON pdf_formatos(nombre_base, nombre_formato)",
    ]),
]


class MigrationManager:
    """Aplica migraciones pendientes al esquema SQLite."""

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn
        self._ensure_version_table()

    def _ensure_version_table(self):
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS schema_version (
                version     INTEGER PRIMARY KEY,
                descripcion TEXT,
                aplicada_en TEXT DEFAULT (datetime('now'))
            )
        """)
        self.conn.commit()

    def _version_actual(self) -> int:
        row = self.conn.execute(
            "SELECT MAX(version) FROM schema_version"
        ).fetchone()
        return row[0] or 0

    def apply_pending(self):
        actual = self._version_actual()
        for version, descripcion, sentencias in _MIGRATIONS:
            if version <= actual:
                continue
            log.info("[migrations] Aplicando v%d: %s", version, descripcion)
            try:
                for sql in sentencias:
                    try:
                        self.conn.execute(sql)
                    except Exception as sql_err:
                        msg = str(sql_err).lower()
                        if 'duplicate column' in msg or 'already exists' in msg:
                            log.debug("[migrations] v%d: columna ya existe (OK): %s", version, sql_err)
                        else:
                            raise
                self.conn.execute(
                    "INSERT INTO schema_version (version, descripcion) VALUES (?, ?)",
                    (version, descripcion)
                )
                self.conn.commit()
                log.info("[migrations] v%d aplicada OK", version)
            except Exception as e:
                self.conn.rollback()
                log.error("[migrations] Error en v%d: %s", version, e, exc_info=True)
                raise
