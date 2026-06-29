"""
Sistema de autenticación multi-usuario offline.
Usuarios almacenados en SQLite con contraseñas hasheadas (bcrypt).
Roles: admin | contador_senior | auxiliar
"""
import hashlib
import hmac
import logging
import os
import secrets
import sqlite3
from datetime import datetime, timedelta
from typing import Optional, Tuple

log = logging.getLogger(__name__)

# ── Roles y permisos ──────────────────────────────────────────────────────────
ROLES = {
    'admin': {
        'label': '🔑 Administrador',
        'permisos': ['conciliar', 'exportar', 'dian', 'usuarios', 'config',
                     'partidas', 'auditoria', 'backup'],
    },
    'contador_senior': {
        'label': '📊 Contador Senior',
        'permisos': ['conciliar', 'exportar', 'dian', 'partidas', 'auditoria'],
    },
    'auxiliar': {
        'label': '📋 Auxiliar Contable',
        'permisos': ['conciliar', 'exportar'],
    },
}

# ── Utilidades de hash ────────────────────────────────────────────────────────
def _hash_password(password: str) -> str:
    """Hash bcrypt-compatible usando hashlib (sin dependencia externa)."""
    try:
        import bcrypt
        return bcrypt.hashpw(password.encode(), bcrypt.gensalt(12)).decode()
    except ImportError:
        # Fallback: PBKDF2 con SHA-256 (seguro si bcrypt no está disponible)
        salt = secrets.token_hex(16)
        dk = hashlib.pbkdf2_hmac('sha256', password.encode(), salt.encode(), 310_000)
        return f"pbkdf2$sha256$310000${salt}${dk.hex()}"


def _verify_password(password: str, stored_hash: str) -> bool:
    """Verifica contraseña contra hash almacenado."""
    try:
        if stored_hash.startswith('$2b$') or stored_hash.startswith('$2a$'):
            import bcrypt
            return bcrypt.checkpw(password.encode(), stored_hash.encode())
        elif stored_hash.startswith('pbkdf2$'):
            _, algo, iters, salt, dk_hex = stored_hash.split('$')
            dk = hashlib.pbkdf2_hmac(algo.replace('sha', 'sha'), password.encode(),
                                     salt.encode(), int(iters))
            return hmac.compare_digest(dk.hex(), dk_hex)
        else:
            # Placeholder hash del usuario admin inicial — forzar cambio
            return False
    except Exception as e:
        log.error("[auth] Error verificando contraseña: %s", e)
        return False


def _get_db() -> sqlite3.Connection:
    from storage.db import _init_db
    return _init_db()


def _ahora() -> str:
    return datetime.now().isoformat(timespec='seconds')


# ── CRUD Usuarios ─────────────────────────────────────────────────────────────
def verificar_credenciales(username: str, password: str) -> Tuple[bool, Optional[dict]]:
    """
    Verifica username/password.
    Retorna (True, usuario_dict) si OK, (False, None) si falla.
    """
    if not username or not password:
        return False, None
    try:
        conn = _get_db()
        row = conn.execute(
            "SELECT id, username, password_hash, rol, nombre_completo, activo "
            "FROM usuarios WHERE username=?", (username.strip().lower(),)
        ).fetchone()
        conn.close()

        if not row:
            log.warning("[auth] Usuario no encontrado: %s", username)
            return False, None

        _, uname, pwd_hash, rol, nombre, activo = row
        if not activo:
            log.warning("[auth] Usuario inactivo: %s", username)
            return False, None

        if not _verify_password(password, pwd_hash):
            log.warning("[auth] Contraseña incorrecta para: %s", username)
            return False, None

        # Actualizar último acceso
        conn2 = _get_db()
        conn2.execute(
            "UPDATE usuarios SET ultimo_acceso=? WHERE username=?",
            (_ahora(), username)
        )
        conn2.commit(); conn2.close()

        usuario = {
            'username': uname,
            'rol': rol,
            'nombre': nombre or uname,
            'permisos': ROLES.get(rol, {}).get('permisos', []),
        }
        log.info("[auth] Login OK: %s (%s)", username, rol)
        return True, usuario

    except Exception as e:
        log.error("[auth] Error en verificar_credenciales: %s", e, exc_info=True)
        return False, None


def crear_usuario(username: str, password: str, rol: str,
                  nombre_completo: str = '', email: str = '',
                  creado_por: str = 'admin') -> Tuple[bool, str]:
    """Crea un nuevo usuario. Retorna (ok, mensaje)."""
    if rol not in ROLES:
        return False, f"Rol inválido. Opciones: {list(ROLES.keys())}"
    if len(password) < 8:
        return False, "La contraseña debe tener mínimo 8 caracteres"

    try:
        pwd_hash = _hash_password(password)
        conn = _get_db()
        conn.execute(
            """INSERT INTO usuarios
               (username, password_hash, rol, nombre_completo, email,
                fecha_creacion, creado_por)
               VALUES (?,?,?,?,?,?,?)""",
            (username.strip().lower(), pwd_hash, rol,
             nombre_completo, email, _ahora(), creado_por)
        )
        conn.commit(); conn.close()
        log.info("[auth] Usuario creado: %s (%s) por %s", username, rol, creado_por)
        return True, f"Usuario '{username}' creado exitosamente"
    except sqlite3.IntegrityError:
        return False, f"El usuario '{username}' ya existe"
    except Exception as e:
        log.error("[auth] Error creando usuario: %s", e, exc_info=True)
        return False, f"Error: {e}"


def cambiar_password(username: str, password_nueva: str,
                     por: str = 'admin') -> Tuple[bool, str]:
    """Cambia la contraseña de un usuario."""
    if len(password_nueva) < 8:
        return False, "Mínimo 8 caracteres"
    try:
        pwd_hash = _hash_password(password_nueva)
        conn = _get_db()
        n = conn.execute(
            "UPDATE usuarios SET password_hash=? WHERE username=?",
            (pwd_hash, username)
        ).rowcount
        conn.commit(); conn.close()
        if n == 0:
            return False, "Usuario no encontrado"
        log.info("[auth] Password cambiado para %s por %s", username, por)
        return True, "Contraseña actualizada"
    except Exception as e:
        log.error("[auth] Error cambiando password: %s", e, exc_info=True)
        return False, str(e)


def toggle_usuario(username: str, activo: bool) -> bool:
    """Activa o desactiva un usuario."""
    try:
        conn = _get_db()
        conn.execute(
            "UPDATE usuarios SET activo=? WHERE username=?",
            (1 if activo else 0, username)
        )
        conn.commit(); conn.close()
        return True
    except Exception as e:
        log.error("[auth] Error en toggle_usuario: %s", e)
        return False


def listar_usuarios() -> list:
    """Lista todos los usuarios (sin hash de contraseña)."""
    try:
        conn = _get_db()
        rows = conn.execute(
            "SELECT username, rol, nombre_completo, email, activo, "
            "fecha_creacion, ultimo_acceso FROM usuarios ORDER BY username"
        ).fetchall()
        conn.close()
        return [
            {'username': r[0], 'rol': r[1], 'nombre': r[2],
             'email': r[3], 'activo': bool(r[4]),
             'creado': r[5], 'ultimo_acceso': r[6]}
            for r in rows
        ]
    except Exception as e:
        log.error("[auth] Error listando usuarios: %s", e)
        return []


def get_usuario(username: str) -> Optional[dict]:
    """Obtiene datos de un usuario."""
    try:
        conn = _get_db()
        row = conn.execute(
            "SELECT username, rol, nombre_completo, email, activo "
            "FROM usuarios WHERE username=?", (username,)
        ).fetchone()
        conn.close()
        if not row:
            return None
        return {'username': row[0], 'rol': row[1], 'nombre': row[2],
                'email': row[3], 'activo': bool(row[4]),
                'permisos': ROLES.get(row[1], {}).get('permisos', [])}
    except Exception as e:
        log.error("[auth] Error en get_usuario: %s", e)
        return None


def tiene_permiso(usuario: dict, permiso: str) -> bool:
    """Verifica si un usuario tiene un permiso específico."""
    return permiso in usuario.get('permisos', [])


# ── Auditoría ─────────────────────────────────────────────────────────────────
def registrar_auditoria(usuario: str, accion: str, modulo: str = '',
                        detalle: str = '', resultado: str = 'OK'):
    """Registra una acción en la tabla de auditoría."""
    try:
        conn = _get_db()
        conn.execute(
            """INSERT INTO auditoria
               (fecha, usuario, accion, modulo, detalle, resultado)
               VALUES (?,?,?,?,?,?)""",
            (_ahora(), usuario, accion, modulo, detalle, resultado)
        )
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[auth] Error registrando auditoría: %s", e)


def listar_auditoria(limite: int = 100, usuario: str = None) -> list:
    """Lista registros de auditoría."""
    try:
        conn = _get_db()
        if usuario:
            rows = conn.execute(
                "SELECT fecha, usuario, accion, modulo, detalle, resultado "
                "FROM auditoria WHERE usuario=? ORDER BY id DESC LIMIT ?",
                (usuario, limite)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT fecha, usuario, accion, modulo, detalle, resultado "
                "FROM auditoria ORDER BY id DESC LIMIT ?", (limite,)
            ).fetchall()
        conn.close()
        return [{'fecha': r[0], 'usuario': r[1], 'accion': r[2],
                 'modulo': r[3], 'detalle': r[4], 'resultado': r[5]}
                for r in rows]
    except Exception as e:
        log.error("[auth] Error listando auditoría: %s", e)
        return []


def _inicializar_admin_si_necesario(password_inicial: str = "Admin2024*"):
    """
    Si el admin existe con hash placeholder, establece la contraseña inicial.
    Se llama una sola vez en el primer arranque.
    """
    try:
        conn = _get_db()
        row = conn.execute(
            "SELECT password_hash FROM usuarios WHERE username='admin'"
        ).fetchone()
        conn.close()
        if row and 'placeholder' in row[0]:
            ok, msg = cambiar_password('admin', password_inicial, 'sistema')
            if ok:
                log.info("[auth] Contraseña admin inicial establecida")
    except Exception as e:
        log.error("[auth] Error inicializando admin: %s", e)
