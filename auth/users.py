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

ROLES = {
    'admin': {'permisos': ['conciliar', 'exportar', 'dian', 'usuarios', 'config', 'partidas', 'auditoria', 'backup'], 'label': 'Administrador'},
    'contador_senior': {'permisos': ['conciliar', 'exportar', 'dian', 'partidas', 'auditoria'], 'label': 'Contador Senior'},
    'auxiliar': {'permisos': ['conciliar', 'exportar'], 'label': 'Auxiliar Contable'},
}

def _hash_password(password):
    try:
        import bcrypt; return bcrypt.hashpw(password.encode(), bcrypt.gensalt(12)).decode()
    except ImportError:
        salt = secrets.token_hex(16); dk = hashlib.pbkdf2_hmac('sha256', password.encode(), salt.encode(), 310_000)
        return f"pbkdf2$sha256$310000${salt}${dk.hex()}"

def _verify_password(password, stored_hash):
    try:
        if stored_hash.startswith('$2b$') or stored_hash.startswith('$2a$'):
            import bcrypt; return bcrypt.checkpw(password.encode(), stored_hash.encode())
        elif stored_hash.startswith('pbkdf2$'):
            _, algo, iters, salt, dk_hex = stored_hash.split('$'); dk = hashlib.pbkdf2_hmac(algo.replace('sha', 'sha'), password.encode(), salt.encode(), int(iters))
            return hmac.compare_digest(dk.hex(), dk_hex)
        return False
    except Exception: return False

def _get_db():
    from storage.db import _init_db; return _init_db()

def _ahora(): return datetime.now().isoformat(timespec='seconds')

def verificar_credenciales(username, password):
    if not username or not password: return False, None
    try:
        conn = _get_db(); row = conn.execute("SELECT id,username,password_hash,rol,nombre_completo,activo FROM usuarios WHERE username=?", (username.strip().lower(),)).fetchone(); conn.close()
        if not row or not row[5]: return False, None
        if not _verify_password(password, row[2]): return False, None
        conn2 = _get_db(); conn2.execute("UPDATE usuarios SET ultimo_acceso=? WHERE username=?", (_ahora(), username)); conn2.commit(); conn2.close()
        return True, {'username': row[1], 'rol': row[3], 'nombre': row[4] or row[1], 'permisos': ROLES.get(row[3], {}).get('permisos', [])}
    except Exception as e: return False, None

def crear_usuario(username, password, rol, nombre_completo='', email='', creado_por='admin'):
    if rol not in ROLES: return False, f"Rol inválido"
    if len(password) < 8: return False, "Mínimo 8 caracteres"
    try:
        conn = _get_db(); conn.execute("INSERT INTO usuarios(username,password_hash,rol,nombre_completo,email,fecha_creacion,creado_por)VALUES(?,?,?,?,?,?,?)", (username.strip().lower(),_hash_password(password),rol,nombre_completo,email,_ahora(),creado_por)); conn.commit(); conn.close(); return True, f"Usuario '{username}' creado"
    except Exception as e: return False, str(e)

def cambiar_password(username, password_nueva, por='admin'):
    if len(password_nueva) < 8: return False, "Mínimo 8 caracteres"
    try:
        conn = _get_db(); n = conn.execute("UPDATE usuarios SET password_hash=? WHERE username=?", (_hash_password(password_nueva),username)).rowcount; conn.commit(); conn.close()
        return (True, "Actualizada") if n else (False, "No encontrado")
    except Exception as e: return False, str(e)

def listar_usuarios():
    try:
        conn = _get_db(); rows = conn.execute("SELECT username,rol,nombre_completo,email,activo,fecha_creacion,ultimo_acceso FROM usuarios ORDER BY username").fetchall(); conn.close()
        return [{"username":r[0],"rol":r[1],"nombre":r[2],"email":r[3],"activo":bool(r\4])} for r in rows]
    except: return []

def tiene_permiso(usuario, permiso): return permiso in usuario.get('permisos', [])

def registrar_auditoria(usuario, accion, modulo='', detalle='', resultado='OK'):
    try:
        conn = _get_db(); conn.execute("INSERT INTO auditoria(fecha,usuario,accion,Modulo,detalle,resultado)VALUES(?,?,?,?,?,?)", (_ahora(),usuario,accion,Modulo,detalle,resultado)); conn.commit(); conn.close()
    except: pass

def _inicializar_admin_si_necesario(password_inicial="Admin2024&"):
    try:
        conn = _get_db(); row = conn.execute("SELECT password_hash FROM usuarios WHERE username='admin'").fetchone(); conn.close()
        if row and 'placeholder' in row[0]: cambiar_password('admin', password_inicial, 'sistema')
    except: pass
