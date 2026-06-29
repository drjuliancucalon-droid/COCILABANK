"""
Cola de notificaciones offline-first — CREDIEXPRESS POPAYÁN SAS
Guarda en SQLite, intenta enviar cuando hay conexión.
Canales: EMAIL, WHATSAPP (via Twilio/API), LOG (siempre disponible).
"""
import logging
import smtplib
import socket
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional

log = logging.getLogger(__name__)

_MAX_INTENTOS = 3


def _ahora() -> str:
    return datetime.now().isoformat(timespec='seconds')


def _get_db():
    from storage.db import _init_db
    return _init_db()


def _hay_internet(host='8.8.8.8', port=53, timeout=2) -> bool:
    """Verificación rápida de conectividad."""
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except (socket.error, OSError):
        return False


# ── Encolar ───────────────────────────────────────────────────────────────────
def encolar_notificacion(tipo: str, canal: str, destinatario: str,
                         asunto: str, cuerpo: str) -> int:
    """
    Agrega una notificación a la cola.
    
    tipo: CONCILIACION_LISTA | PARTIDA_PENDIENTE | ALERTA_COMISION | VENCIMIENTO_DIAN
    canal: EMAIL | WHATSAPP | LOG
    
    Retorna el ID de la notificación en cola, o -1 si falla.
    """
    try:
        conn = _get_db()
        cur = conn.execute(
            """INSERT INTO notificaciones_queue
               (tipo, canal, destinatario, asunto, cuerpo, fecha_creacion)
               VALUES (?,?,?,?,?,?)""",
            (tipo, canal.upper(), destinatario, asunto, cuerpo, _ahora())
        )
        notif_id = cur.lastrowid
        conn.commit(); conn.close()

        # Si es LOG, marcar inmediatamente
        if canal.upper() == 'LOG':
            log.info("[notif] %s → %s: %s", tipo, destinatario, asunto)
            _marcar_enviada(notif_id)

        log.debug("[notif] Encolada #%d (%s via %s)", notif_id, tipo, canal)
        return notif_id
    except Exception as e:
        log.error("[notif] Error encolando: %s", e, exc_info=True)
        return -1


def encolar_backup(archivo_local: str, destino: str = 'AUTO',
                   tipo: str = 'LOCAL') -> int:
    """Encola un archivo para backup (local o Google Drive cuando haya internet)."""
    try:
        conn = _get_db()
        cur = conn.execute(
            """INSERT INTO backup_queue
               (archivo_local, destino, tipo, fecha_creacion)
               VALUES (?,?,?,?)""",
            (archivo_local, destino, tipo, _ahora())
        )
        bid = cur.lastrowid
        conn.commit(); conn.close()
        log.debug("[backup] Encolado #%d: %s", bid, archivo_local)
        return bid
    except Exception as e:
        log.error("[notif] Error encolando backup: %s", e)
        return -1


# ── Procesamiento ─────────────────────────────────────────────────────────────
def procesar_cola_notificaciones(config: dict = None) -> dict:
    """
    Intenta enviar notificaciones pendientes.
    Retorna dict con {enviadas, fallidas, omitidas}.
    config: {email_host, email_port, email_user, email_pass,
             twilio_sid, twilio_token, twilio_from}
    """
    config = config or {}
    enviadas = fallidas = omitidas = 0

    if not _hay_internet():
        log.info("[notif] Sin internet — notificaciones quedan en cola")
        return {'enviadas': 0, 'fallidas': 0, 'omitidas': 0, 'sin_internet': True}

    try:
        conn = _get_db()
        pendientes = conn.execute(
            """SELECT id, tipo, canal, destinatario, asunto, cuerpo, intentos
               FROM notificaciones_queue
               WHERE estado='PENDIENTE' AND intentos < ?
               ORDER BY id LIMIT 20""",
            (_MAX_INTENTOS,)
        ).fetchall()
        conn.close()
    except Exception as e:
        log.error("[notif] Error leyendo cola: %s", e)
        return {'enviadas': 0, 'fallidas': 0, 'omitidas': 0}

    for nid, tipo, canal, dest, asunto, cuerpo, intentos in pendientes:
        if canal == 'LOG':
            _marcar_enviada(nid)
            enviadas += 1
            continue

        if canal == 'EMAIL':
            ok, err = _enviar_email(dest, asunto, cuerpo, config)
        elif canal == 'WHATSAPP':
            ok, err = _enviar_whatsapp(dest, cuerpo, config)
        else:
            omitidas += 1
            continue

        if ok:
            _marcar_enviada(nid)
            enviadas += 1
        else:
            _marcar_fallo(nid, err)
            fallidas += 1

    return {'enviadas': enviadas, 'fallidas': fallidas, 'omitidas': omitidas}


def procesar_cola_backup(carpeta_backup: str = './backups') -> dict:
    """
    Copia archivos pendientes de backup a la carpeta local.
    Si hay internet y el destino es GOOGLE_DRIVE, intenta subir.
    """
    import os, shutil
    os.makedirs(carpeta_backup, exist_ok=True)
    copiados = fallidos = 0

    try:
        conn = _get_db()
        pendientes = conn.execute(
            "SELECT id, archivo_local, destino, tipo FROM backup_queue "
            "WHERE estado='PENDIENTE' AND intentos < ? ORDER BY id",
            (_MAX_INTENTOS,)
        ).fetchall()
        conn.close()
    except Exception as e:
        log.error("[backup] Error leyendo cola: %s", e)
        return {'copiados': 0, 'fallidos': 0}

    for bid, archivo, destino, tipo in pendientes:
        try:
            if not os.path.exists(archivo):
                _marcar_fallo_backup(bid, "Archivo no encontrado")
                fallidos += 1
                continue

            # Siempre hacer copia local
            nombre = os.path.basename(archivo)
            destino_local = os.path.join(carpeta_backup, nombre)
            shutil.copy2(archivo, destino_local)

            # Intentar Google Drive si hay internet
            if tipo == 'GOOGLE_DRIVE' and _hay_internet():
                ok = _subir_google_drive(archivo, destino)
                if ok:
                    _marcar_enviada_backup(bid)
                    copiados += 1
                    continue

            _marcar_enviada_backup(bid)
            copiados += 1
            log.info("[backup] Copiado: %s → %s", nombre, carpeta_backup)

        except Exception as e:
            log.error("[backup] Error copiando %s: %s", archivo, e)
            _marcar_fallo_backup(bid, str(e))
            fallidos += 1

    return {'copiados': copiados, 'fallidos': fallidos}


# ── Envíos ────────────────────────────────────────────────────────────────────
def _enviar_email(destinatario: str, asunto: str, cuerpo: str,
                  config: dict) -> tuple:
    """Envía email via SMTP."""
    host  = config.get('email_host', '')
    port  = int(config.get('email_port', 587))
    user  = config.get('email_user', '')
    pwd   = config.get('email_pass', '')
    if not all([host, user, pwd]):
        return False, "Config SMTP incompleta"
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = asunto
        msg['From']    = user
        msg['To']      = destinatario
        msg.attach(MIMEText(cuerpo, 'html', 'utf-8'))
        with smtplib.SMTP(host, port, timeout=10) as s:
            s.starttls()
            s.login(user, pwd)
            s.sendmail(user, [destinatario], msg.as_string())
        log.info("[notif] Email enviado a %s: %s", destinatario, asunto)
        return True, ''
    except Exception as e:
        return False, str(e)


def _enviar_whatsapp(numero: str, mensaje: str, config: dict) -> tuple:
    """Envía WhatsApp via Twilio API."""
    sid    = config.get('twilio_sid', '')
    token  = config.get('twilio_token', '')
    desde  = config.get('twilio_from', '')
    if not all([sid, token, desde]):
        return False, "Config Twilio incompleta"
    try:
        import urllib.request, urllib.parse, base64, json
        url  = f"https://api.twilio.com/2010-04-01/Accounts/{sid}/Messages.json"
        data = urllib.parse.urlencode({
            'From': f"whatsapp:{desde}",
            'To':   f"whatsapp:{numero}",
            'Body': mensaje[:1600],
        }).encode()
        creds = base64.b64encode(f"{sid}:{token}".encode()).decode()
        req = urllib.request.Request(url, data=data, method='POST')
        req.add_header('Authorization', f'Basic {creds}')
        resp = urllib.request.urlopen(req, timeout=10)
        result = json.loads(resp.read())
        log.info("[notif] WhatsApp enviado a %s: %s", numero, result.get('sid'))
        return True, ''
    except Exception as e:
        return False, str(e)


def _subir_google_drive(archivo: str, carpeta_id: str) -> bool:
    """Sube un archivo a Google Drive (requiere credenciales en secrets)."""
    try:
        import streamlit as st
        import json as _json
        creds_json = st.secrets.get("GOOGLE_SHEETS_CREDS", None)
        if not creds_json:
            return False
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        creds = Credentials.from_service_account_info(
            _json.loads(creds_json),
            scopes=['https://www.googleapis.com/auth/drive']
        )
        service = build('drive', 'v3', credentials=creds)
        import os
        meta = {'name': os.path.basename(archivo)}
        if carpeta_id and carpeta_id != 'AUTO':
            meta['parents'] = [carpeta_id]
        media = MediaFileUpload(archivo)
        service.files().create(body=meta, media_body=media).execute()
        return True
    except Exception as e:
        log.warning("[backup] Error subiendo a Drive: %s", e)
        return False


# ── Helpers de estado ─────────────────────────────────────────────────────────
def _marcar_enviada(nid: int):
    try:
        conn = _get_db()
        conn.execute(
            "UPDATE notificaciones_queue SET estado='ENVIADA', fecha_envio=? WHERE id=?",
            (_ahora(), nid))
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[notif] Error marcando enviada: %s", e)


def _marcar_fallo(nid: int, error: str):
    try:
        conn = _get_db()
        conn.execute(
            "UPDATE notificaciones_queue SET intentos=intentos+1, "
            "error_msg=?, estado=CASE WHEN intentos+1>= ? THEN 'FALLIDA' ELSE 'PENDIENTE' END "
            "WHERE id=?", (error[:200], _MAX_INTENTOS, nid))
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[notif] Error marcando fallo: %s", e)


def _marcar_enviada_backup(bid: int):
    try:
        conn = _get_db()
        conn.execute(
            "UPDATE backup_queue SET estado='COMPLETADO', fecha_subida=? WHERE id=?",
            (_ahora(), bid))
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[backup] Error marcando enviado: %s", e)


def _marcar_fallo_backup(bid: int, error: str):
    try:
        conn = _get_db()
        conn.execute(
            "UPDATE backup_queue SET intentos=intentos+1, error_msg=? WHERE id=?",
            (error[:200], bid))
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[backup] Error marcando fallo backup: %s", e)


def listar_notificaciones_pendientes() -> List[dict]:
    """Lista notificaciones pendientes."""
    try:
        conn = _get_db()
        rows = conn.execute(
            "SELECT id, tipo, canal, destinatario, asunto, intentos, fecha_creacion "
            "FROM notificaciones_queue WHERE estado='PENDIENTE' ORDER BY id"
        ).fetchall()
        conn.close()
        return [{'id': r[0], 'tipo': r[1], 'canal': r[2], 'dest': r[3],
                 'asunto': r[4], 'intentos': r[5], 'fecha': r[6]} for r in rows]
    except Exception as e:
        log.error("[notif] Error listando: %s", e)
        return []
