"""
Módulo de Partidas Conciliatorias — CREDIEXPRESS POPAYÁN SAS
Seguimiento de movimientos que concilian en meses posteriores.
"""
import hashlib
import logging
import sqlite3
from datetime import datetime
from typing import List, Optional, Tuple

log = logging.getLogger(__name__)

ESTADOS = {
    'PENDIENTE': '⏳ Pendiente',
    'EN_PROCESO': '🔄 En proceso',
    'CONCILIADA': '✅ Conciliada',
    'ANULADA': '❌ Anulada',
}

TIPOS_PARTIDA = {
    'DEPOSITO_TRANSITO':   '🏦 Depósito en tránsito',
    'CHEQUE_PENDIENTE':    '📄 Cheque pendiente cobro',
    'ERROR_BANCO':         '⚠️ Error del banco',
    'ERROR_AUXILIAR':      '⚠️ Error en auxiliar',
    'COMISION_NO_CONT':    '💸 Comisión no contabilizada',
    'NOTA_DEBITO':         '📉 Nota débito bancaria',
    'NOTA_CREDITO_BANCO':  '📈 Nota crédito bancaria',
    'RETENCION':           '🏛️ Retención en la fuente',
    'OTRO':                '🔹 Otro concepto',
}


def _get_db() -> sqlite3.Connection:
    from storage.db import _init_db
    return _init_db()


def _ahora() -> str:
    return datetime.now().isoformat(timespec='seconds')


def _uuid_partida(periodo: str, descripcion: str, valor: float) -> str:
    raw = f"{periodo}_{descripcion}_{valor:.2f}"
    return hashlib.md5(raw.encode()).hexdigest()[:12]


# ── CRUD Partidas ─────────────────────────────────────────────────────────────
def registrar_partida(periodo_origen: str, tipo: str, descripcion: str,
                      valor: float, banco: str = '', doc_referencia: str = '',
                      usuario: str = 'admin', observaciones: str = '') -> Tuple[bool, str]:
    """Registra una nueva partida conciliatoria."""
    if tipo not in TIPOS_PARTIDA:
        tipo = 'OTRO'
    uuid = _uuid_partida(periodo_origen, descripcion, valor)
    try:
        conn = _get_db()
        conn.execute(
            """INSERT OR IGNORE INTO partidas_conciliatorias
               (uuid, periodo_origen, tipo, descripcion, valor, banco,
                doc_referencia, usuario_registro, fecha_registro, observaciones)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (uuid, periodo_origen, tipo, descripcion, valor, banco,
             doc_referencia, usuario, _ahora(), observaciones)
        )
        # Historial
        conn.execute(
            """INSERT INTO partidas_historial
               (partida_uuid, periodo, accion, usuario, detalle, fecha)
               VALUES (?,?,?,?,?,?)""",
            (uuid, periodo_origen, 'REGISTRO', usuario,
             f"Partida registrada: {descripcion} ${valor:,.0f}", _ahora())
        )
        conn.commit(); conn.close()
        log.info("[partidas] Registrada: %s %s $%.0f", tipo, descripcion, valor)
        return True, uuid
    except Exception as e:
        log.error("[partidas] Error registrando: %s", e, exc_info=True)
        return False, str(e)


def listar_partidas(estado: str = None, banco: str = None,
                    periodo: str = None, limite: int = 200) -> List[dict]:
    """Lista partidas conciliatorias con filtros opcionales."""
    try:
        conn = _get_db()
        sql = """SELECT uuid, periodo_origen, tipo, descripcion, valor,
                        banco, doc_referencia, estado, periodo_cierre,
                        usuario_registro, fecha_registro, observaciones
                 FROM partidas_conciliatorias WHERE 1=1"""
        params = []
        if estado:
            sql += " AND estado=?"; params.append(estado)
        if banco:
            sql += " AND banco LIKE ?"; params.append(f"%{banco}%")
        if periodo:
            sql += " AND periodo_origen=?"; params.append(periodo)
        sql += " ORDER BY fecha_registro DESC LIMIT ?"
        params.append(limite)
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return [
            {'uuid': r[0], 'periodo_origen': r[1], 'tipo': r[2],
             'descripcion': r[3], 'valor': r[4], 'banco': r[5],
             'doc_referencia': r[6], 'estado': r[7], 'periodo_cierre': r[8],
             'usuario': r[9], 'fecha': r[10], 'observaciones': r[11],
             'tipo_label': TIPOS_PARTIDA.get(r[2], r[2]),
             'estado_label': ESTADOS.get(r[7], r[7])}
            for r in rows
        ]
    except Exception as e:
        log.error("[partidas] Error listando: %s", e)
        return []


def conciliar_partida(uuid: str, periodo_cierre: str,
                      usuario: str = 'admin', detalle: str = '') -> Tuple[bool, str]:
    """Marca una partida como conciliada en el período indicado."""
    try:
        conn = _get_db()
        n = conn.execute(
            """UPDATE partidas_conciliatorias
               SET estado='CONCILIADA', periodo_cierre=?, fecha_cierre=?
               WHERE uuid=? AND estado='PENDIENTE'""",
            (periodo_cierre, _ahora(), uuid)
        ).rowcount
        if n > 0:
            conn.execute(
                """INSERT INTO partidas_historial
                   (partida_uuid, periodo, accion, usuario, detalle, fecha)
                   VALUES (?,?,?,?,?,?)""",
                (uuid, periodo_cierre, 'CONCILIADA', usuario,
                 detalle or f"Conciliada en {periodo_cierre}", _ahora())
            )
        conn.commit(); conn.close()
        if n == 0:
            return False, "Partida no encontrada o ya conciliada"
        return True, "Partida conciliada"
    except Exception as e:
        log.error("[partidas] Error conciliando: %s", e, exc_info=True)
        return False, str(e)


def resumen_partidas(banco: str = None) -> dict:
    """Resumen estadístico de partidas."""
    try:
        conn = _get_db()
        q = "SELECT estado, COUNT(*), COALESCE(SUM(valor),0) FROM partidas_conciliatorias"
        params = []
        if banco:
            q += " WHERE banco LIKE ?"; params.append(f"%{banco}%")
        q += " GROUP BY estado"
        rows = conn.execute(q, params).fetchall()
        conn.close()
        resumen = {'PENDIENTE': (0, 0), 'CONCILIADA': (0, 0),
                   'EN_PROCESO': (0, 0), 'ANULADA': (0, 0)}
        for estado, cnt, total in rows:
            resumen[estado] = (cnt, total)
        return resumen
    except Exception as e:
        log.error("[partidas] Error resumen: %s", e)
        return {}


def detectar_partidas_automaticas(df_solo_banco, df_solo_aux,
                                  periodo: str, banco: str = '',
                                  usuario: str = 'sistema') -> int:
    """
    Sugiere partidas conciliatorias automáticamente desde
    movimientos sin conciliar (Solo_Banco y Solo_Auxiliar).
    """
    n = 0
    try:
        # Movimientos solo en banco → posibles depósitos en tránsito o errores
        for _, row in df_solo_banco.iterrows():
            desc  = str(row.get('DESCRIPCION', ''))
            valor = float(row.get('VALOR', 0) or 0)
            tipo  = row.get('TIPO', '')
            if abs(valor) < 1:
                continue
            t_partida = 'DEPOSITO_TRANSITO' if tipo == 'DEBITO' else 'NOTA_CREDITO_BANCO'
            ok, _ = registrar_partida(periodo, t_partida, desc, abs(valor),
                                      banco=banco, usuario=usuario,
                                      observaciones='Auto-detectada desde Solo_Banco')
            if ok:
                n += 1

        # Movimientos solo en auxiliar → cheques pendientes
        for _, row in df_solo_aux.iterrows():
            desc  = str(row.get('CONCEPTO', ''))
            valor = float(row.get('VALOR', 0) or 0)
            if abs(valor) < 1:
                continue
            ok, _ = registrar_partida(periodo, 'CHEQUE_PENDIENTE', desc, abs(valor),
                                      banco=banco, usuario=usuario,
                                      observaciones='Auto-detectada desde Solo_Auxiliar')
            if ok:
                n += 1
    except Exception as e:
        log.error("[partidas] Error auto-detectando: %s", e)
    return n
