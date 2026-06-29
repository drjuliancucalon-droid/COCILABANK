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


def registrar_partida(periodo_origen: str, tipo: str, descripcion: str,
                      valor: float, banco: str = '', doc_referencia: str = '',
                      usuario: str = 'admin', observaciones: str = '') -> Tuple[bool, str]:
    if tipo not in TIPOS_PARTIDA: tipo = 'OTRO'
    uuid = _uuid_partida(periodo_origen, descripcion, valor)
    try:
        conn = _get_db()
        conn.execute("INSERT OR IGNORE INTO partidas_conciliatorias(uuid,periodo_origen,tipo,descripcion,valor,banco,doc_referencia,usuario_registro,fecha_registro,observaciones)VALUES(?,?,?,?,?,?,?,?,?,?)",(uuid,periodo_origen,tipo,descripcion,valor,banco,doc_referencia,usuario,_ahora(),observaciones))
        conn.execute("INSERT INTO partidas_historial(partida_uuid,periodo,accion,usuario,detalle,fecha)VALUES(?,?,?,?,?,?)",(uuid,periodo_origen,'REGISTRO',usuario,f"Partida registrada: {descripcion} ${valor:,.0f}",_ahora()))
        conn.commit(); conn.close(); return True,uuid
    except Exception as e: return False,str(e)

def listar_partidas(estado=None,banco=None,periodo=None,limite=200):
    try:
        conn = _get_db(); sql="SELECT uuid,periodo_origen,tipo,descripcion,valor,banco,doc_referencia,estado,periodo_cierre,usuario_registro,fecha_registro,observaciones FROM partidas_conciliatorias WHERE 1=1"; params=[]
        if estado: sql+=" AND estado=?"; params.append(estado)
        if banco: sql+=" AND banco LIKE ?"; params.append(f"%{banco}%")
        if periodo: sql+=" AND periodo_origen=?"; params.append(periodo)
        sql+=" ORDER BY fecha_registro DESC LIMIT?"; params.append(limite)
        rows=conn.execute(sql,params).fetchall(); conn.close()
        return [{'uuid':r[0],'periodo_origen':r[1],'tipo':r[2],'descripcion':r[3],'valor':r[4],'banco':r[5],'doc_referencia':r[6],'estado':r[7],'periodo_cierre':r[8],'usuario':r[9],'fecha':r[10],'observaciones':r[11],'tipo_label':TIPOS_PARTIDBI.get(r[2],r[2]),'estado_label':ESTADOS.get(r[7],r[7])} for r in rows]
    except: return []

def resumen_partidas(banco=None):
    try:
        conn=_get_db(); q="SELECT estado,COUNT(*),COALESCE(SUM(valor),0) FROM partidas_conciliatorias"; params=[]
        if banco: q+=" WHERE banco LIKE ?"; params.append(f"%{banco}%")
        rows=conn.execute(q+" GROUP BY estado",params).fetchall(); conn.close()
        r={'PENDIENTE':(0,0),'CONCILIADA':(0,0),'EN_PROCESO':(0,0),'ANULADA':(0,0)}
        for e,c,t in rows: r[e]=(c,t)
        return r
    except: return {}
