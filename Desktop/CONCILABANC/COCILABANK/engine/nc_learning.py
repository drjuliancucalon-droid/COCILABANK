"""
engine/nc_learning.py — Sistema de aprendizaje de conceptos NC
Aprende correspondencias banco<->Nota Contable y las promueve al catálogo.
"""
import re
import json
import hashlib as _hashlib
import unicodedata as _ud2
import logging
from datetime import datetime

from config import _STOP_NC

log = logging.getLogger(__name__)

# ── Normalización y tokens ─────────────────────────────────────────────────────
def _norm_nc(s):
    return _ud2.normalize('NFKD', (s or '').lower()).encode('ascii','ignore').decode()

def _extraer_tokens_nc(texto):
    """Extrae tokens significativos (3+ chars, sin stopwords) de un concepto NC."""
    norm   = _norm_nc(texto)
    tokens = re.findall(r'[a-z0-9]{3,}', norm)
    return sorted(set(t for t in tokens if t not in _STOP_NC))

def _uuid_par_nc(banco_tokens, aux_tokens):
    """UUID determinista del par banco<->auxiliar (MD5 de tokens ordenados)."""
    key = '|'.join(sorted(banco_tokens)) + '::' + '|'.join(sorted(aux_tokens))
    return _hashlib.md5(key.encode()).hexdigest()[:16]

def _similitud_tokens_nc(t1, t2):
    """Similitud Jaccard entre dos listas de tokens."""
    s1, s2 = set(t1), set(t2)
    if not s1 or not s2:
        return 0.0
    return len(s1 & s2) / len(s1 | s2)

# ── Búsqueda en catálogo ──────────────────────────────────────────────────────
def buscar_en_catalogo_nc(banco_desc, aux_concepto, umbral=0.30):
    """
    Busca si el par banco<->NC tiene una regla aprobada.
    Devuelve (uuid, similitud) o (None, 0.0).
    """
    try:
        banco_tok = _extraer_tokens_nc(banco_desc)
        aux_tok   = _extraer_tokens_nc(aux_concepto)
        if len(banco_tok) < 1 or len(aux_tok) < 1:
            return None, 0.0
        from storage.db import _init_db
        conn  = _init_db()
        rows  = conn.execute(
            "SELECT uuid, banco_tokens, aux_tokens, confirmaciones FROM nc_catalogo "
            "WHERE nivel IN ('ALTA','MEDIA') ORDER BY confirmaciones DESC LIMIT 200"
        ).fetchall()
        conn.close()
        mejor_sim, mejor_uuid = 0.0, None
        for uuid, bt_j, at_j, _ in rows:
            bt  = json.loads(bt_j or '[]')
            at  = json.loads(at_j or '[]')
            sim = (_similitud_tokens_nc(banco_tok, bt) +
                   _similitud_tokens_nc(aux_tok,   at)) / 2
            if sim > mejor_sim and sim >= umbral:
                mejor_sim, mejor_uuid = sim, uuid
        return mejor_uuid, mejor_sim
    except Exception as e:
        log.error("[buscar_en_catalogo_nc] %s", e, exc_info=True)
        return None, 0.0

# ── Promoción de candidatos ───────────────────────────────────────────────────
def _promover_candidatos_nc(min_veces=3):
    """Promueve candidatos con suficientes confirmaciones al catálogo."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        candidatos = conn.execute(
            "SELECT uuid, banco_desc_raw, aux_concepto_raw, banco_tokens, "
            "aux_tokens, veces_visto, fecha_primera, fecha_ultima "
            "FROM nc_aprendizaje WHERE veces_visto >= ?", (min_veces,)
        ).fetchall()
        n = 0
        for uuid, bd, ac, bt, at, vv, fp, fu in candidatos:
            nivel = 'ALTA' if vv >= 5 else 'MEDIA'
            existe = conn.execute(
                "SELECT id FROM nc_catalogo WHERE uuid=?", (uuid,)).fetchone()
            if not existe:
                conn.execute("""INSERT INTO nc_catalogo
                    (uuid, banco_tokens, aux_tokens, confirmaciones, nivel,
                     aprobado_por, fecha_primera, fecha_ultima, sync_status)
                    VALUES (?,?,?,?,?,'AUTO',?,?,'PENDIENTE_SYNC')""",
                    (uuid, bt, at, vv, nivel, fp, fu))
                n += 1
            conn.execute("DELETE FROM nc_aprendizaje WHERE uuid=?", (uuid,))
        conn.commit(); conn.close()
        return n
    except Exception as e:
        log.error("[_promover_candidatos_nc] %s", e, exc_info=True)
        return 0

# ── Aprendizaje desde match confirmado ───────────────────────────────────────
def _aprender_match_nc(banco_desc, aux_doc, aux_concepto, metodo,
                       valor_banco=None, valor_aux=None):
    """
    Registra un match NC confirmado y alimenta el aprendizaje.
    Solo aprende si el documento es NC- (notas contables).
    """
    if not aux_doc or not str(aux_doc).upper().startswith('NC-'):
        return
    try:
        banco_tok = _extraer_tokens_nc(banco_desc)
        aux_tok   = _extraer_tokens_nc(aux_concepto)
        if len(banco_tok) < 2 or len(aux_tok) < 2:
            return
        uuid  = _uuid_par_nc(banco_tok, aux_tok)
        ahora = datetime.now().isoformat(timespec='seconds')
        from storage.db import _init_db
        conn  = _init_db()
        # Log histórico
        conn.execute("""INSERT INTO nc_historial_match
            (fecha, banco_desc, aux_doc, aux_concepto, metodo, valor_banco, valor_aux)
            VALUES (?,?,?,?,?,?,?)""",
            (ahora, (banco_desc or '')[:200], aux_doc,
             (aux_concepto or '')[:200], metodo, valor_banco, valor_aux))
        # Actualizar catálogo si ya existe la regla
        en_cat = conn.execute(
            "SELECT id, confirmaciones FROM nc_catalogo WHERE uuid=?", (uuid,)).fetchone()
        if en_cat:
            nuevo_nivel = 'ALTA' if en_cat[1]+1 >= 5 else 'MEDIA'
            conn.execute(
                "UPDATE nc_catalogo SET confirmaciones=?, nivel=?, "
                "fecha_ultima=?, sync_status='PENDIENTE_SYNC' WHERE uuid=?",
                (en_cat[1]+1, nuevo_nivel, ahora, uuid))
        else:
            # Actualizar o insertar en aprendizaje
            cand = conn.execute(
                "SELECT id, veces_visto FROM nc_aprendizaje WHERE uuid=?", (uuid,)).fetchone()
            if cand:
                conn.execute(
                    "UPDATE nc_aprendizaje SET veces_visto=?, fecha_ultima=? WHERE uuid=?",
                    (cand[1]+1, ahora, uuid))
            else:
                conn.execute("""INSERT INTO nc_aprendizaje
                    (uuid, banco_desc_raw, aux_concepto_raw, banco_tokens, aux_tokens,
                     veces_visto, fecha_primera, fecha_ultima)
                    VALUES (?,?,?,?,?,1,?,?)""",
                    (uuid, (banco_desc or '')[:200], (aux_concepto or '')[:200],
                     json.dumps(banco_tok), json.dumps(aux_tok), ahora, ahora))
        conn.commit(); conn.close()
        _promover_candidatos_nc()
    except Exception as e:
        log.error("[_aprender_match_nc] %s", e, exc_info=True)

# ── Listado del catálogo ──────────────────────────────────────────────────────
def listar_catalogo_nc(limite=8):
    """Devuelve las reglas del catálogo ordenadas por confirmaciones."""
    try:
        import os
        from config import DB_PATH
        if not os.path.exists(DB_PATH):
            return [], 0, 0
        from storage.db import _init_db
        conn  = _init_db()
        total = conn.execute("SELECT COUNT(*) FROM nc_catalogo").fetchone()[0]
        rows  = conn.execute(
            "SELECT uuid, banco_tokens, aux_tokens, confirmaciones, nivel, "
            "aprobado_por, fecha_ultima "
            "FROM nc_catalogo ORDER BY confirmaciones DESC LIMIT ?", (limite,)
        ).fetchall()
        pend  = conn.execute(
            "SELECT COUNT(*) FROM nc_aprendizaje").fetchone()[0]
        conn.close()
        return rows, total, pend
    except Exception as e:
        log.error("[listar_catalogo_nc] %s", e, exc_info=True)
        return [], 0, 0
