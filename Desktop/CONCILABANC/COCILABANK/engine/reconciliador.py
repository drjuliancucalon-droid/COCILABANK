"""
engine/reconciliador.py — Motor de conciliación bancaria CREDIEXPRESS
Matching inteligente en 4 fases: A) tipo documento, B) número doc,
C) tolerancia aproximada, D) agrupados N:1 + aprendizaje NC.
"""
import re
import os
import sqlite3
import json
import logging
import numpy as np
import pandas as pd
from datetime import datetime

from config import TOL_EXACTA, TOL_APROX, DB_PATH

log = logging.getLogger(__name__)

# ── Patrones ──────────────────────────────────────────────────────────────────
_PAT_PREFIJO  = re.compile(r'^([A-Z]{2,3})-', re.I)
_PAT_NUMERICO = re.compile(r'[A-Z]{2,3}-(\d+)', re.I)

# ── Helpers de score ──────────────────────────────────────────────────────────
def _prefijo_doc(doc_str):
    """Extrae prefijo: CE / CG / NC / CON / desconocido."""
    m = _PAT_PREFIJO.match(str(doc_str or ''))
    return m.group(1).upper() if m else ''

def _num_doc(doc_str):
    """Extrae la parte numérica de CE-250201 → '250201'."""
    m = _PAT_NUMERICO.match(str(doc_str or ''))
    return m.group(1) if m else ''

def score_concepto(desc_banco, concepto_aux):
    """
    Similitud rápida entre descripción bancaria y concepto auxiliar.
    Devuelve 0.0–1.0 basado en palabras comunes (sin stopwords).
    """
    STOP = {'de','la','el','en','a','y','con','por','para','del','un','una',
            'los','las','al','se','su','que','no','es','pago','transferencia'}
    def _tokens(s):
        return {w.lower() for w in re.findall(r'[a-z0-9]{3,}', (s or '').lower())
                if w.lower() not in STOP}
    t1 = _tokens(desc_banco)
    t2 = _tokens(concepto_aux)
    if not t1 or not t2:
        return 0.0
    return len(t1 & t2) / max(len(t1), len(t2))

# Alias interno usado en comparar_documentos
_score_concepto = score_concepto

def comparar_documentos(df_b, df_a):
    """
    Reconciliación con matching inteligente por tipo de documento.

    Fase A — Restricción por prefijo:
        Movimiento ABONO  (vb > 0) → solo candidatos CG- / CON- (débitos)
        Movimiento CARGO  (vb < 0) → solo candidatos CE- / NC-   (créditos)
        Sin prefijo conocido        → candidatos libres de ambas columnas

    Fase B — Bonus por número de documento:
        Si el número del doc auxiliar (250201 de CE-250201) aparece
        en la descripción bancaria → ese candidato sube en prioridad.

    Dentro de los candidatos filtrados:
        1) Match exacto por monto + bonus de doc-num / similitud de concepto
        2) Match aproximado (±0.5 %) como fallback
        3) Sin match → '❌ SOLO EN BANCO'
    """
    if df_b.empty or df_a.empty:
        return pd.DataFrame(), df_a.copy() if not df_a.empty else pd.DataFrame()

    # ── Pre-cómputo ÚNICO antes del loop principal ───────────────────────────
    df_a = df_a.copy()
    df_a['_PREFIJO']  = df_a['DOCUMENTO'].apply(_prefijo_doc)
    df_a['_NUMERICO'] = df_a['DOCUMENTO'].apply(_num_doc)

    # Tokenizar conceptos auxiliares UNA SOLA VEZ (evita 1M+ re.findall)
    _STOP_SIM = {'de','la','el','en','a','y','con','por','para','del','un','una',
                 'los','las','al','se','su','que','no','es','pago','transferencia'}
    def _tok(s):
        return frozenset(
            w for w in re.findall(r'[a-z0-9]{3,}', (s or '').lower())
            if w not in _STOP_SIM
        )
    df_a['_CONC_TOK'] = df_a['CONCEPTO'].fillna('').apply(_tok)

    # Catálogo NC: cargar en memoria UNA vez
    _catalogo_nc_cache = []
    try:
        if os.path.exists(DB_PATH):
            _cn = sqlite3.connect(DB_PATH)
            _catalogo_nc_cache = _cn.execute(
                "SELECT banco_tokens, aux_tokens FROM nc_catalogo "
                "WHERE nivel IN ('ALTA','MEDIA') LIMIT 200"
            ).fetchall()
            _cn.close()
    except Exception:
        _catalogo_nc_cache = []

    # Pre-parsear tokens del catálogo NC (evita json.loads en el loop)
    _cat_parsed = []
    for _bt_j, _at_j in _catalogo_nc_cache:
        try:
            _cat_parsed.append((
                frozenset(json.loads(_bt_j or '[]')),
                frozenset(json.loads(_at_j or '[]'))
            ))
        except Exception as e:
            log.warning("[reconciliador] Error parseando tokens NC del catálogo: %s", e)

    def _jaccard(a, b):
        if not a or not b: return 0.0
        return len(a & b) / len(a | b)

    idx_usados = set()
    filas      = []

    # Patrón de rechazo definido ANTES del loop
    _PAT_REC_B_LOOP = re.compile(
        r'RECHAZOS?|DEBITO\s+POR\s+RECHAZO|ND\s+POR|DEVOLUCI|ANULACI|RETORNO|REVERSO|COBRO\s+INV|REINTEGRO', re.I)

    # ── PRE-PASE: reservar NC- para entradas bancarias de rechazo ────────────
    # Se procesan ANTES del loop principal para que las NC no sean consumidas
    # por otros cargos con montos similares (ej. IVA, comisiones regulares).
    _TOL_PRE_E   = 0.05   # ±5%
    _pre_e_rows  = {}     # idx_banco → (idx_aux, row_aux)

    _nc_pre = df_a[df_a['_PREFIJO'] == 'NC'].copy()
    _nc_pre_usados = set()

    for _idx_b, _row_b in df_b.iterrows():
        _vb = _row_b.get('VALOR', np.nan)
        if pd.isna(_vb) or float(_vb) >= 0:
            continue   # solo CARGO
        _desc = str(_row_b.get('DESCRIPCION', '') or '')
        if not _PAT_REC_B_LOOP.search(_desc):
            continue   # solo rechazos
        _m = abs(float(_vb))
        _cands = _nc_pre[
            ~_nc_pre.index.isin(_nc_pre_usados) &
            ((_nc_pre['CREDITO'] - _m).abs() / max(_m, 1) <= _TOL_PRE_E)
        ].copy()
        if _cands.empty:
            continue
        _cands['_dr'] = (_cands['CREDITO'] - _m).abs()
        _mejor = _cands.sort_values('_dr').iloc[0]
        _pre_e_rows[_idx_b] = (_mejor.name, _mejor)
        _nc_pre_usados.add(_mejor.name)

    # Reservar las NC emparejadas en el pre-pase: el loop principal no las usa
    idx_usados.update(_nc_pre_usados)
    # ── Fin PRE-PASE ──────────────────────────────────────────────────────────

    for idx_b, row_b in df_b.iterrows():
        vb = row_b['VALOR']
        if pd.isna(vb):
            continue

        monto_abs  = abs(vb)
        desc_banco = str(row_b.get('DESCRIPCION', '') or '')

        # Tokens de la descripción bancaria (calculados UNA vez por fila banco)
        banco_tok = _tok(desc_banco)

        # ── Fase A: filtrar candidatos por tipo ───────────────────────────
        es_abono = vb >= 0
        libres   = df_a[~df_a.index.isin(idx_usados)]

        if es_abono:
            col_buscar = 'DEBITO'
            candidatos = libres[
                libres['_PREFIJO'].isin(['CG','CON']) & libres[col_buscar].notna()
            ]
            if candidatos.empty:
                candidatos = libres[libres[col_buscar].notna()]
        else:
            col_buscar = 'CREDITO'
            candidatos = libres[
                libres['_PREFIJO'].isin(['CE','NC']) & libres[col_buscar].notna()
            ]
            if candidatos.empty:
                candidatos = libres[libres[col_buscar].notna()]

        candidatos = candidatos.copy()

        match_tipo = match_monto = match_idx = None
        match_doc  = match_conc = match_fecha_aux = ''
        match_metodo  = ''
        match_sim_val = 0.0

        if not candidatos.empty:
            candidatos['_diff'] = (candidatos[col_buscar] - monto_abs).abs()

            # ── Fase B: doc-num bonus (vectorizado) ───────────────────────
            candidatos['_doc_bonus'] = candidatos['_NUMERICO'].apply(
                lambda n: 1 if n and n in desc_banco else 0
            )

            # ── Similitud concepto con tokens PRE-computados ──────────────
            candidatos['_sim'] = candidatos['_CONC_TOK'].apply(
                lambda t: _jaccard(banco_tok, t)
            )

            # ── Bonus de proximidad de fecha (±5 días → +0.15) ───────────
            _fecha_b = row_b.get('FECHA')
            if pd.notna(_fecha_b) and 'FECHA' in candidatos.columns:
                candidatos['_fecha_bonus'] = candidatos['FECHA'].apply(
                    lambda f: 0.15 if (pd.notna(f) and
                        abs((pd.Timestamp(f) - pd.Timestamp(_fecha_b)).days) <= 10)
                    else 0.0
                )
            else:
                candidatos['_fecha_bonus'] = 0.0

            # ── Fase D: catálogo NC (solo si hay reglas y solo NC-) ───────
            candidatos['_cat_sim'] = 0.0
            if _cat_parsed:
                _nc_mask = candidatos['_PREFIJO'] == 'NC'
                if _nc_mask.any():
                    def _nc_cat_sim(aux_tok):
                        mejor = 0.0
                        for bt, at in _cat_parsed:
                            s = (_jaccard(banco_tok, bt) + _jaccard(aux_tok, at)) / 2
                            if s > mejor:
                                mejor = s
                        return mejor
                    candidatos.loc[_nc_mask, '_cat_sim'] = \
                        candidatos.loc[_nc_mask, '_CONC_TOK'].apply(_nc_cat_sim)

            # ── Score combinado ───────────────────────────────────────────────
            exactos = candidatos[candidatos['_diff'] <= TOL_EXACTA].copy()
            if not exactos.empty:
                exactos = exactos.sort_values(
                    ['_doc_bonus', '_cat_sim', '_fecha_bonus', '_sim', '_diff'],
                    ascending=[False, False, False, False, True]
                )
                mejor = exactos.iloc[0]
                match_tipo    = 'EXACTO'
                if mejor['_doc_bonus']:
                    match_metodo = 'DOC+MONTO'
                elif mejor['_cat_sim'] >= 0.30:
                    match_metodo = 'CATALOGO+MONTO'
                else:
                    match_metodo = 'MONTO'
                match_monto     = mejor[col_buscar]
                match_idx       = mejor.name
                match_doc       = mejor['DOCUMENTO']
                match_conc      = mejor['CONCEPTO']
                match_fecha_aux = mejor['FECHA_RAW']
                match_sim_val   = float(mejor.get('_sim', 0.0))

            # ── Fallback: match aproximado ────────────────────────────────────
            if match_tipo is None and monto_abs > 0:
                aprox = candidatos[
                    (candidatos['_diff'] / monto_abs) <= TOL_APROX
                ].copy()
                if not aprox.empty:
                    aprox = aprox.sort_values(
                        ['_doc_bonus', '_cat_sim', '_fecha_bonus', '_sim', '_diff'],
                        ascending=[False, False, False, False, True]
                    )
                    mejor = aprox.iloc[0]
                    match_tipo    = 'APROX'
                    if mejor['_doc_bonus']:
                        match_metodo = 'DOC+APROX'
                    elif mejor['_cat_sim'] >= 0.30:
                        match_metodo = 'CATALOGO+APROX'
                    else:
                        match_metodo = 'APROX'
                    match_monto     = mejor[col_buscar]
                    match_idx       = mejor.name
                    match_doc       = mejor['DOCUMENTO']
                    match_conc      = mejor['CONCEPTO']
                    match_fecha_aux = mejor['FECHA_RAW']
                    match_sim_val   = float(mejor.get('_sim', 0.0))

        if match_idx is not None:
            idx_usados.add(match_idx)

        estado   = ('✅ COINCIDE EXACTO' if match_tipo == 'EXACTO'
                    else '🔶 COINCIDE APROX.' if match_tipo == 'APROX'
                    else '❌ SOLO EN BANCO')
        diff_val = abs(monto_abs - match_monto) if match_monto is not None else None

        # Calcular confianza del match
        _confianza = {
            'DOC+MONTO'      : 95,
            'CATALOGO+MONTO' : 85,
            'MONTO'          : max(60, 60 + int(match_sim_val * 25)),
            'DOC+APROX'      : 75,
            'CATALOGO+APROX' : 60,
            'APROX'          : max(40, 40 + int(match_sim_val * 20)),
        }.get(match_metodo, 0)

        filas.append({
            'N'              : idx_b,
            'FECHA_BANCO'    : row_b['FECHA_RAW'],
            'TIPO_MOV'       : row_b['TIPO'],
            'DESCRIPCION'    : desc_banco,
            'VALOR_BANCO'    : vb,
            'DOC_AUXILIAR'   : match_doc,
            'FECHA_AUXILIAR' : match_fecha_aux,
            'CONCEPTO_AUX'   : match_conc,
            'MONTO_AUXILIAR' : match_monto,
            'DIFERENCIA'     : diff_val,
            'ESTADO'         : estado,
            'MATCH_TIPO'     : match_tipo or 'SIN_MATCH',
            'METODO_MATCH'   : match_metodo,
            'CONFIANZA'      : _confianza,
            'PAGINA_PDF'     : row_b.get('PAGINA', ''),
        })

    df_comp = pd.DataFrame(filas)

    # ── Aplicar resultados del PRE-PASE a df_comp ────────────────────────────
    # Los banco entries de rechazo quedaron como SOLO EN BANCO en el loop
    # (sus NC estaban reservadas). Aquí los marcamos como RECHAZO-CONFIRMAR.
    if _pre_e_rows and not df_comp.empty:
        for _idx_b_pre, (_idx_a_pre, _row_a_pre) in _pre_e_rows.items():
            _mask = (df_comp['N'] == _idx_b_pre) & df_comp['ESTADO'].str.contains('SOLO EN BANCO', na=False)
            if _mask.any():
                _nc_val = float(_row_a_pre.get('CREDITO', 0))
                _bv     = abs(float(df_comp.loc[_mask, 'VALOR_BANCO'].iloc[0]))
                _diff_pct = abs(_nc_val - _bv) / max(_bv, 1)
                _conf_pre = max(45, int((1 - _diff_pct) * 90))
                df_comp.loc[_mask, 'DOC_AUXILIAR']   = _row_a_pre.get('DOCUMENTO', '')
                df_comp.loc[_mask, 'FECHA_AUXILIAR']  = _row_a_pre.get('FECHA_RAW', '')
                df_comp.loc[_mask, 'CONCEPTO_AUX']   = _row_a_pre.get('CONCEPTO', '')
                df_comp.loc[_mask, 'MONTO_AUXILIAR']  = _nc_val
                df_comp.loc[_mask, 'DIFERENCIA']      = round(abs(_bv - _nc_val), 2)
                df_comp.loc[_mask, 'ESTADO']          = '🔄 RECHAZO — CONFIRMAR'
                df_comp.loc[_mask, 'MATCH_TIPO']      = 'RECHAZO'
                df_comp.loc[_mask, 'METODO_MATCH']    = 'PRE_FASE_E'
                df_comp.loc[_mask, 'CONFIANZA']       = _conf_pre
    # ── Fin aplicación PRE-PASE ───────────────────────────────────────────────

    df_solo_aux = df_a[~df_a.index.isin(idx_usados)].copy()
    # Limpiar columnas internas del auxiliar
    for _c in ['_PREFIJO', '_NUMERICO', '_CONC_TOK']:
        if _c in df_solo_aux.columns:
            df_solo_aux.drop(columns=[_c], inplace=True)
    df_solo_aux['ESTADO'] = '📋 SOLO EN AUXILIAR'

    # ══════════════════════════════════════════════════════════════════════
    # FASE E — Segundo paso: cargos rechazados sin asiento (tolerancia ±3%)
    # Detecta cargos bancarios con keywords de rechazo/devolución y los
    # empareja con NC- del auxiliar que quedaron sin match en el loop principal.
    # Se muestra como '🔄 RECHAZO — CONFIRMAR' para revisión humana.
    # ══════════════════════════════════════════════════════════════════════
    _PAT_REC_B = re.compile(
        r'RECHAZOS?|DEBITO\s+POR|ND\s+POR|DEVOLUCI|ANULACI|RETORNO|REVERSO|COBRO\s+INV|REINTEGRO', re.I)
    _PAT_REC_A = re.compile(
        r'RECHAZOS?|DEVOLUCI|ANULACI|RETORNO|REVERSO|REINTEGRO|NOTA\s+CONT|COMISI|PAGOS\s+A', re.I)
    _TOL_RECHAZO = 0.05   # ±5 % — cubre diferencias de comisión (ej. 2682.17 vs 2659.92 = 0.84%)

    if not df_comp.empty and not df_solo_aux.empty:
        # NC- libres en el auxiliar (solo las que quedaron sin emparejar)
        _nc_libres = df_solo_aux[
            df_solo_aux.get('DOCUMENTO', pd.Series(dtype=str))
                        .str.startswith('NC-', na=False) &
            df_solo_aux['CREDITO'].notna()
        ].copy()

        _usados_fase_e = set()

        for _fi in df_comp[df_comp['ESTADO'] == '❌ SOLO EN BANCO'].index:
            _rc = df_comp.loc[_fi]
            if str(_rc.get('TIPO_MOV', '') or '') != 'CARGO':
                continue
            _desc_b = str(_rc.get('DESCRIPCION', '') or '')
            if not _PAT_REC_B.search(_desc_b):
                continue
            _monto_b = abs(float(_rc.get('VALOR_BANCO', 0) or 0))
            if _monto_b < 1:
                continue

            # NC libres con diferencia de monto dentro de la tolerancia
            _cands = _nc_libres[
                ~_nc_libres.index.isin(_usados_fase_e) &
                ((_nc_libres['CREDITO'] - _monto_b).abs() / _monto_b <= _TOL_RECHAZO)
            ].copy()
            if _cands.empty:
                continue

            # Priorizar NC que también tengan keywords de rechazo en su concepto
            _cands['_rb'] = _cands['CONCEPTO'].fillna('').apply(
                lambda _c: 2 if _PAT_REC_A.search(_c) else 0)
            _cands['_dr'] = (_cands['CREDITO'] - _monto_b).abs()
            _mejor_r = _cands.sort_values(['_rb', '_dr'],
                                          ascending=[False, True]).iloc[0]

            df_comp.loc[_fi, 'DOC_AUXILIAR']  = _mejor_r.get('DOCUMENTO', '')
            df_comp.loc[_fi, 'FECHA_AUXILIAR'] = _mejor_r.get('FECHA_RAW', '')
            df_comp.loc[_fi, 'CONCEPTO_AUX']  = _mejor_r.get('CONCEPTO', '')
            df_comp.loc[_fi, 'MONTO_AUXILIAR'] = _mejor_r['CREDITO']
            df_comp.loc[_fi, 'DIFERENCIA']    = abs(_monto_b - _mejor_r['CREDITO'])
            df_comp.loc[_fi, 'ESTADO']        = '🔄 RECHAZO — CONFIRMAR'
            df_comp.loc[_fi, 'MATCH_TIPO']    = 'RECHAZO'
            df_comp.loc[_fi, 'METODO_MATCH']  = 'FASE_E'
            df_comp.loc[_fi, 'CONFIANZA']     = 45

            _usados_fase_e.add(_mejor_r.name)

        # Quitar del auxiliar suelto las NC que Fase E emparejó
        if _usados_fase_e:
            df_solo_aux = df_solo_aux.drop(index=list(_usados_fase_e), errors='ignore')

    # ══════════════════════════════════════════════════════════════════════
    # FASE F — N cargos bancarios → 1 NC (matching por agrupación)
    # Cuando el banco cobra N veces el mismo tipo de cargo (ej: IVA por
    # cada transacción) y el auxiliar tiene UNA sola NC por el total.
    # Tolerancia ±1 % para cubrir redondeos del contador.
    # ══════════════════════════════════════════════════════════════════════
    _TOL_GRUPO = 0.01   # ±1 %

    if not df_comp.empty and not df_solo_aux.empty:
        # Cargos bancarios que siguen SOLO EN BANCO
        _sb_f = df_comp[
            (df_comp['ESTADO'] == '❌ SOLO EN BANCO') &
            (df_comp['TIPO_MOV'] == 'CARGO')
        ].copy()

        # NC libres en el auxiliar
        _nc_f = df_solo_aux[
            df_solo_aux.get('DOCUMENTO', pd.Series(dtype=str))
                        .str.startswith('NC-', na=False) &
            df_solo_aux['CREDITO'].notna()
        ].copy()

        if not _sb_f.empty and not _nc_f.empty:
            # Clave de agrupación: tokens significativos de la descripción
            def _clave_grupo(s):
                words = re.findall(r'[A-Z]{3,}', (s or '').upper())
                # quitar palabras muy genéricas
                _skip = {'CARG','CARGO','PAGO','PROV','BANC','COBR'}
                return '|'.join(w for w in words if w not in _skip)

            _sb_f['_gkey'] = _sb_f['DESCRIPCION'].apply(_clave_grupo)
            # Subgrupo por monto: bucket de ±5% para separar montos distintos
            # con la misma descripción (ej. 509.62 vs 2682.17 en DEBITO RECHAZOS)
            _sb_f['_gkey_monto'] = _sb_f.apply(
                lambda r: _clave_grupo(str(r.get('DESCRIPCION',''))) + '||' +
                          str(round(abs(float(r.get('VALOR_BANCO', 0) or 0)) / 50) * 50),
                axis=1
            )

            _usados_f_b  = set()
            _usados_f_nc = set()

            # Para cada NC libre, buscar un grupo cuya suma coincida
            for _nci, _ncrow in _nc_f.iterrows():
                if _nci in _usados_f_nc:
                    continue
                _nc_val = float(_ncrow['CREDITO'])
                if _nc_val < 1:
                    continue

                _pendientes = _sb_f[~_sb_f.index.isin(_usados_f_b)]
                if _pendientes.empty:
                    break

                _mejor_grupo_idx  = None
                _mejor_grupo_diff = None

                # Primero intenta subgrupos por descripción+monto (más preciso)
                # Luego intenta solo por descripción (más amplio)
                for _gfield in ['_gkey_monto', '_gkey']:
                    for _gkey, _grp in _pendientes.groupby(_gfield):
                        if not _gkey:
                            continue
                        _suma = float(_grp['VALOR_BANCO'].abs().sum())
                        if _suma < 1:
                            continue
                        _diff_pct = abs(_suma - _nc_val) / max(_nc_val, 1)
                        if _diff_pct <= _TOL_GRUPO:
                            # Preferir el grupo cuya suma sea más cercana
                            if _mejor_grupo_diff is None or _diff_pct < _mejor_grupo_diff:
                                _mejor_grupo_idx  = list(_grp.index)
                                _mejor_grupo_diff = _diff_pct
                    if _mejor_grupo_idx is not None:
                        break  # encontró match con subgrupo más preciso

                if _mejor_grupo_idx is None:
                    continue

                _n = len(_mejor_grupo_idx)
                _conf_f = max(55, int((1 - _mejor_grupo_diff) * 85))
                for _bidx in _mejor_grupo_idx:
                    df_comp.loc[_bidx, 'DOC_AUXILIAR']  = _ncrow.get('DOCUMENTO', '')
                    df_comp.loc[_bidx, 'FECHA_AUXILIAR'] = _ncrow.get('FECHA_RAW', '')
                    df_comp.loc[_bidx, 'CONCEPTO_AUX']  = _ncrow.get('CONCEPTO', '')
                    df_comp.loc[_bidx, 'MONTO_AUXILIAR'] = round(_nc_val / _n, 2)
                    df_comp.loc[_bidx, 'DIFERENCIA']    = round(
                        abs(float(df_comp.loc[_bidx,'VALOR_BANCO']) + _nc_val/_n), 2)
                    df_comp.loc[_bidx, 'ESTADO']        = f'🔵 AGRUPADO N:1 ({_n} cargos → 1 NC)'
                    df_comp.loc[_bidx, 'MATCH_TIPO']    = 'AGRUPADO'
                    df_comp.loc[_bidx, 'METODO_MATCH']  = f'FASE_F_N{_n}'
                    df_comp.loc[_bidx, 'CONFIANZA']     = _conf_f
                    _usados_f_b.add(_bidx)
                _usados_f_nc.add(_nci)

        # Quitar NC usadas en Fase F del auxiliar suelto
        if not _sb_f.empty and '_usados_f_nc' in dir() and _usados_f_nc:
            df_solo_aux = df_solo_aux.drop(index=list(_usados_f_nc), errors='ignore')

    # ══════════════════════════════════════════════════════════════════════
    # POST-PROCESO: candidato más cercano para entradas SOLO EN BANCO
    # Para cada movimiento sin match, busca el asiento auxiliar con monto
    # más próximo (sin importar tolerancia) y explica por qué no coincidió.
    # Estas columnas se usan en la UI y en la hoja Excel de acciones.
    # ══════════════════════════════════════════════════════════════════════
    for _col_new in ['CANDIDATO_DOC','CANDIDATO_CONCEPTO','CANDIDATO_MONTO',
                     'CANDIDATO_DIFF_PCT','RAZON_NO_MATCH']:
        if _col_new not in df_comp.columns:
            df_comp[_col_new] = ''

    _sb_mask = df_comp['ESTADO'].str.contains('SOLO EN BANCO', na=False)
    if _sb_mask.any() and not df_a.empty:
        _aux_cred = df_a[df_a.get('CREDITO', pd.Series(0, index=df_a.index)) > 0].copy()
        _aux_deb  = df_a[df_a.get('DEBITO',  pd.Series(0, index=df_a.index)) > 0].copy()

        for _si, _sr in df_comp[_sb_mask].iterrows():
            _vb = abs(float(_sr.get('VALOR_BANCO', 0) or 0))
            if _vb < 1:
                df_comp.loc[_si, 'RAZON_NO_MATCH'] = 'Monto cero o no definido'
                continue
            _es_cargo = float(_sr.get('VALOR_BANCO', 0) or 0) < 0
            _pool = _aux_cred if _es_cargo else _aux_deb
            _col_v = 'CREDITO' if _es_cargo else 'DEBITO'
            if _pool.empty:
                df_comp.loc[_si, 'RAZON_NO_MATCH'] = 'Sin asientos en auxiliar'
                continue
            _p = _pool.copy()
            _p['_da'] = (_p[_col_v] - _vb).abs()
            _p['_dp'] = _p['_da'] / max(_vb, 1) * 100
            _mc = _p.sort_values('_da').iloc[0]
            _dpct = round(float(_mc['_dp']), 1)
            df_comp.loc[_si, 'CANDIDATO_DOC']     = str(_mc.get('DOCUMENTO', ''))
            df_comp.loc[_si, 'CANDIDATO_CONCEPTO'] = str(_mc.get('CONCEPTO', ''))[:80]
            df_comp.loc[_si, 'CANDIDATO_MONTO']   = float(_mc[_col_v])
            df_comp.loc[_si, 'CANDIDATO_DIFF_PCT'] = _dpct
            # Razón detallada
            if _dpct <= 0.5:
                _r = 'Monto coincide pero concepto/doc no matchea'
            elif _dpct <= 5.0:
                _r = f'Diferencia de monto {_dpct:.1f}% (fuera de tolerancia ±0.5%)'
            elif _dpct <= 20.0:
                _r = f'Diferencia de monto significativa ({_dpct:.1f}%)'
            else:
                _r = 'Sin asiento con monto similar — registrar NC'
            df_comp.loc[_si, 'RAZON_NO_MATCH'] = _r

    return df_comp, df_solo_aux


# ── Validación aritmética post-conciliación ───────────────────────────────────
UMBRAL_DIFERENCIA_NETA = 100.0  # COP — configurable

def validar_diferencia_neta(saldo_banco, saldo_auxiliar):
    """
    Calcula y retorna la diferencia neta banco vs auxiliar.
    Retorna (diferencia_neta, alerta:bool, mensaje:str).
    """
    diferencia_neta = (saldo_banco or 0.0) - (saldo_auxiliar or 0.0)
    alerta = abs(diferencia_neta) > UMBRAL_DIFERENCIA_NETA
    if alerta:
        msg = (f"⚠️ Diferencia neta de ${diferencia_neta:,.2f} COP entre extracto bancario "
               f"y auxiliar contable. Revise asientos faltantes.")
    else:
        msg = f"\u2705 Diferencia neta: ${diferencia_neta:,.2f} COP (dentro del umbral aceptable)."
    return diferencia_neta, alerta, msg
