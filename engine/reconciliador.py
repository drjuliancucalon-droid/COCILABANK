"""
engine/reconciliador.py
Motor principal de conciliacion bancaria para COCILABANK / CREDIEXPRESS POPAYAN SAS.
Implementa el algoritmo de conciliacion segun NIC 7 y normativa colombiana DIAN.
"""
from __future__ import annotations
import logging
from dataclasses import dataclass, field
from typing import Optional
import pandas as pd

logger = logging.getLogger(__name__)

TOLERANCIA_DEFAULT = 0.01


@dataclass
class ResultadoConciliacion:
    saldo_banco: float = 0.0
    saldo_auxiliar: float = 0.0
    diferencia: float = 0.0
    conciliado: bool = False
    tasa_conciliacion: float = 0.0
    total_movimientos_banco: int = 0
    total_movimientos_auxiliar: int = 0
    movimientos_cruzados: int = 0
    partidas_abiertas_banco: list = field(default_factory=list)
    partidas_abiertas_auxiliar: list = field(default_factory=list)
    depositos_transito: list = field(default_factory=list)
    cheques_pendientes: list = field(default_factory=list)
    notas_credito: list = field(default_factory=list)
    notas_debito: list = field(default_factory=list)
    errores_banco: list = field(default_factory=list)
    errores_auxiliar: list = field(default_factory=list)
    valor_riesgo: float = 0.0
    advertencias: list = field(default_factory=list)
    metadata: dict = field(default_factory=dict)


def _normalizar_valor(val) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if ',' in s and '.' in s:
        if s.rindex(',') > s.rindex('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    s = s.replace('$', '').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        return 0.0


def _normalizar_df(df, col_fecha, col_valor, col_desc):
    df = df.copy()
    df[col_valor] = df[col_valor].apply(_normalizar_valor)
    df[col_fecha] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
    df[col_desc] = df[col_desc].fillna('').astype(str).str.upper().str.strip()
    return df


def _cruzar_movimientos(df_banco, df_aux, tolerancia=TOLERANCIA_DEFAULT,
                        col_valor_banco='VALOR', col_valor_aux='VALOR',
                        col_fecha_banco='FECHA', col_fecha_aux='FECHA',
                        col_desc_banco='DESCRIPCION', col_desc_aux='DESCRIPCION',
                        ventana_dias=5):
    cruces = []
    usados_banco = set()
    usados_aux = set()
    banco_idx = df_banco.index.tolist()
    aux_idx = df_aux.index.tolist()
    banco_vals = df_banco[col_valor_banco].to_dict()
    aux_vals = df_aux[col_valor_aux].to_dict()
    banco_fechas = df_banco[col_fecha_banco].to_dict()
    aux_fechas = df_aux[col_fecha_aux].to_dict()
    for bi in sorted(banco_idx, key=lambda i: abs(banco_vals.get(i, 0)), reverse=True):
        if bi in usados_banco:
            continue
        bv = banco_vals.get(bi, 0)
        bf = banco_fechas.get(bi)
        for ai in aux_idx:
            if ai in usados_aux:
                continue
            av = aux_vals.get(ai, 0)
            af = aux_fechas.get(ai)
            if abs(bv - av) > tolerancia:
                continue
            if pd.notna(bf) and pd.notna(af):
                if abs((bf - af).days) > ventana_dias:
                    continue
            cruces.append({'idx_banco': bi, 'idx_aux': ai, 'valor': bv})
            usados_banco.add(bi)
            usados_aux.add(ai)
            break
    return cruces, list(usados_banco), list(usados_aux)


def conciliar(df_banco, df_auxiliar, saldo_banco=None, saldo_auxiliar=None,
              col_fecha_banco='FECHA', col_valor_banco='VALOR', col_desc_banco='DESCRIPCION',
              col_fecha_aux='FECHA', col_valor_aux='VALOR', col_desc_aux='DESCRIPCION',
              tolerancia=TOLERANCIA_DEFAULT, ventana_dias=5, periodo='') -> ResultadoConciliacion:
    resultado = ResultadoConciliacion()
    resultado.metadata['periodo'] = periodo
    try:
        df_b = _normalizar_df(df_banco.copy(), col_fecha_banco, col_valor_banco, col_desc_banco)
        df_a = _normalizar_df(df_auxiliar.copy(), col_fecha_aux, col_valor_aux, col_desc_aux)
    except Exception as e:
        resultado.advertencias.append(f"Error normalizando datos: {e}")
        return resultado
    resultado.total_movimientos_banco = len(df_b)
    resultado.total_movimientos_auxiliar = len(df_a)
    resultado.saldo_banco = float(saldo_banco) if saldo_banco is not None else float(df_b[col_valor_banco].sum())
    resultado.saldo_auxiliar = float(saldo_auxiliar) if saldo_auxiliar is not None else float(df_a[col_valor_aux].sum())
    cruces, cruzados_banco, cruzados_aux = _cruzar_movimientos(
        df_b, df_a, tolerancia, col_valor_banco, col_valor_aux,
        col_fecha_banco, col_fecha_aux, col_desc_banco, col_desc_aux, ventana_dias)
    resultado.movimientos_cruzados = len(cruces)
    df_ba = df_b[~df_b.index.isin(cruzados_banco)].copy()
    df_aa = df_a[~df_a.index.isin(cruzados_aux)].copy()
    resultado.partidas_abiertas_banco = df_ba.to_dict('records')
    resultado.partidas_abiertas_auxiliar = df_aa.to_dict('records')
    for row in resultado.partidas_abiertas_banco:
        val = row.get(col_valor_banco, 0)
        desc = str(row.get(col_desc_banco, '')).upper()
        if val > 0:
            (resultado.notas_credito if any(k in desc for k in ['NOTA', 'NC', 'CREDITO'])
             else resultado.depositos_transito).append(row)
        elif val < 0:
            (resultado.notas_debito if any(k in desc for k in ['NOTA', 'ND', 'DEBITO'])
             else resultado.cheques_pendientes).append(row)
    resultado.diferencia = round(resultado.saldo_banco - resultado.saldo_auxiliar, 2)
    resultado.conciliado = abs(resultado.diferencia) <= tolerancia
    total = resultado.total_movimientos_banco + resultado.total_movimientos_auxiliar
    if total > 0:
        resultado.tasa_conciliacion = round((resultado.movimientos_cruzados * 2 / total) * 100, 2)
    resultado.valor_riesgo = sum(abs(r.get(col_valor_banco, 0)) for r in resultado.partidas_abiertas_banco)
    if not resultado.conciliado:
        resultado.advertencias.append(f"Diferencia sin conciliar: ${resultado.diferencia:,.2f} COP")
    return resultado


def resumen_como_dict(resultado):
    return {
        'saldo_banco': resultado.saldo_banco,
        'saldo_auxiliar': resultado.saldo_auxiliar,
        'diferencia': resultado.diferencia,
        'conciliado': resultado.conciliado,
        'tasa_conciliacion': resultado.tasa_conciliacion,
        'total_banco': resultado.total_movimientos_banco,
        'total_auxiliar': resultado.total_movimientos_auxiliar,
        'valor_riesgo': resultado.valor_riesgo,
        'advertencias': resultado.advertencias,
        'periodo': resultado.metadata.get('periodo', '')
    }
