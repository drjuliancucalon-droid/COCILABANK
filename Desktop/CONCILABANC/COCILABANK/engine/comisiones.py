"""
Detección de comisiones bancarias — CREDIEXPRESS POPAYÁN SAS
Identifica 4×1000, comisiones de manejo, intereses y otros cobros bancarios.
"""
import logging
import re
from datetime import datetime
from typing import List, Optional, Tuple

import pandas as pd

log = logging.getLogger(__name__)

# ── Patrones de comisiones colombianas ───────────────────────────────────────
_PATRONES_COMISION = [
    # Gravamen 4×1000 (GMF)
    ('GMF',       re.compile(r'4\s*[×Xx×]\s*1000|GMF|GRAVAMEN\s+MOV|TRANSACCI[ÓO]N\s+FINANC', re.I),
                  "Gravamen a los Movimientos Financieros (4×1000)"),
    # Comisiones de manejo
    ('MANEJO',    re.compile(r'COMISI[ÓO]N|CUOTA\s+MANEJO|COSTO\s+SERV|COBRA\s+SERVICIO', re.I),
                  "Comisión de manejo / costo servicio"),
    # Intereses de sobregiro / mora
    ('INTERES',   re.compile(r'INTER[EÉ]S\s+(MORA|SOBREGIRO|VENCIDO|CORRIENTE)|CARGO\s+INTER', re.I),
                  "Intereses bancarios"),
    # Transferencias ACH
    ('ACH',       re.compile(r'\bACH\b|TRANSF\s+ELECTR|PSE\b|DÉBITO\s+ACH', re.I),
                  "Comisión transferencia ACH/PSE"),
    # Chequeras
    ('CHEQUERA',  re.compile(r'CHEQUERA|TALONARIO|LIBRO\s+CHEQUES', re.I),
                  "Chequera / talonario"),
    # Datáfonos / datafono
    ('DATAFONO',  re.compile(r'DATÁFONO|DATAFONO|TERMINAL\s+POS|COMISI[ÓO]N\s+VENTA', re.I),
                  "Comisión datáfono"),
    # Certificaciones
    ('CERTIF',    re.compile(r'CERTIFICACI[ÓO]N|EXTRACTO\s+COBRO|CONSULTA\s+CERTIF', re.I),
                  "Certificación bancaria"),
    # Seguros asociados
    ('SEGURO',    re.compile(r'SEGURO\s+(VIDA|DESEMPLEO|INCENDIO)|PRIMA\s+SEGURO', re.I),
                  "Prima seguro asociado a cuenta"),
    # Descuentos de cartera / factoring
    ('DESCUENTO', re.compile(r'DESCUENTO\s+CART|FACTORING|ENDOSO', re.I),
                  "Descuento de cartera"),
]

# Monto mínimo para considerar una comisión (evitar falsos positivos < $100)
_UMBRAL_MINIMO_COMISION = 100.0


def detectar_comisiones(df: pd.DataFrame, banco: str = '', periodo: str = '') -> List[dict]:
    """
    Detecta comisiones bancarias en un DataFrame de movimientos.
    Retorna lista de comisiones encontradas.
    """
    if df is None or df.empty:
        return []

    desc_col = next((c for c in ['DESCRIPCION', 'CONCEPTO', 'descripcion'] if c in df.columns), None)
    val_col  = next((c for c in ['VALOR', 'valor', 'MONTO'] if c in df.columns), None)
    if not desc_col:
        return []

    comisiones = []
    for _, row in df.iterrows():
        desc  = str(row.get(desc_col, ''))
        valor = float(row.get(val_col, 0) or 0) if val_col else 0.0
        fecha = str(row.get('FECHA_RAW', row.get('FECHA', '')))
        tipo  = str(row.get('TIPO', ''))

        # Las comisiones suelen ser débitos (salidas de dinero)
        if tipo not in ('DEBITO', 'CREDITO', ''):
            pass  # procesar igual, la detección es por descripción

        for tipo_com, patron, descripcion_larga in _PATRONES_COMISION:
            if patron.search(desc) and abs(valor) >= _UMBRAL_MINIMO_COMISION:
                comisiones.append({
                    'tipo_comision':    tipo_com,
                    'descripcion':      descripcion_larga,
                    'descripcion_banco': desc[:120],
                    'valor':            abs(valor),
                    'fecha_transaccion': fecha,
                    'banco':            banco,
                    'periodo':          periodo,
                })
                break  # Solo un tipo por movimiento

    log.info("[comisiones] Detectadas %d comisiones en %s", len(comisiones), banco or 'banco')
    return comisiones


def guardar_comisiones(comisiones: List[dict]) -> int:
    """Guarda comisiones detectadas en SQLite. Retorna cantidad guardada."""
    if not comisiones:
        return 0
    try:
        from storage.db import _init_db
        conn = _init_db()
        n = 0
        for c in comisiones:
            conn.execute(
                """INSERT INTO comisiones_detectadas
                   (periodo, banco, tipo_comision, descripcion, valor,
                    fecha_transaccion, fecha_deteccion)
                   VALUES (?,?,?,?,?,?,?)""",
                (c.get('periodo', ''), c.get('banco', ''),
                 c['tipo_comision'], c['descripcion_banco'],
                 c['valor'], c.get('fecha_transaccion', ''),
                 datetime.now().isoformat(timespec='seconds'))
            )
            n += 1
        conn.commit(); conn.close()
        return n
    except Exception as e:
        log.error("[comisiones] Error guardando: %s", e, exc_info=True)
        return 0


def resumen_comisiones(banco: str = None, periodo: str = None) -> dict:
    """
    Resumen estadístico de comisiones.
    Retorna dict con total por tipo, total COP, alertas.
    """
    try:
        from storage.db import _init_db
        conn = _init_db()
        sql = ("SELECT tipo_comision, COUNT(*), COALESCE(SUM(valor),0) "
               "FROM comisiones_detectadas WHERE 1=1")
        params = []
        if banco:
            sql += " AND banco LIKE ?"; params.append(f"%{banco}%")
        if periodo:
            sql += " AND periodo=?"; params.append(periodo)
        sql += " GROUP BY tipo_comision"
        rows = conn.execute(sql, params).fetchall()

        total_all = conn.execute(
            "SELECT COALESCE(SUM(valor),0) FROM comisiones_detectadas", []
        ).fetchone()[0]
        conn.close()

        por_tipo = {r[0]: {'count': r[1], 'total': r[2]} for r in rows}
        total_periodo = sum(v['total'] for v in por_tipo.values())

        alertas = []
        if por_tipo.get('GMF', {}).get('total', 0) > 500_000:
            alertas.append("⚠️ GMF supera $500.000 — revisar si aplica exención")
        if por_tipo.get('MANEJO', {}).get('count', 0) > 3:
            alertas.append("💡 Múltiples comisiones de manejo — considerar negociar tarifa")

        return {
            'por_tipo': por_tipo,
            'total_periodo': total_periodo,
            'total_historico': total_all,
            'alertas': alertas,
        }
    except Exception as e:
        log.error("[comisiones] Error resumen: %s", e)
        return {}


def listar_comisiones(banco: str = None, periodo: str = None,
                      limite: int = 100) -> List[dict]:
    """Lista comisiones detectadas."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        sql = ("SELECT id, periodo, banco, tipo_comision, descripcion, valor, "
               "fecha_transaccion, revisado, fecha_deteccion "
               "FROM comisiones_detectadas WHERE 1=1")
        params = []
        if banco:
            sql += " AND banco LIKE ?"; params.append(f"%{banco}%")
        if periodo:
            sql += " AND periodo=?"; params.append(periodo)
        sql += " ORDER BY id DESC LIMIT ?"
        params.append(limite)
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return [
            {'id': r[0], 'periodo': r[1], 'banco': r[2], 'tipo': r[3],
             'descripcion': r[4], 'valor': r[5], 'fecha': r[6],
             'revisado': bool(r[7]), 'detectado': r[8]}
            for r in rows
        ]
    except Exception as e:
        log.error("[comisiones] Error listando: %s", e)
        return []


def alertas_gmf(df: pd.DataFrame) -> Optional[str]:
    """Genera alerta si el GMF supera el esperado (0.4% de débitos totales)."""
    try:
        comisiones = detectar_comisiones(df)
        gmf_total = sum(c['valor'] for c in comisiones if c['tipo_comision'] == 'GMF')
        if gmf_total <= 0:
            return None
        debitos = df[df.get('TIPO', pd.Series()) == 'DEBITO']['VALOR'].sum() if 'TIPO' in df.columns else 0
        if debitos > 0:
            pct = (gmf_total / abs(debitos)) * 100
            if pct > 0.45:
                return (f"⚠️ GMF detectado: ${gmf_total:,.0f} COP "
                        f"({pct:.2f}% de débitos — esperado 0.4%)")
        return f"💸 GMF total período: ${gmf_total:,.0f} COP"
    except Exception as e:
        log.error("[comisiones] Error alertas GMF: %s", e)
        return None
