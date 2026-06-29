"""
Plan Único de Cuentas (PUC) colombiano — CREDIEXPRESS POPAYÁN SAS
Asignación automática de códigos contables a movimientos bancarios.
"""
import logging
import re
import sqlite3
from typing import Optional, Tuple

log = logging.getLogger(__name__)

# ── Catálogo PUC simplificado (cuentas bancarias más usadas) ─────────────────
PUC_CATALOGO = {
    # Activo — Disponible
    '111005': ('BANCOS NACIONALES',                'DEBITO',  '1110', 'Disponible'),
    '111010': ('BANCOS DEL EXTERIOR',              'DEBITO',  '1110', 'Disponible'),
    '111505': ('BANCOS - MONEDA EXTRANJERA',       'DEBITO',  '1115', 'Disponible'),
    '112005': ('FONDOS FIDUCIARIOS',               'DEBITO',  '1120', 'Disponible'),
    # Gastos bancarios
    '530505': ('COMISIONES BANCARIAS',             'DEBITO',  '5305', 'Financieros'),
    '530510': ('GASTOS 4×1000',                    'DEBITO',  '5305', 'Financieros'),
    '530515': ('INTERESES BANCARIOS',              'DEBITO',  '5305', 'Financieros'),
    '530525': ('MANTENIMIENTO CUENTA',             'DEBITO',  '5305', 'Financieros'),
    '530530': ('GMF GRAVAMEN MOVIMIENTOS',         'DEBITO',  '5305', 'Financieros'),
    # Ingresos financieros
    '421005': ('INTERESES ACTIVOS',                'CREDITO', '4210', 'Financieros'),
    '421010': ('RENDIMIENTOS FINANCIEROS',         'CREDITO', '4210', 'Financieros'),
    # Impuestos
    '236540': ('RETENCIÓN EN LA FUENTE POR PAGAR', 'CREDITO', '2365', 'Impuestos'),
    '240805': ('IVA POR PAGAR',                    'CREDITO', '2408', 'Impuestos'),
    '240850': ('ICA POR PAGAR',                    'CREDITO', '2408', 'Impuestos'),
    '135515': ('ANTICIPOS DE IMPUESTOS - RENTA',   'DEBITO',  '1355', 'Impuestos'),
    '135520': ('AUTORRETENCIÓN',                   'DEBITO',  '1355', 'Impuestos'),
    # Proveedores / Cuentas por pagar
    '220505': ('PROVEEDORES NACIONALES',           'CREDITO', '2205', 'C×Pagar'),
    '232005': ('ACREEDORES VARIOS',                'CREDITO', '2320', 'C×Pagar'),
    # Clientes / C×Cobrar
    '130505': ('CLIENTES NACIONALES',              'DEBITO',  '1305', 'C×Cobrar'),
    # Nómina
    '251005': ('SUELDOS Y SALARIOS POR PAGAR',     'CREDITO', '2510', 'Nómina'),
    '251010': ('PRESTACIONES SOCIALES',            'CREDITO', '2510', 'Nómina'),
    # Préstamos
    '211005': ('OBLIGACIONES BANCARIAS CP',        'CREDITO', '2110', 'Financiero'),
    '221005': ('OBLIGACIONES LARGO PLAZO',         'CREDITO', '2210', 'Financiero'),
    # Socios
    '320505': ('CAPITAL SUSCRITO Y PAGADO',        'CREDITO', '3205', 'Patrimonio'),
    '330505': ('RESERVA LEGAL',                    'CREDITO', '3305', 'Patrimonio'),
    # Otros
    '159905': ('OTROS ACTIVOS',                    'DEBITO',  '1599', 'Otros'),
    '249905': ('OTROS PASIVOS',                    'CREDITO', '2499', 'Otros'),
}

# ── Reglas de clasificación automática ───────────────────────────────────────
_REGLAS_PUC = [
    # Comisiones y gastos bancarios
    (re.compile(r'4\s*[×Xx]\s*1000|GMF|GRAVAMEN|TRANS[AÁ]CION FINANCIER', re.I), '530510'),
    (re.compile(r'COMISI[ÓO]N|COBRO SERVICIO|CUOTA MANEJO', re.I),               '530505'),
    (re.compile(r'INTER[EÉ]S|INTERES MORA',                 re.I),               '530515'),
    (re.compile(r'MANTENIMIENTO|COSTO MENSUAL',             re.I),               '530525'),
    # Impuestos
    (re.compile(r'RETENCI[ÓO]N EN LA FUENTE|RETEFUENTE',   re.I),               '236540'),
    (re.compile(r'IVA\b|IMPUESTO AL VALOR',                 re.I),               '240805'),
    (re.compile(r'\bICA\b|IND\.?\s*COMERC',                 re.I),               '240850'),
    (re.compile(r'AUTORRETENCION|AUTORRETENCI[ÓO]N',        re.I),               '135520'),
    # Nómina
    (re.compile(r'N[ÓO]MINA|SALARIO|SUELDO|PAGO\s+EMPLEAD', re.I),              '251005'),
    (re.compile(r'CESANT|PRIMA|VACACION|DOTACI[ÓO]N',       re.I),              '251010'),
    # Préstamos
    (re.compile(r'ABONO\s+A\s+PRÉSTAMO|PRÉSTAMO|CUOTA\s+CRÉDITO', re.I),        '211005'),
    # Proveedores
    (re.compile(r'PROVEEDOR|PAGO\s+FACTURA|PAGO\s+A\s+',   re.I),              '220505'),
    # Ingresos
    (re.compile(r'RENDIMIENTO|INTERÉS\s+GANADO|INTER[EÉ]S\s+ACRED', re.I),      '421010'),
    (re.compile(r'ABONO\s+INTERÉS|INTERESES\s+PERIOD',      re.I),              '421005'),
    # Clientes
    (re.compile(r'CONSIGNACI[ÓO]N|RECAUDO|PAGO\s+CLIENTE', re.I),              '130505'),
    # Default banco
    (re.compile(r'BANCOLOMBIA|DAVIVIENDA|BBVA|BOGOT[AÁ]|POPULAR|OCCIDENTE|AV\s+VILLAS', re.I), '111005'),
]


def clasificar_movimiento(descripcion: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Clasifica un movimiento bancario y retorna (codigo_puc, nombre_cuenta, naturaleza).
    Primero busca en la tabla aprendida (SQLite), luego en las reglas.
    """
    desc = descripcion or ''

    # 1) Buscar en tabla aprendida
    aprendida = _buscar_en_tabla(desc)
    if aprendida:
        return aprendida

    # 2) Aplicar reglas
    for patron, codigo in _REGLAS_PUC:
        if patron.search(desc):
            entry = PUC_CATALOGO.get(codigo, (codigo, 'DEBITO', '', ''))
            return codigo, entry[0], entry[1]

    # 3) Default
    return '111005', 'BANCOS NACIONALES', 'DEBITO'


def _buscar_en_tabla(descripcion: str) -> Optional[Tuple[str, str, str]]:
    """Busca clasificación PUC aprendida en SQLite."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        # Buscar coincidencia exacta primero
        row = conn.execute(
            "SELECT codigo_puc, nombre_cuenta, naturaleza FROM puc_asignaciones "
            "WHERE lower(descripcion_banco)=lower(?) ORDER BY confirmaciones DESC LIMIT 1",
            (descripcion[:100],)
        ).fetchone()
        conn.close()
        if row:
            return row[0], row[1], row[2]
    except Exception as e:
        log.error("[puc] Error buscando en tabla: %s", e)
    return None


def aprender_clasificacion(descripcion_banco: str, codigo_puc: str,
                           usuario: str = 'admin') -> bool:
    """Guarda o actualiza una clasificación PUC aprendida."""
    entry = PUC_CATALOGO.get(codigo_puc)
    if not entry:
        log.warning("[puc] Código PUC desconocido: %s", codigo_puc)
        return False
    nombre, naturaleza = entry[0], entry[1]
    try:
        from storage.db import _init_db
        conn = _init_db()
        existe = conn.execute(
            "SELECT id, confirmaciones FROM puc_asignaciones "
            "WHERE lower(descripcion_banco)=lower(?)",
            (descripcion_banco[:100],)
        ).fetchone()
        if existe:
            conn.execute(
                "UPDATE puc_asignaciones SET codigo_puc=?, nombre_cuenta=?, "
                "naturaleza=?, usuario=?, fecha=?, confirmaciones=confirmaciones+1 "
                "WHERE id=?",
                (codigo_puc, nombre, naturaleza, usuario,
                 _ahora(), existe[0])
            )
        else:
            from datetime import datetime
            conn.execute(
                """INSERT INTO puc_asignaciones
                   (descripcion_banco, codigo_puc, nombre_cuenta, naturaleza,
                    usuario, fecha) VALUES (?,?,?,?,?,?)""",
                (descripcion_banco[:100], codigo_puc, nombre, naturaleza,
                 usuario, _ahora())
            )
        conn.commit(); conn.close()
        return True
    except Exception as e:
        log.error("[puc] Error aprendiendo: %s", e, exc_info=True)
        return False


def _ahora() -> str:
    from datetime import datetime
    return datetime.now().isoformat(timespec='seconds')


def enriquecer_dataframe_con_puc(df):
    """Agrega columnas PUC a un DataFrame de movimientos bancarios."""
    import pandas as pd
    if df is None or df.empty:
        return df
    desc_col = next((c for c in ['DESCRIPCION', 'CONCEPTO', 'descripcion'] if c in df.columns), None)
    if not desc_col:
        return df
    codigos, nombres, naturalezas = [], [], []
    for _, row in df.iterrows():
        c, n, nat = clasificar_movimiento(str(row.get(desc_col, '')))
        codigos.append(c); nombres.append(n); naturalezas.append(nat)
    df = df.copy()
    df['PUC']          = codigos
    df['CUENTA_PUC']   = nombres
    df['NATURALEZA_PUC'] = naturalezas
    return df


def listar_catalogo_puc() -> list:
    """Lista el catálogo PUC completo."""
    return [
        {'codigo': k, 'nombre': v[0], 'naturaleza': v[1], 'grupo': v[2], 'clase': v[3]}
        for k, v in sorted(PUC_CATALOGO.items())
    ]
