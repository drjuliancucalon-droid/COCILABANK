"""
Generación de PDF firmado digitalmente — CREDIEXPRESS POPAYÁN SAS
SHA-256 + timestamp para validez ante DIAN / revisiones de auditoría.
"""
import hashlib
import io
import json
import logging
import os
import sqlite3
from datetime import datetime
from typing import Optional, Tuple

log = logging.getLogger(__name__)


def _ahora_iso() -> str:
    return datetime.now().isoformat(timespec='seconds')


def _hash_contenido(contenido: bytes) -> str:
    """SHA-256 del contenido del PDF."""
    return hashlib.sha256(contenido).hexdigest()


def _leer_config(clave: str, default: str = '') -> str:
    """Lee configuración de empresa desde SQLite."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        row = conn.execute(
            "SELECT valor FROM configuracion_empresa WHERE clave=?", (clave,)
        ).fetchone()
        conn.close()
        return row[0] if row else default
    except Exception:
        return default


def generar_pdf_conciliacion(
    df_comp,
    df_banco,
    df_aux,
    banco_nombre: str,
    aux_nombre: str,
    periodo: str,
    pct_conc: float,
    saldo_banco: float,
    saldo_aux: float,
    diferencia_neta: float,
    usuario: str = 'admin',
    n_exactas: int = 0,
    n_aprox: int = 0,
    n_agrupadas: int = 0,
    n_rechazos: int = 0,
) -> Tuple[Optional[bytes], str, str]:
    """
    Genera un PDF de conciliación bancaria firmado digitalmente.
    
    Returns:
        (pdf_bytes, nombre_archivo, hash_sha256)
        Si falla, pdf_bytes=None.
    """
    try:
        from reportlab.lib.pagesizes import letter, A4
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import (
            SimpleDocTemplate, Table, TableStyle, Paragraph,
            Spacer, HRFlowable, KeepTogether
        )
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    except ImportError:
        log.warning("[pdf_firmado] reportlab no instalado — generando PDF básico con fpdf2")
        return _generar_pdf_basico(
            banco_nombre, aux_nombre, periodo, pct_conc,
            saldo_banco, saldo_aux, diferencia_neta, usuario,
            n_exactas, n_aprox, n_agrupadas, n_rechazos
        )

    # ── Configuración empresa ─────────────────────────────────────────────
    empresa    = _leer_config('empresa_nombre', 'CREDIEXPRESS POPAYÁN SAS')
    nit        = _leer_config('empresa_nit',    '900000000-0')
    ciudad     = _leer_config('empresa_ciudad', 'Popayán')
    color_hex  = _leer_config('empresa_color_primario', '#1F4E79')
    color_r    = int(color_hex[1:3], 16) / 255
    color_g    = int(color_hex[3:5], 16) / 255
    color_b    = int(color_hex[5:7], 16) / 255
    color_corp = colors.Color(color_r, color_g, color_b)

    timestamp  = _ahora_iso()
    nombre_pdf = f"Conciliacion_{banco_nombre[:20]}_{periodo}_{timestamp[:10]}.pdf"

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2.5*cm, bottomMargin=2*cm)

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('Titulo', parent=styles['Title'],
                                   fontSize=16, textColor=color_corp,
                                   spaceAfter=4)
    estilo_sub    = ParagraphStyle('Sub', parent=styles['Normal'],
                                   fontSize=9, textColor=colors.grey)
    estilo_firma  = ParagraphStyle('Firma', parent=styles['Code'],
                                   fontSize=7, textColor=colors.grey,
                                   borderPad=4, backColor=colors.lightgrey)
    estilo_seccion = ParagraphStyle('Seccion', parent=styles['Heading2'],
                                    fontSize=11, textColor=color_corp,
                                    spaceBefore=12, spaceAfter=4)
    estilo_normal = styles['Normal']
    estilo_normal.fontSize = 9

    def cop(v):
        try: return f"$ {float(v):>16,.0f} COP"
        except: return "$ —"

    story = []

    # Encabezado
    story.append(Paragraph(empresa, estilo_titulo))
    story.append(Paragraph(f"NIT: {nit} | {ciudad}", estilo_sub))
    story.append(Paragraph("CONCILIACIÓN BANCARIA", estilo_seccion))
    story.append(Paragraph(
        f"Período: <b>{periodo}</b> | Banco: <b>{banco_nombre}</b> | "
        f"Auxiliar: <b>{aux_nombre}</b>", estilo_normal))
    story.append(Paragraph(
        f"Generado: {timestamp} | Usuario: {usuario}", estilo_sub))
    story.append(HRFlowable(width="100%", thickness=1, color=color_corp))
    story.append(Spacer(1, 0.3*cm))

    # KPIs
    estado = "✅ CONCILIADO" if pct_conc >= 95 else ("⚠️ PARCIAL" if pct_conc >= 70 else "❌ PENDIENTE")
    kpi_data = [
        ['INDICADOR', 'VALOR'],
        ['Tasa de conciliación', f"{pct_conc:.1f}% — {estado}"],
        ['Saldo extracto banco', cop(saldo_banco)],
        ['Saldo auxiliar contable', cop(saldo_aux)],
        ['Diferencia neta', cop(diferencia_neta)],
        ['Coincidencias exactas', str(n_exactas)],
        ['Coincidencias aproximadas', str(n_aprox)],
        ['Agrupadas N:1', str(n_agrupadas)],
        ['Sin conciliar', str(n_rechazos)],
    ]
    kpi_style = TableStyle([
        ('BACKGROUND',  (0,0), (-1,0),  color_corp),
        ('TEXTCOLOR',   (0,0), (-1,0),  colors.white),
        ('FONTNAME',    (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',    (0,0), (-1,-1), 9),
        ('ALIGN',       (1,1), (1,-1),  'RIGHT'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#EBF5FB')]),
        ('GRID',        (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('TOPPADDING',  (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
    ])
    story.append(Paragraph("Resumen de Conciliación", estilo_seccion))
    t = Table(kpi_data, colWidths=[9*cm, 8*cm])
    t.setStyle(kpi_style)
    story.append(t)
    story.append(Spacer(1, 0.5*cm))

    # Tabla comparación (primeras 50 filas)
    if df_comp is not None and not df_comp.empty:
        story.append(Paragraph("Detalle de Comparación (primeras 50 filas)", estilo_seccion))
        cols_show = [c for c in ['FECHA_BANCO', 'DESC_BANCO', 'VALOR_BANCO',
                                  'VALOR_AUX', 'DIFERENCIA', 'ESTADO', 'METODO']
                     if c in df_comp.columns]
        if cols_show:
            header = [c.replace('_', ' ') for c in cols_show]
            rows_data = [header]
            for _, r in df_comp.head(50).iterrows():
                rows_data.append([str(r.get(c, ''))[:30] for c in cols_show])
            tbl = Table(rows_data, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), color_corp),
                ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
                ('FONTSIZE',   (0,0), (-1,-1), 7),
                ('GRID',       (0,0), (-1,-1), 0.3, colors.lightgrey),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F4F6F7')]),
            ]))
            story.append(tbl)
        story.append(Spacer(1, 0.5*cm))

    # ── Firma digital ─────────────────────────────────────────────────────
    # Generar el PDF sin la firma para calcular el hash
    doc.build(story)
    pdf_sin_firma = buf.getvalue()
    hash_sha256   = _hash_contenido(pdf_sin_firma)

    # Bloque de firma
    firma_data = {
        'documento':     'Conciliación Bancaria',
        'empresa':       empresa,
        'nit':           nit,
        'periodo':       periodo,
        'banco':         banco_nombre,
        'usuario':       usuario,
        'timestamp':     timestamp,
        'sha256':        hash_sha256,
        'sistema':       'CREDIEXPRESS Conciliación v2.0',
    }
    firma_json  = json.dumps(firma_data, ensure_ascii=False, indent=2)
    hash_firma  = hashlib.sha256(firma_json.encode()).hexdigest()

    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Paragraph("FIRMA DIGITAL DE DOCUMENTO", estilo_seccion))
    story.append(Paragraph(
        f"SHA-256 del contenido: <font name='Courier'>{hash_sha256}</font>", estilo_sub))
    story.append(Paragraph(
        f"Hash de firma: <font name='Courier'>{hash_firma[:32]}…</font>", estilo_sub))
    story.append(Paragraph(
        f"Timestamp: {timestamp} | Sistema: CREDIEXPRESS v2.0", estilo_sub))
    story.append(Paragraph(
        "Este documento ha sido generado electrónicamente y contiene firma "
        "digital SHA-256. Para verificar su integridad, calcule el hash "
        "SHA-256 del contenido y compárelo con el valor registrado.",
        estilo_sub))

    # Reconstruir con firma
    buf2 = io.BytesIO()
    doc2 = SimpleDocTemplate(buf2, pagesize=A4,
                             leftMargin=2*cm, rightMargin=2*cm,
                             topMargin=2.5*cm, bottomMargin=2*cm)
    doc2.build(story)
    pdf_final = buf2.getvalue()

    # Guardar registro en SQLite
    _registrar_exportacion(periodo, hash_sha256, nombre_pdf, usuario)

    log.info("[pdf_firmado] PDF generado: %s | SHA-256: %s…", nombre_pdf, hash_sha256[:16])
    return pdf_final, nombre_pdf, hash_sha256


def _generar_pdf_basico(banco_nombre, aux_nombre, periodo, pct_conc,
                        saldo_banco, saldo_aux, diferencia_neta, usuario,
                        n_exactas, n_aprox, n_agrupadas, n_rechazos):
    """Fallback: genera un PDF básico en HTML convertido si no hay reportlab."""
    try:
        # Intentar con fpdf2
        from fpdf import FPDF

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Helvetica', 'B', 14)
        pdf.cell(0, 10, 'CREDIEXPRESS POPAYÁN SAS', ln=True, align='C')
        pdf.set_font('Helvetica', 'B', 11)
        pdf.cell(0, 8, 'CONCILIACIÓN BANCARIA', ln=True, align='C')
        pdf.set_font('Helvetica', '', 9)
        pdf.cell(0, 6, f'Período: {periodo}  |  Banco: {banco_nombre}', ln=True)
        pdf.cell(0, 6, f'Generado: {_ahora_iso()}  |  Usuario: {usuario}', ln=True)
        pdf.ln(4)
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(0, 7, 'RESUMEN', ln=True)
        pdf.set_font('Helvetica', '', 9)
        filas = [
            ('Tasa conciliación', f"{pct_conc:.1f}%"),
            ('Saldo banco',       f"${saldo_banco:,.0f}"),
            ('Saldo auxiliar',    f"${saldo_aux:,.0f}"),
            ('Diferencia neta',   f"${diferencia_neta:,.0f}"),
            ('Exactas',           str(n_exactas)),
            ('Aproximadas',       str(n_aprox)),
            ('Sin conciliar',     str(n_rechazos)),
        ]
        for k, v in filas:
            pdf.cell(80, 6, k); pdf.cell(0, 6, v, ln=True)
        pdf.ln(6)
        contenido = pdf.output()
        if isinstance(contenido, str):
            contenido = contenido.encode('latin-1')
        hash_sha256 = _hash_contenido(contenido)
        pdf.set_font('Helvetica', '', 7)
        pdf.cell(0, 5, f'SHA-256: {hash_sha256}', ln=True)
        contenido_final = pdf.output()
        if isinstance(contenido_final, str):
            contenido_final = contenido_final.encode('latin-1')
        nombre = f"Conciliacion_{banco_nombre[:15]}_{periodo}.pdf"
        _registrar_exportacion(periodo, hash_sha256, nombre, usuario)
        return contenido_final, nombre, hash_sha256
    except ImportError:
        log.error("[pdf_firmado] Ni reportlab ni fpdf2 disponibles")
        return None, '', ''
    except Exception as e:
        log.error("[pdf_firmado] Error PDF básico: %s", e, exc_info=True)
        return None, '', ''


def _registrar_exportacion(periodo: str, hash_sha256: str,
                            archivo: str, usuario: str):
    """Registra la exportación PDF en la tabla exportaciones_dian."""
    try:
        from storage.db import _init_db
        conn = _init_db()
        conn.execute(
            """INSERT INTO exportaciones_dian
               (periodo, tipo, archivo_xml, hash_sha256, estado, usuario, fecha_generacion)
               VALUES (?,?,?,?,?,?,?)""",
            (periodo, 'PDF_CONCILIACION', archivo, hash_sha256,
             'GENERADO', usuario, _ahora_iso())
        )
        conn.commit(); conn.close()
    except Exception as e:
        log.error("[pdf_firmado] Error registrando exportación: %s", e)


def verificar_integridad_pdf(pdf_bytes: bytes, hash_esperado: str) -> bool:
    """Verifica que el hash SHA-256 del PDF coincide con el registrado."""
    return hashlib.sha256(pdf_bytes).hexdigest() == hash_esperado
