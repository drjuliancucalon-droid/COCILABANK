"""
utils/pdf_diagnostico.py - Diagnostico de legibilidad de PDFs + OCR opcional
"""
import re
import logging

import pdfplumber

log = logging.getLogger(__name__)

try:
    from pdf2image import convert_from_path
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


def ocr_pdf_page(pdf_path, page_number):
    """Devuelve el texto de una pagina especifica usando OCR."""
    if not OCR_AVAILABLE:
        return ""
    try:
        images = convert_from_path(pdf_path, first_page=page_number, last_page=page_number)
        if images:
            return pytesseract.image_to_string(images[0], lang='spa')
    except Exception as e:
        log.warning("[ocr_pdf_page] pagina %d: %s", page_number, e)
    return ""


def diagnosticar_pdf(ruta, tipo):
    resultado = {
        'archivo': ruta, 'tipo': tipo,
        'paginas_total': 0, 'paginas_con_texto': 0, 'paginas_sin_texto': 0,
        'total_chars': 0, 'total_words': 0, 'lineas_doc_encontradas': 0,
        'pct_paginas_legibles': 0.0, 'pct_estimado_datos': 0.0,
        'calidad': '', 'advertencias': [], 'ocr_usado': False
    }

    pat_doc = re.compile(
        r'(?:CON|CE|CG|NC|RE|RG|FA|RI|REC|CXP|CXC|OC|NI|AJ|TF|EG|EI|NE|CT|ND|JR)-\d+'
    )

    # Formato A: d/MM (Bancolombia clasico)
    # Formato B: dd-MM-yyyy o dd/MM/yyyy (otros bancos colombianos)
    pat_mov_banco = re.compile(
        r'(?:^\d{1,2}/\d{2}\s+|^\d{2}[-/]\d{2}[-/]\d{4}\s+)',
        re.MULTILINE
    )

    try:
        with pdfplumber.open(ruta) as pdf:
            resultado['paginas_total'] = len(pdf.pages)
            for pag in pdf.pages:
                texto = pag.extract_text() or ''
                if len(texto.strip()) > 30:
                    resultado['paginas_con_texto'] += 1
                    resultado['total_chars'] += len(texto)
                    resultado['total_words'] += len(texto.split())
                    if tipo == 'AUXILIAR':
                        resultado['lineas_doc_encontradas'] += len(pat_doc.findall(texto))
                    else:
                        resultado['lineas_doc_encontradas'] += len(pat_mov_banco.findall(texto))
                else:
                    if OCR_AVAILABLE:
                        ocr_text = ocr_pdf_page(ruta, pag.page_number)
                        if len(ocr_text.strip()) > 30:
                            resultado['paginas_con_texto'] += 1
                            resultado['total_chars'] += len(ocr_text)
                            resultado['total_words'] += len(ocr_text.split())
                            resultado['advertencias'].append(
                                'Pag. %d: texto extraido con OCR' % pag.page_number)
                            resultado['ocr_usado'] = True
                            if tipo == 'AUXILIAR':
                                resultado['lineas_doc_encontradas'] += len(pat_doc.findall(ocr_text))
                            else:
                                resultado['lineas_doc_encontradas'] += len(pat_mov_banco.findall(ocr_text))
                        else:
                            resultado['paginas_sin_texto'] += 1
                            resultado['advertencias'].append(
                                'Pag. %d: sin texto (imagen sin OCR o ilegible)' % pag.page_number)
                    else:
                        resultado['paginas_sin_texto'] += 1
                        resultado['advertencias'].append(
                            'Pag. %d: sin texto (imagen escaneada, OCR no instalado)' % pag.page_number)
    except Exception as e:
        log.error("[diagnosticar_pdf] %s: %s", ruta, e, exc_info=True)
        resultado['advertencias'].append('Error al abrir: %s' % e)
        return resultado

    n_tot = resultado['paginas_total']
    n_ok  = resultado['paginas_con_texto']
    resultado['pct_paginas_legibles'] = (n_ok / n_tot * 100) if n_tot > 0 else 0

    n_lineas = resultado['lineas_doc_encontradas']

    if n_lineas >= 5:
        # Datos reales encontrados: portadas y paginas de publicidad no penalizan.
        # Bonus progresivo hasta 5 puntos adicionales segun volumen de transacciones.
        pct_datos = min(100.0, 95.0 + min(5.0, n_lineas / 100))
    elif n_lineas > 0:
        # Pocos datos - combinar paginas legibles + bonus por datos encontrados
        pct_datos = min(100.0, resultado['pct_paginas_legibles'] * 0.98
                        + min(2.0, n_lineas / 10))
    else:
        # Sin datos de transacciones: pagina no financiera o PDF imagen sin OCR
        pct_datos = resultado['pct_paginas_legibles'] * 0.5

    resultado['pct_estimado_datos'] = round(pct_datos, 1)

    if pct_datos >= 95:
        resultado['calidad'] = '\U0001f7e2 EXCELENTE'
    elif pct_datos >= 80:
        resultado['calidad'] = '\U0001f7e1 BUENA'
    elif pct_datos >= 50:
        resultado['calidad'] = '\U0001f7e0 PARCIAL'
    else:
        resultado['calidad'] = '\U0001f534 BAJA'

    return resultado
