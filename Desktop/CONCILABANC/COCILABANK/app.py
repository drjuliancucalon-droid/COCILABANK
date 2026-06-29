"""
CREDIEXPRESS POPAYÁN SAS — Conciliación Bancaria Comercial v3.0
Sistema profesional para contadores colombianos.
Offline-first · Multi-banco · Multi-usuario · DIAN · PUC · ML
"""
import hmac
import io
import json
import logging
import os
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    _PLOTLY_DISPONIBLE = True
except ImportError:
    _PLOTLY_DISPONIBLE = False
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker

log = logging.getLogger(__name__)

# ── Módulos base ───────────────────────────────────────────────────────────────
from config import BASE_DIR, OFFLINE_MODE, DB_PATH

from storage import (
    guardar_historial,
    registrar_aprendizaje_nc,
    leer_historial_sqlite,
    listar_formatos_aprendidos,
    _auto_guardar_excel,
    _auto_guardar_archivo,
)
from storage.sheets import sincronizar_catalogo_nc

from engine import comparar_documentos, validar_diferencia_neta
from engine.nc_learning import listar_catalogo_nc

from parsers import cargar_y_parsear_uploaded_file as cargar_y_parsear

from utils import cop, pct_bar, semaforo_conciliacion, OCR_AVAILABLE
from utils.formatters import (
    _cop_limpio, _semaforo_conciliacion,
    _inferir_cuenta_sugerida, _guia_banco_sin_aux, _guia_aux_sin_banco,
)
from utils.periodo import _extraer_periodo, _MESES_ES, inferir_periodo_desde_df, validar_periodo_archivos

from exports import generar_excel

# ── Módulos comerciales (sprints 1-6) ─────────────────────────────────────────
try:
    from auth.users import (
        verificar_credenciales, listar_usuarios, crear_usuario,
        cambiar_password, toggle_usuario, registrar_auditoria,
        listar_auditoria, ROLES, _inicializar_admin_si_necesario,
    )
    _AUTH_DISPONIBLE = True
except Exception as _e:
    log.warning("auth no disponible: %s", _e)
    _AUTH_DISPONIBLE = False

try:
    from engine.partidas import (
        registrar_partida, listar_partidas, conciliar_partida,
        resumen_partidas, detectar_partidas_automaticas, TIPOS_PARTIDA,
    )
    _PARTIDAS_DISPONIBLE = True
except Exception as _e:
    log.warning("partidas no disponible: %s", _e)
    _PARTIDAS_DISPONIBLE = False

try:
    from engine.puc import (
        clasificar_movimiento, enriquecer_dataframe_con_puc,
        aprender_clasificacion, listar_catalogo_puc, PUC_CATALOGO,
    )
    _PUC_DISPONIBLE = True
except Exception as _e:
    log.warning("puc no disponible: %s", _e)
    _PUC_DISPONIBLE = False

try:
    from engine.comisiones import (
        detectar_comisiones, guardar_comisiones, resumen_comisiones,
        listar_comisiones, alertas_gmf,
    )
    _COMISIONES_DISPONIBLE = True
except Exception as _e:
    log.warning("comisiones no disponible: %s", _e)
    _COMISIONES_DISPONIBLE = False

try:
    from exports.pdf_firmado import generar_pdf_conciliacion, verificar_integridad_pdf
    _PDF_FIRMADO_DISPONIBLE = True
except Exception as _e:
    log.warning("pdf_firmado no disponible: %s", _e)
    _PDF_FIRMADO_DISPONIBLE = False

try:
    from exports.dian_xml import (
        generar_xml_medios_magneticos, generar_formato_1007_retenciones,
        generar_formato_1008_pagos, listar_exportaciones_dian,
    )
    _DIAN_DISPONIBLE = True
except Exception as _e:
    log.warning("dian_xml no disponible: %s", _e)
    _DIAN_DISPONIBLE = False

try:
    from notifications.queue import (
        encolar_notificacion, procesar_cola_notificaciones,
        encolar_backup, procesar_cola_backup, listar_notificaciones_pendientes,
    )
    _NOTIF_DISPONIBLE = True
except Exception as _e:
    log.warning("notifications no disponible: %s", _e)
    _NOTIF_DISPONIBLE = False

try:
    from ml.predictor import (
        predecir_partidas_proximas, listar_predicciones,
        confirmar_prediccion, accuracy_modelo,
    )
    _ML_DISPONIBLE = True
except Exception as _e:
    log.warning("ml no disponible: %s", _e)
    _ML_DISPONIBLE = False

try:
    from utils.roi import calcular_roi, roi_acumulado_mes, calendario_fiscal_colombia
    _ROI_DISPONIBLE = True
except Exception as _e:
    log.warning("roi no disponible: %s", _e)
    _ROI_DISPONIBLE = False

try:
    import white_label as WL
    _WL_DISPONIBLE = True
except Exception as _e:
    log.warning("white_label no disponible: %s", _e)
    _WL_DISPONIBLE = False

warnings.filterwarnings('ignore')
pd.set_option('display.float_format', lambda x: f'{x:,.2f}')
pd.set_option('display.max_colwidth', 90)
pd.set_option('display.max_rows', 800)

st.set_page_config(page_title="Conciliación CREDIEXPRESS", page_icon="\U0001f3e6", layout="wide")
# ── Inicializar admin SQLite (primer arranque) ─────────────────────────────────
if _AUTH_DISPONIBLE:
    try:
        _inicializar_admin_si_necesario()
    except Exception:
        pass

# ── Configuración white label ──────────────────────────────────────────────────
_empresa    = WL.get("empresa_nombre",       "CREDIEXPRESS POPAYÁN SAS") if _WL_DISPONIBLE else "CREDIEXPRESS POPAYÁN SAS"
_color_corp = WL.get("empresa_color_primario",  "#1F4E79")               if _WL_DISPONIBLE else "#1F4E79"
_color_gold = WL.get("empresa_color_secundario","#C9A227")               if _WL_DISPONIBLE else "#C9A227"
_tema       = WL.get("tema_default", "oscuro")                           if _WL_DISPONIBLE else "oscuro"

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

_corp_r, _corp_g, _corp_b = _hex_to_rgb(_color_corp)
_gold_r, _gold_g, _gold_b = _hex_to_rgb(_color_gold)

if _tema == "claro":
    _bg_main  = "#F4F6F9"; _bg_card  = "#FFFFFF"; _bg_card2 = "#EBF5FB"
    _txt_main = "#1A1A2E"; _txt_sec  = "#555";    _borde_op = "30"
else:
    _bg_main  = "#0D1B2A"; _bg_card  = "#162032"; _bg_card2 = "#1E3A5F"
    _txt_main = "#E8F4FD"; _txt_sec  = "#90CAF9"; _borde_op = "44"

st.markdown("""
<style>
/* ══════════════════════════════════════════════════════════════
   COCILABANK — PREMIUM DARK EDITION v4.0
   Sistema de diseño completo: variables · animaciones · componentes
   ══════════════════════════════════════════════════════════════ */

/* ── Variables CSS globales ── */
:root {
  --c-blue:    #3b82f6;
  --c-blue-lg: #60a5fa;
  --c-green:   #22c55e;
  --c-gold:    #f59e0b;
  --c-red:     #ef4444;
  --c-purple:  #a855f7;
  --c-cyan:    #06b6d4;
  --c-bg:      #0d1b2a;
  --c-card:    #162032;
  --c-card2:   #1e2d42;
  --c-border:  rgba(59,130,246,0.2);
  --c-text:    #e2e8f0;
  --c-muted:   #94a3b8;
  --radius:    12px;
  --shadow-blue: 0 4px 20px rgba(59,130,246,0.25);
  --shadow-gold: 0 4px 20px rgba(245,158,11,0.25);
}

/* ── Fuente global ── */
html, body, [class*="css"] {
  font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
}

/* ══ ANIMACIONES KEYFRAMES ══ */

/* 1. Entrada suave desde abajo */
@keyframes fadeInSlide {
  from { opacity:0; transform:translateY(14px); }
  to   { opacity:1; transform:translateY(0); }
}
/* 2. Escala de entrada para badges */
@keyframes badgePop {
  0%   { opacity:0; transform:scale(0.6); }
  70%  { transform:scale(1.08); }
  100% { opacity:1; transform:scale(1); }
}
/* 3. Pulso de brillo para cards críticas */
@keyframes pulseGlow {
  0%,100% { box-shadow: 0 0 0 0 rgba(239,68,68,0); }
  50%      { box-shadow: 0 0 18px 4px rgba(239,68,68,0.35); }
}
/* 4. Pulso dorado para header */
@keyframes headerPulse {
  0%,100% { border-color: rgba(245,158,11,0.25); }
  50%      { border-color: rgba(245,158,11,0.70); }
}
/* 5. Shimmer de carga en botones */
@keyframes shimmer {
  0%   { background-position: -400px 0; }
  100% { background-position:  400px 0; }
}
/* 6. Ripple al hacer click */
@keyframes rippleEffect {
  to { transform:scale(4); opacity:0; }
}
/* 7. Barra de progreso fill */
@keyframes fillBar {
  from { width:0%; }
  to   { width:var(--fill); }
}
/* 8. Toast slide-in */
@keyframes slideInRight {
  from { transform:translateX(110%); opacity:0; }
  to   { transform:translateX(0);    opacity:1; }
}
@keyframes fadeOutToast {
  from { opacity:1; }
  to   { opacity:0; }
}

/* ══ HEADER GLASSMORPHISM PREMIUM ══ */
.main-header {
  background: linear-gradient(135deg,
    rgba(15,40,90,0.95) 0%,
    rgba(21,60,130,0.92) 50%,
    rgba(30,80,160,0.90) 100%);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  color: #ffffff;
  padding: 28px 36px;
  border-radius: 16px;
  margin-bottom: 20px;
  border: 1px solid rgba(245,158,11,0.30);
  box-shadow: 0 8px 32px rgba(13,71,161,0.40),
              inset 0 1px 0 rgba(255,255,255,0.08);
  animation: fadeInSlide 0.5s ease both,
             headerPulse 4s ease-in-out 1s infinite;
  position: relative;
  overflow: hidden;
}
.main-header::before {
  content:'';
  position:absolute; top:0; left:0; right:0; height:2px;
  background: linear-gradient(90deg, transparent, #f59e0b, #60a5fa, transparent);
}
.main-header h1 {
  color: #ffffff !important;
  margin: 0;
  font-size: 1.70rem;
  font-weight: 900;
  letter-spacing: -.02em;
  text-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
.main-header p {
  color: rgba(255,255,255,0.75) !important;
  margin: 7px 0 0 0;
  font-size: .90rem;
  letter-spacing: .01em;
}
.main-header .header-badge {
  display: inline-block;
  background: rgba(245,158,11,0.20);
  border: 1px solid rgba(245,158,11,0.45);
  color: #fbbf24;
  border-radius: 20px;
  padding: 2px 10px;
  font-size: .72rem;
  font-weight: 700;
  letter-spacing: .05em;
  margin-left: 10px;
  vertical-align: middle;
}

/* ══ STATUS BAR FLOTANTE ══ */
.status-bar {
  position: sticky;
  top: 0;
  z-index: 999;
  background: rgba(13,27,42,0.96);
  backdrop-filter: blur(10px);
  border-bottom: 1px solid rgba(59,130,246,0.2);
  padding: 7px 20px;
  display: flex;
  align-items: center;
  gap: 18px;
  font-size: .76rem;
  font-weight: 700;
  margin: -10px -20px 16px -20px;
  animation: fadeInSlide 0.3s ease both;
}
.status-bar .sb-item {
  display: flex;
  align-items: center;
  gap: 5px;
  color: var(--c-muted);
}
.status-bar .sb-item span { color: var(--c-text); }
.status-bar .sb-dot {
  width: 6px; height: 6px;
  border-radius: 50%;
  display: inline-block;
}
.status-bar .sb-dot.green { background:#22c55e; box-shadow:0 0 6px #22c55e; }
.status-bar .sb-dot.amber { background:#f59e0b; box-shadow:0 0 6px #f59e0b; }
.status-bar .sb-dot.red   { background:#ef4444; box-shadow:0 0 6px #ef4444; }

/* ══ KPI CARDS ══ */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(5,1fr);
  gap: 14px;
  margin: 0 0 22px 0;
}
.kpi-card {
  background: var(--c-card);
  border: 1px solid var(--c-border);
  border-radius: var(--radius);
  padding: 18px 20px 14px;
  animation: fadeInSlide 0.45s ease both;
  position: relative;
  overflow: hidden;
  transition: transform .2s, box-shadow .2s;
  cursor: default;
}
.kpi-card:hover {
  transform: translateY(-3px);
  box-shadow: var(--shadow-blue);
}
.kpi-card::before {
  content:'';
  position:absolute; top:0; left:0; right:0; height:3px;
  border-radius: var(--radius) var(--radius) 0 0;
}
.kpi-card::after {
  content:'';
  position:absolute; top:0; right:0;
  width:60px; height:60px;
  border-radius: 50%;
  opacity: .05;
  transform: translate(15px,-15px);
}
.kpi-card.azul::before  { background:linear-gradient(90deg,#1d4ed8,#60a5fa); }
.kpi-card.verde::before { background:linear-gradient(90deg,#15803d,#4ade80); }
.kpi-card.ambar::before { background:linear-gradient(90deg,#b45309,#fbbf24); }
.kpi-card.rojo::before  { background:linear-gradient(90deg,#991b1b,#f87171); }
.kpi-card.morado::before{ background:linear-gradient(90deg,#6d28d9,#c084fc); }
.kpi-card.azul::after   { background:#3b82f6; }
.kpi-card.verde::after  { background:#22c55e; }
.kpi-card.ambar::after  { background:#f59e0b; }
.kpi-card.rojo::after   { background:#ef4444; }
.kpi-card.morado::after { background:#a855f7; }
.kpi-card.azul   { border-color:rgba(59,130,246,0.30); }
.kpi-card.verde  { border-color:rgba(34,197,94,0.30);  }
.kpi-card.ambar  { border-color:rgba(245,158,11,0.30); }
.kpi-card.rojo   { border-color:rgba(239,68,68,0.30);  }
.kpi-card.morado { border-color:rgba(168,85,247,0.30); }
.kpi-card.critico { animation: pulseGlow 2s ease-in-out infinite; }
.kpi-label {
  font-size:.68rem; font-weight:700; letter-spacing:.08em;
  text-transform:uppercase; color:var(--c-muted); margin-bottom:6px;
}
.kpi-icon {
  position:absolute; top:14px; right:16px;
  font-size:1.3rem; opacity:.35;
}
.kpi-value {
  font-size:1.55rem; font-weight:900; line-height:1.1;
  margin-bottom:2px; letter-spacing:-.02em;
}
.kpi-sub { font-size:.72rem; color:var(--c-muted); margin-top:3px; }
.kpi-delta {
  font-size:.70rem; font-weight:700; margin-top:4px;
  padding: 1px 7px; border-radius: 10px; display:inline-block;
}
.kpi-delta.up   { background:rgba(34,197,94,.15); color:#4ade80; }
.kpi-delta.down { background:rgba(239,68,68,.15); color:#f87171; }
.kpi-delta.neu  { background:rgba(148,163,184,.12); color:#94a3b8; }
.kpi-bar-bg {
  background:rgba(255,255,255,0.08);
  border-radius:4px; height:5px; margin-top:10px; overflow:hidden;
}
.kpi-bar-fill {
  height:5px; border-radius:4px;
  width:var(--fill);
  animation: fillBar 1.2s cubic-bezier(.4,0,.2,1) both;
  animation-delay: var(--delay, 0.1s);
}
.kpi-card.azul  .kpi-bar-fill { background:linear-gradient(90deg,#1d4ed8,#60a5fa); }
.kpi-card.verde .kpi-bar-fill { background:linear-gradient(90deg,#15803d,#4ade80); }
.kpi-card.ambar .kpi-bar-fill { background:linear-gradient(90deg,#b45309,#fbbf24); }
.kpi-card.rojo  .kpi-bar-fill { background:linear-gradient(90deg,#991b1b,#f87171); }
.kpi-card.morado .kpi-bar-fill{ background:linear-gradient(90deg,#6d28d9,#c084fc); }
/* Gauge */
.kpi-gauge { margin-top:8px; text-align:center; }
@media (max-width:900px) { .kpi-grid { grid-template-columns:repeat(3,1fr); } }
@media (max-width:600px) { .kpi-grid { grid-template-columns:repeat(2,1fr); } }

/* ══ BOTONES PREMIUM — 5 variantes ══ */

/* Base para todos los botones */
.stButton > button {
  border-radius: 10px !important;
  font-weight: 700 !important;
  font-size: .84rem !important;
  letter-spacing: .03em !important;
  transition: transform .15s, box-shadow .2s, opacity .15s !important;
  position: relative !important;
  overflow: hidden !important;
  border: none !important;
  padding: 9px 18px !important;
}
.stButton > button:active {
  transform: scale(0.96) !important;
}

/* Ripple via pseudo — activado con JS */
.stButton > button .ripple {
  position: absolute;
  border-radius: 50%;
  background: rgba(255,255,255,0.3);
  transform: scale(0);
  animation: rippleEffect 0.5s linear;
  pointer-events: none;
}

/* Variante PRIMARY — azul gradiente */
.stButton > button[kind="primary"],
.stButton > button[data-testid="baseButton-primary"] {
  background: linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%) !important;
  color: #ffffff !important;
  box-shadow: 0 3px 14px rgba(59,130,246,0.40) !important;
}
.stButton > button[kind="primary"]:hover {
  transform: translateY(-2px) !important;
  box-shadow: 0 6px 22px rgba(59,130,246,0.55) !important;
}

/* Variante SECONDARY — ghost azul */
.stButton > button[kind="secondary"],
.stButton > button[data-testid="baseButton-secondary"] {
  background: rgba(59,130,246,0.10) !important;
  border: 1px solid rgba(59,130,246,0.40) !important;
  color: #60a5fa !important;
}
.stButton > button[kind="secondary"]:hover {
  background: rgba(59,130,246,0.20) !important;
  border-color: rgba(59,130,246,0.70) !important;
  transform: translateY(-1px) !important;
}

/* Sidebar — todos los botones con variante ghost dorada */
[data-testid="stSidebar"] .stButton > button {
  background: rgba(59,130,246,0.12) !important;
  border: 1px solid rgba(59,130,246,0.30) !important;
  color: #e2e8f0 !important;
  border-radius: 9px !important;
  font-weight: 700 !important;
  transition: all .2s !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: rgba(59,130,246,0.25) !important;
  border-color: rgba(96,165,250,0.60) !important;
  transform: translateX(2px) !important;
  color: #60a5fa !important;
}

/* Botón de logout — rojo suave */
[data-testid="stSidebar"] .stButton > button[kind="secondary"] {
  background: rgba(239,68,68,0.10) !important;
  border-color: rgba(239,68,68,0.30) !important;
  color: #fca5a5 !important;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
  background: rgba(239,68,68,0.20) !important;
  border-color: rgba(239,68,68,0.60) !important;
  color: #f87171 !important;
}

/* Botón disabled */
.stButton > button:disabled {
  opacity: .38 !important;
  cursor: not-allowed !important;
  transform: none !important;
}

/* Download button — verde */
[data-testid="stDownloadButton"] > button {
  background: linear-gradient(135deg,#15803d,#22c55e) !important;
  border: none !important;
  color: #ffffff !important;
  font-weight: 700 !important;
  border-radius: 10px !important;
  box-shadow: 0 3px 14px rgba(34,197,94,0.35) !important;
  transition: transform .15s, box-shadow .2s !important;
}
[data-testid="stDownloadButton"] > button:hover {
  transform: translateY(-2px) !important;
  box-shadow: 0 6px 22px rgba(34,197,94,0.50) !important;
}

/* ══ TABS PREMIUM ══ */
[data-baseweb="tab-list"] {
  background: rgba(15,25,45,0.6) !important;
  border-radius: 12px !important;
  padding: 4px !important;
  gap: 2px !important;
  border: 1px solid rgba(59,130,246,0.15) !important;
}
button[data-baseweb="tab"] {
  border-radius: 9px !important;
  font-weight: 600 !important;
  font-size: .80rem !important;
  padding: 7px 12px !important;
  transition: all .2s !important;
  color: var(--c-muted) !important;
}
button[data-baseweb="tab"]:hover {
  background: rgba(59,130,246,0.12) !important;
  color: #93c5fd !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
  background: rgba(59,130,246,0.22) !important;
  color: #60a5fa !important;
  border-bottom: 3px solid #3b82f6 !important;
  font-weight: 800 !important;
}

/* ══ MÉTRICAS NATIVAS ══ */
[data-testid="metric-container"] {
  background: rgba(59,130,246,0.07) !important;
  border: 1px solid rgba(59,130,246,0.22) !important;
  border-left: 4px solid #3b82f6 !important;
  border-radius: 10px !important;
  padding: 14px 18px !important;
  transition: transform .2s, box-shadow .2s !important;
}
[data-testid="metric-container"]:hover {
  transform: translateY(-2px) !important;
  box-shadow: 0 4px 16px rgba(59,130,246,0.20) !important;
}
[data-testid="metric-container"] label {
  font-size:.72rem !important; font-weight:700 !important;
  letter-spacing:.07em !important; text-transform:uppercase !important;
  color:var(--c-muted) !important;
}
[data-testid="stMetricValue"] {
  color:#60a5fa !important; font-size:1.25rem !important; font-weight:800 !important;
}
[data-testid="stMetricDelta"] { font-size:.76rem !important; font-weight:700 !important; }

/* ══ BADGES PREMIUM ══ */
.badge-verde, .badge-rojo, .badge-naranja, .badge-azul,
.badge-purple, .badge-cyan, .badge-gold {
  display:inline-block; border-radius:20px;
  padding:3px 12px; font-size:.72rem; font-weight:800;
  animation: badgePop .35s ease both;
  letter-spacing:.03em;
}
.badge-verde   { background:rgba(34,197,94,.18);  color:#4ade80; border:1px solid rgba(34,197,94,.35); }
.badge-rojo    { background:rgba(239,68,68,.18);  color:#f87171; border:1px solid rgba(239,68,68,.35); }
.badge-naranja { background:rgba(245,158,11,.18); color:#fbbf24; border:1px solid rgba(245,158,11,.35); }
.badge-azul    { background:rgba(59,130,246,.18); color:#60a5fa; border:1px solid rgba(59,130,246,.35); }
.badge-purple  { background:rgba(168,85,247,.18); color:#c084fc; border:1px solid rgba(168,85,247,.35); }
.badge-cyan    { background:rgba(6,182,212,.18);  color:#67e8f9; border:1px solid rgba(6,182,212,.35); }
.badge-gold    { background:rgba(245,158,11,.22); color:#fbbf24; border:1px solid rgba(245,158,11,.50); }

/* ══ BADGES EN TABLA (columna ESTADO) ══ */
.estado-badge {
  display:inline-flex; align-items:center; gap:5px;
  border-radius:20px; padding:3px 11px;
  font-size:.72rem; font-weight:800;
  animation: badgePop .4s ease both;
  white-space:nowrap;
}
.estado-exacto  { background:rgba(34,197,94,.15);  color:#4ade80; border:1px solid rgba(34,197,94,.35); }
.estado-aprox   { background:rgba(245,158,11,.15); color:#fbbf24; border:1px solid rgba(245,158,11,.35); }
.estado-solo    { background:rgba(239,68,68,.15);  color:#f87171; border:1px solid rgba(239,68,68,.35); }
.estado-agrup   { background:rgba(59,130,246,.15); color:#60a5fa; border:1px solid rgba(59,130,246,.35); }
.estado-rechazo { background:rgba(168,85,247,.15); color:#c084fc; border:1px solid rgba(168,85,247,.35); }
.estado-otros   { background:rgba(148,163,184,.12);color:#94a3b8; border:1px solid rgba(148,163,184,.25);}

/* ══ CALLOUT BOXES PREMIUM ══ */
.callout-info, .callout-success, .callout-warning, .callout-danger, .callout-accion {
  border-radius: 10px;
  padding: 14px 18px;
  margin: 10px 0;
  color: inherit;
  line-height: 1.65;
  animation: fadeInSlide .4s ease both;
  transition: box-shadow .2s;
}
.callout-info    { background:rgba(59,130,246,.09);  border-left:4px solid #3b82f6; }
.callout-success { background:rgba(34,197,94,.09);   border-left:4px solid #22c55e; }
.callout-warning { background:rgba(245,158,11,.09);  border-left:4px solid #f59e0b; }
.callout-danger  { background:rgba(239,68,68,.09);   border-left:4px solid #ef4444; }
.callout-accion  { background:rgba(168,85,247,.09);  border-left:4px solid #a855f7;
  font-family:'Consolas','Courier New',monospace; font-size:.87rem; }
.callout-info:hover    { box-shadow:0 2px 12px rgba(59,130,246,.15); }
.callout-success:hover { box-shadow:0 2px 12px rgba(34,197,94,.15); }
.callout-warning:hover { box-shadow:0 2px 12px rgba(245,158,11,.15); }
.callout-danger:hover  { box-shadow:0 2px 12px rgba(239,68,68,.15); }

/* ══ TÍTULOS DE SECCIÓN ══ */
.section-title {
  color:#60a5fa; font-size:.95rem; font-weight:800;
  letter-spacing:.05em; text-transform:uppercase;
  border-bottom:1px solid rgba(59,130,246,.25);
  padding-bottom:7px; margin:20px 0 13px 0;
  display:flex; align-items:center; gap:8px;
}
.section-title::after {
  content:''; flex:1; height:1px;
  background:linear-gradient(90deg,rgba(59,130,246,.25),transparent);
}

/* ══ DATAFRAMES ══ */
[data-testid="stDataFrame"] {
  border-radius:10px !important; overflow:hidden !important;
  border:1px solid rgba(59,130,246,.18) !important;
  animation: fadeInSlide .4s ease both;
}

/* ══ FILE UPLOADER ══ */
[data-testid="stFileUploader"] {
  border:2px dashed rgba(59,130,246,.25) !important;
  border-radius:12px !important;
  transition: border-color .25s, background .25s !important;
}
[data-testid="stFileUploader"]:hover {
  border-color:rgba(96,165,250,.60) !important;
  background:rgba(59,130,246,.04) !important;
}

/* ══ SIDEBAR PREMIUM ══ */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg,#0a1628 0%,#0d1e35 50%,#091422 100%) !important;
  border-right:1px solid rgba(59,130,246,.18) !important;
}
[data-testid="stSidebar"] * { color:#e2e8f0 !important; }
[data-testid="stSidebar"] hr {
  border-color:rgba(59,130,246,.15) !important;
}

/* ══ PROGRESS BAR ══ */
[data-testid="stProgressBar"] > div > div {
  background:linear-gradient(90deg,#1d4ed8,#60a5fa) !important;
  border-radius:4px !important;
}

/* ══ EXPANDERS ══ */
details > summary {
  font-weight:700 !important; font-size:.90rem !important;
  border-radius:8px !important;
  transition:background .2s !important;
}
details > summary:hover { background:rgba(59,130,246,.07) !important; }

/* ══ SELECTBOX / INPUTS ══ */
[data-baseweb="select"] > div, [data-baseweb="input"] > div {
  border-radius:9px !important;
  border-color:rgba(59,130,246,.22) !important;
  transition:border-color .2s !important;
}
[data-baseweb="select"] > div:hover, [data-baseweb="input"] > div:hover {
  border-color:rgba(96,165,250,.55) !important;
}
[data-baseweb="select"] > div:focus-within, [data-baseweb="input"] > div:focus-within {
  border-color:#3b82f6 !important;
  box-shadow:0 0 0 3px rgba(59,130,246,.15) !important;
}

/* ══ CHECKBOX ══ */
[data-baseweb="checkbox"] svg { fill:#3b82f6 !important; }

/* ══ SPINNER ══ */
[data-testid="stSpinner"] > div { border-top-color:#3b82f6 !important; }

/* ══ TOAST NOTIFICATIONS PREMIUM ══ */
.toast-container {
  position:fixed; top:14px; right:14px; z-index:99999;
  display:flex; flex-direction:column; gap:8px; pointer-events:none;
}
.toast {
  display:flex; align-items:center; gap:10px;
  padding:12px 16px; border-radius:12px;
  font-size:.84rem; font-weight:700;
  min-width:270px; max-width:360px;
  animation:slideInRight .35s ease, fadeOutToast .5s ease 3.5s both;
  border:1px solid transparent; pointer-events:auto;
  box-shadow:0 4px 20px rgba(0,0,0,0.4);
  position:relative; overflow:hidden;
}
.toast::after {
  content:''; position:absolute; bottom:0; left:0; right:0; height:3px;
  background:rgba(255,255,255,.3);
  animation:fillBar 3.5s linear both; --fill:100%;
}
.toast-success { background:rgba(15,60,30,.97); border-color:rgba(34,197,94,.4); color:#4ade80; }
.toast-warning { background:rgba(90,40,0,.97);  border-color:rgba(245,158,11,.4);color:#fbbf24; }
.toast-error   { background:rgba(80,10,10,.97); border-color:rgba(239,68,68,.4); color:#f87171; }
.toast-info    { background:rgba(10,40,100,.97);border-color:rgba(59,130,246,.4);color:#60a5fa; }

/* ══ GUIA-ROW ══ */
.guia-row {
  background:rgba(59,130,246,.06);
  border-radius:9px; padding:12px 16px; margin:6px 0;
  border:1px solid rgba(59,130,246,.18);
  line-height:1.7;
  transition:background .2s, border-color .2s, transform .15s;
}
.guia-row:hover {
  background:rgba(59,130,246,.12);
  border-color:rgba(96,165,250,.45);
  transform:translateX(3px);
}
</style>
""", unsafe_allow_html=True)

# ── Control de acceso multi-usuario ──────────────────────────────────────────
def _pantalla_login() -> bool:
    """Pantalla de login multi-usuario o fallback a contraseña simple."""
    # Ya autenticado
    if st.session_state.get("usuario_actual"):
        return True

    st.markdown(f"""
    <div style='max-width:440px;margin:60px auto 0;'>
      <div style='background:{_bg_card};border:1px solid {_color_corp}{_borde_op};
                  border-radius:16px;padding:40px;text-align:center;'>
        <div style='font-size:3rem;margin-bottom:4px;'>🏦</div>
        <div style='font-size:1.5rem;font-weight:800;color:{_color_gold};
                    letter-spacing:.04em;margin-bottom:4px;'>{_empresa}</div>
        <div style='font-size:.9rem;color:{_txt_sec};margin-bottom:28px;'>
          Sistema de Conciliación Bancaria</div>
    </div></div>""", unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        if _AUTH_DISPONIBLE:
            # Login multi-usuario SQLite
            usr = st.text_input("👤 Usuario", placeholder="admin",
                                key="_login_usr")
            pwd = st.text_input("🔒 Contraseña", type="password",
                                placeholder="••••••••", key="_login_pwd")
            if st.button("🚀 Ingresar", use_container_width=True, type="primary"):
                ok, usuario = verificar_credenciales(usr, pwd)
                if ok and usuario:
                    st.session_state["usuario_actual"] = usuario
                    registrar_auditoria(usr, "LOGIN", "app", "Inicio sesión exitoso")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos")
                    registrar_auditoria(usr or "—", "LOGIN_FAIL", "app",
                                        "Intento fallido", "ERROR")
        else:
            # Fallback: contraseña simple (hmac)
            def _pw_entered():
                if hmac.compare_digest(
                    st.session_state.get("_pw", ""),
                    st.secrets.get("APP_PASSWORD", "crediexpress2025")
                ):
                    st.session_state["usuario_actual"] = {
                        "username": "admin", "rol": "admin",
                        "nombre": "Administrador",
                        "permisos": ["conciliar","exportar","dian","usuarios",
                                     "config","partidas","auditoria","backup"],
                    }
                    del st.session_state["_pw"]
                else:
                    st.session_state["_pw_err"] = True
            st.text_input("🔒 Contraseña", type="password",
                          on_change=_pw_entered, key="_pw",
                          placeholder="Contraseña de acceso")
            if st.session_state.get("_pw_err"):
                st.error("❌ Contraseña incorrecta")

        st.caption("💡 Contacte al administrador si olvidó su contraseña.")
    return bool(st.session_state.get("usuario_actual"))

if not _pantalla_login():
    st.stop()

# Datos del usuario en sesión
_usuario_actual = st.session_state.get("usuario_actual", {
    "username": "admin", "rol": "admin", "nombre": "Administrador",
    "permisos": ["conciliar","exportar","dian","usuarios","config",
                 "partidas","auditoria","backup"],
})
def _tiene_permiso(p: str) -> bool:
    return p in _usuario_actual.get("permisos", [])

st.markdown(f"""
<div class='main-header'>
  <h1>🏦 Conciliación Bancaria — {_empresa}
    <span class="header-badge">v3.0 PREMIUM</span>
  </h1>
  <p>⚡ Extracto Bancario &nbsp;↔&nbsp; Auxiliar Contable &nbsp;·&nbsp;
     🤖 Detección automática de formato &nbsp;·&nbsp;
     📊 Análisis inteligente de diferencias &nbsp;·&nbsp;
     🇨🇴 Motor DIAN Colombia</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    # ── Logo + perfil usuario ─────────────────────────────────────────────────
    _usr_rol = _usuario_actual.get("rol", "admin")
    _usr_nom = _usuario_actual.get("nombre", "Admin")
    _icon_rol = {"admin":"🔑","contador_senior":"📊","auxiliar":"📋"}.get(_usr_rol,"👤")
    st.markdown(f"""
    <div style='text-align:center;padding:10px 0 14px 0;'>
      <div style='font-size:2.6rem;filter:drop-shadow(0 0 12px rgba(59,130,246,.5));'>🏦</div>
      <div style='font-size:1.05rem;font-weight:900;letter-spacing:.04em;
                  color:{_color_gold};text-shadow:0 0 12px rgba(245,158,11,.35);
                  margin-top:4px;'>{_empresa[:22]}</div>
      <div style='font-size:.70rem;color:#64748b;letter-spacing:.06em;
                  text-transform:uppercase;margin-top:2px;'>Conciliación Bancaria v3.0</div>
      <div style='margin-top:10px;background:rgba(59,130,246,.10);
                  border:1px solid rgba(59,130,246,.25);border-radius:10px;
                  padding:8px 12px;font-size:.78rem;
                  display:flex;align-items:center;gap:8px;justify-content:center;'>
        <span style='font-size:1.1rem;'>{_icon_rol}</span>
        <div style='text-align:left;'>
          <div style='font-weight:800;color:#e2e8f0;'>{_usr_nom}</div>
          <div style='font-size:.68rem;color:#64748b;text-transform:uppercase;
                      letter-spacing:.05em;'>{_usr_rol}</div>
        </div>
        <span style='margin-left:auto;width:8px;height:8px;background:#22c55e;
                     border-radius:50%;box-shadow:0 0 6px #22c55e;
                     display:inline-block;'></span>
      </div>
    </div>""", unsafe_allow_html=True)

    # Botón logout
    if st.button("🚪 Cerrar sesión", use_container_width=True, type="secondary"):
        st.session_state.pop("usuario_actual", None)
        st.rerun()

    st.markdown("---")

    # ── Paso 1: Archivos ──────────────────────────────────────────────────────
    st.markdown("**📂 Paso 1 — Cargar archivos**")
    st.markdown("#### 📂 Cargar Archivos")
    banco_files = st.file_uploader(
        "🏦 Extracto Bancario",
        type=["pdf","csv","xlsx","txt"],
        accept_multiple_files=True,
        help="Puedes subir 1 o más archivos del mismo mes (PDF, CSV o Excel). "
             "El sistema los unifica automáticamente.",
    )
    aux_files = st.file_uploader(
        "📋 Auxiliar Contable",
        type=["pdf","csv","xlsx","txt"],
        accept_multiple_files=True,
        help="Puedes subir 1 o más archivos: por tipo de comprobante (NC, CE, CG) o por quincena.",
    )

    # Compatibilidad: banco_file / aux_file = primer archivo (para código heredado)
    banco_file = banco_files[0] if banco_files else None
    aux_file   = aux_files[0]   if aux_files   else None

    usar_ocr = st.checkbox("🔍 Forzar OCR en PDF escaneados", value=True,
                            help="Requiere Tesseract + Poppler instalados.")

    # Resumen de archivos cargados
    def _mostrar_archivos(files, icono):
        for f in files:
            periodo_nom = _extraer_periodo(f.name)
            per_lbl = (f" · {_MESES_ES[periodo_nom[1]-1]} {periodo_nom[0]}"
                       if periodo_nom else "")
            st.markdown(
                f"<div style='background:#ffffff22;border-radius:6px;"
                f"padding:5px 10px;font-size:.79rem;margin-top:3px;'>"
                f"{icono} <b>{f.name}</b>{per_lbl}<br>"
                f"<span style='opacity:.7;'>{f.size/1024:.1f} KB</span></div>",
                unsafe_allow_html=True,
            )

    if banco_files:
        _mostrar_archivos(banco_files, "🏦")
    if aux_files:
        _mostrar_archivos(aux_files, "📋")

    st.markdown("<br>", unsafe_allow_html=True)
    ejecutar = st.button(
        "🚀 Ejecutar análisis completo",
        disabled=not (banco_files and aux_files),
        use_container_width=True,
    )
    if ejecutar:
        st.session_state.run = True

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════════════
    # 💾 BACKUP MULTICAPA — IndexedDB + File System API + Descarga + Restaurar
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown("#### 💾 Datos & Backup")

    # Detectar modo nube: en Streamlit Cloud no existe el .bat local
    _en_nube = not os.path.isfile(os.path.join(BASE_DIR, '2_ABRIR_APP.bat'))
    if _en_nube:
        st.markdown(
            "<div style='background:#1a3a5e;border-radius:8px;padding:8px 12px;"
            "font-size:.78rem;margin-bottom:8px;'>☁️ <b>Modo nube</b> — "
            "descarga tu backup al terminar cada sesión para no perder datos.</div>",
            unsafe_allow_html=True,
        )

    # ── Descargar backup ──────────────────────────────────────────────────────
    _db_b64   = ""
    _fecha_bk = datetime.now().strftime('%Y-%m-%d')
    if os.path.exists(DB_PATH):
        with open(DB_PATH, 'rb') as _f:
            _db_bytes = _f.read()
        _db_b64 = __import__('base64').b64encode(_db_bytes).decode()
        st.download_button(
            "⬇️ Descargar backup",
            data                = _db_bytes,
            file_name           = f"CREDIEXPRESS_{_fecha_bk}.db",
            mime                = "application/octet-stream",
            use_container_width = True,
            help                = "Guarda el backup completo en tu computador",
        )
    else:
        st.button("⬇️ Descargar backup", disabled=True,
                  use_container_width=True, help="Aún no hay datos guardados")

    # ── Restaurar backup ──────────────────────────────────────────────────────
    _bk_up = st.file_uploader(
        "⬆️ Restaurar backup (.db)",
        type = ["db"],
        key  = "bk_restore",
        help = "Sube un archivo .db descargado previamente para restaurar tus datos",
    )
    if _bk_up is not None:
        with open(DB_PATH, 'wb') as _f:
            _f.write(_bk_up.getvalue())
        st.success("✅ Backup restaurado correctamente.")
        st.rerun()

    # ── Componente JS: IndexedDB + File System Access API ─────────────────────
    if _db_b64 and len(_db_b64) < 15_000_000:
        import streamlit.components.v1 as _stcv1
        _stcv1.html(f"""
<style>
  body{{margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}}
  .bk-btn{{display:block;width:100%;padding:8px 6px;margin:4px 0;border:none;
           border-radius:7px;cursor:pointer;font-size:12px;font-weight:700;
           letter-spacing:.01em;transition:opacity .15s;}}
  .bk-btn:hover{{opacity:.82;}}
  #btn-idb {{background:#1a3a5e;color:#cce4ff;}}
  #btn-fs  {{background:#14532d;color:#bbf7d0;}}
  #btn-ridb{{background:#2d1f00;color:#ffe08a;}}
  #bk-st   {{font-size:11px;margin-top:6px;padding:5px 8px;border-radius:5px;
             background:rgba(255,255,255,.06);min-height:18px;color:#8ab4d4;
             word-break:break-word;}}
</style>
<button id="btn-idb"  class="bk-btn" onclick="guardarIDB()">🗄️ Guardar en este navegador (IndexedDB)</button>
<button id="btn-fs"   class="bk-btn" onclick="guardarFS()">📁 Auto-guardar en carpeta del PC</button>
<button id="btn-ridb" class="bk-btn" onclick="restaurarIDB()">🔄 Descargar backup guardado en navegador</button>
<div id="bk-st">Verificando backup en navegador...</div>
<script>
const DB_B64="{_db_b64}";const FECHA="{_fecha_bk}";const IDB_DB="CREDIEXPRESS_BACKUPS";
let _dh=null;
function logSt(m,c){{const s=document.getElementById("bk-st");s.textContent=m;s.style.color=c||"#8ab4d4";}}
function b64ToBytes(b64){{const bin=atob(b64),arr=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);return arr;}}
function abrirIDB(cb){{const r=indexedDB.open(IDB_DB,1);r.onupgradeneeded=e=>e.target.result.createObjectStore("backups");r.onsuccess=e=>cb(e.target.result);r.onerror=()=>logSt("❌ IndexedDB no disponible","#f44336");}}
function guardarIDB(){{logSt("Guardando...","#90a4ae");abrirIDB(db=>{{const tx=db.transaction("backups","readwrite");tx.objectStore("backups").put({{data:b64ToBytes(DB_B64),fecha:FECHA,ts:Date.now()}},"ultimo");tx.oncomplete=()=>logSt("✅ Guardado en navegador — "+FECHA,"#4caf50");tx.onerror=e=>logSt("❌ "+e.target.error,"#f44336");}});}}
async function guardarFS(){{if(!window.showDirectoryPicker){{logSt("⚠️ Requiere Chrome o Edge 86+","#ff9800");return;}}logSt("Selecciona carpeta...","#90a4ae");try{{if(!_dh)_dh=await window.showDirectoryPicker({{id:"crediexpress_bk",mode:"readwrite",startIn:"documents"}});const fh=await _dh.getFileHandle("CREDIEXPRESS_"+FECHA+".db",{{create:true}});const wr=await fh.createWritable();await wr.write(b64ToBytes(DB_B64));await wr.close();logSt("✅ Guardado en PC — CREDIEXPRESS_"+FECHA+".db","#4caf50");guardarIDB();}}catch(e){{if(e.name!=="AbortError")logSt("❌ "+e.message,"#f44336");else logSt("Cancelado.","#90a4ae");}}}}
function restaurarIDB(){{abrirIDB(db=>{{const tx=db.transaction("backups","readonly");const req=tx.objectStore("backups").get("ultimo");req.onsuccess=e=>{{const rec=e.target.result;if(!rec){{logSt("⚠️ Sin backup en este navegador","#ff9800");return;}}const blob=new Blob([rec.data],{{type:"application/octet-stream"}});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="CREDIEXPRESS_"+rec.fecha+".db";document.body.appendChild(a);a.click();setTimeout(()=>{{URL.revokeObjectURL(url);a.remove();}},1200);logSt("⬇️ Descargando backup del "+rec.fecha+" — luego súbelo con Restaurar","#42a5f5");}};req.onerror=()=>logSt("❌ Error leyendo IndexedDB","#f44336");}});}}
(function checkExisting(){{try{{const r=indexedDB.open(IDB_DB,1);r.onupgradeneeded=e=>e.target.result.createObjectStore("backups");r.onsuccess=e=>{{const tx=e.target.result.transaction("backups","readonly");const req=tx.objectStore("backups").get("ultimo");req.onsuccess=e=>{{const rec=e.target.result;if(rec&&rec.fecha)logSt("🗄️ Backup en navegador disponible — "+rec.fecha,"#42a5f5");else logSt("Sin backup en navegador aún — usa los botones de arriba.","#607d8b");}};}}}}catch(ex){{logSt("IndexedDB no disponible.","#607d8b");}}}}());
</script>
""", height=210, scrolling=False)
    elif _db_b64:
        st.warning("⚠️ Base de datos > 10 MB — usa solo la descarga manual.", icon="⚠️")

    st.markdown("---")
    st.markdown("#### 💡 Formatos soportados")
    st.markdown("""
- **PDF** — extracto bancario original
- **CSV** — SIIGO / Helisa / World Office
- **Excel (.xlsx)** — cualquier formato tabular
- **TXT** — texto plano con columnas
    """)
    if OCR_AVAILABLE:
        st.markdown("<span class='badge-verde'>✅ OCR disponible</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='badge-naranja'>⚠️ OCR no instalado</span>", unsafe_allow_html=True)
        st.caption("Instale pytesseract + Poppler para PDFs escaneados.")

    st.markdown("---")
    st.markdown("<div style='font-size:.72rem;opacity:.7;text-align:center;'>v2.0 · CREDIEXPRESS POPAYÁN SAS<br>Desarrollado con ❤️ en Python + Streamlit</div>", unsafe_allow_html=True)
    # ── Auto-guardar archivos subidos (solo offline) ─────────────────────────
    if OFFLINE_MODE:
        for _uf, _sub in [(f, "datos_entrada") for f in (banco_files or []) + (aux_files or [])]:
            if _uf:
                _ruta, _nuevo = _auto_guardar_archivo(_uf, _sub)
                if _nuevo and _ruta:
                    st.caption(f"💾 Guardado: .../{_uf.name}")

    # ── Panel historial (solo offline) ───────────────────────────────────────
    if OFFLINE_MODE:
        # ── Fase C: formatos aprendidos ──────────────────────────────────
        _formatos = listar_formatos_aprendidos()
        if _formatos:
            st.markdown("---")
            st.markdown("**🧠 Formatos Aprendidos**")
            for _fma, _tipo, _banco, _usos, _ultima in _formatos[:5]:
                _ico2 = "📄" if _tipo == "auxiliar" else "🏦"
                st.markdown(
                    f"<small>{_ico2} <b>{_fma}</b><br>"
                    f"&nbsp;&nbsp;{_usos} uso{'s' if _usos!=1 else ''}"
                    f" · {(_ultima or '')[:10]}</small>",
                    unsafe_allow_html=True
                )
        # ── Fase D5: catalogo NC aprendido ────────────────────────────────
        try:
            _cat_rows, _cat_total, _cat_pend = listar_catalogo_nc(5)
        except Exception:
            _cat_rows, _cat_total, _cat_pend = [], 0, 0
        st.markdown("---")
        st.markdown(
            f"**📚 Catálogo NC** &nbsp;"
            f"<span style='color:#42a5f5;font-size:.8rem;'>"
            f"{_cat_total} reglas · {_cat_pend} pendientes</span>",
            unsafe_allow_html=True
        )
        if _cat_rows:
            for _cr in _cat_rows:
                _uuid, _bt, _at, _conf, _nivel, _apr, _ult = _cr
                _ico_nv = ("🟢" if _nivel=='ALTA'
                           else "🟡" if _nivel=='MEDIA'
                           else "⏳")
                try:
                    _bt_tok = json.loads(_bt or '[]')[:3]
                    _at_tok = json.loads(_at or '[]')[:3]
                    _lbl = ' '.join(_bt_tok) + ' ↔ ' + ' '.join(_at_tok)
                except Exception:
                    _lbl = _uuid
                st.markdown(
                    f"<small>{_ico_nv} {_lbl}<br>"
                    f"&nbsp;&nbsp;{_conf} confirmaci{'o' if _conf==1 else 'o'}nes"
                    f" · {(_ult or '')[:10]}</small>",
                    unsafe_allow_html=True
                )
        else:
            st.markdown(
                "<small style='opacity:.5;'>Aun sin reglas aprendidas.<br>"
                "Se llenara automaticamente al procesar PDFs.</small>",
                unsafe_allow_html=True
            )
        # ── Boton de sincronizacion ───────────────────────────────────────
        st.markdown("")
        if st.button("🔄 Sincronizar con Cloud", use_container_width=True,
                     help="Sube reglas nuevas a Google Sheets y baja las del cloud"):
            with st.spinner("Sincronizando..."):
                _n_up, _n_down = sincronizar_catalogo_nc()
            st.success(
                f"✅ Sync OK — "
                f"↑{_n_up} subidas · ↓{_n_down} bajadas"
            )
        _hist = leer_historial_sqlite(6)
        if _hist:
            st.markdown("---")
            st.markdown("#### 📜 Historial")
            for _h in _hist:
                _fh, _fb, _fa, _per, _tasa, _ex, _nb, _dif = _h
                _ico = "🟢" if _tasa >= 90 else ("🟡" if _tasa >= 75 else "🔴")
                _per_lbl = f" · {_per}" if _per else ""
                st.markdown(f"""
<div style='background:rgba(255,255,255,0.08);border-radius:6px;
            padding:7px 10px;margin:3px 0;font-size:.74rem;line-height:1.5;'>
  {_ico} <b>{_fh[:16]}</b>{_per_lbl}<br>
  <span style='opacity:.65;'>{os.path.basename(_fb)[:26]}</span><br>
  <span style='color:#42a5f5;font-weight:700;'>{_tasa:.0f}% conciliado
    &nbsp;·&nbsp; {_ex}/{_nb} mov.</span>
</div>""", unsafe_allow_html=True)


if 'run' in st.session_state and st.session_state.run:
    with st.spinner("Procesando archivos..."):
        try:
            # ── Parsear TODOS los extractos bancarios y unificar ─────────────
            _dfs_banco, _res_banco_list, _leg_banco_list = [], [], []
            for _bf in banco_files:
                _df_i, _res_i, _leg_i = cargar_y_parsear(_bf, 'BANCO', usar_ocr=usar_ocr)
                _dfs_banco.append(_df_i)
                _res_banco_list.append(_res_i)
                _leg_banco_list.append(_leg_i)

            # Consolidar DataFrame banco
            df_banco = pd.concat(_dfs_banco, ignore_index=True) if _dfs_banco else pd.DataFrame()

            # Saldos: primer archivo = SA, último archivo = SAC; totales = suma
            _res_banco_ord = _res_banco_list  # ya están en orden de carga
            _r0 = _res_banco_ord[0]
            _rn = _res_banco_ord[-1]
            # Usar 'in' para no tratar saldo=0 como ausente (bug del operador 'or')
            sa  = _r0['SALDO_INICIAL']  if 'SALDO_INICIAL'  in _r0 else (_r0.get('SALDO_ANTERIOR') or 0)
            sac = _rn['SALDO_FINAL']    if 'SALDO_FINAL'    in _rn else (_rn.get('SALDO_ACTUAL')   or 0)
            # Fallback: si sac sigue en 0, tomar último saldo del DataFrame
            if not sac and not df_banco.empty and 'SALDO' in df_banco.columns:
                _sb = df_banco['SALDO'].dropna()
                if not _sb.empty:
                    sac = float(_sb.iloc[-1])
            tab_s = sum(r.get('TOTAL_ABONOS', 0) or 0 for r in _res_banco_list)
            tca_s = sum(abs(r.get('TOTAL_CARGOS', 0) or 0) for r in _res_banco_list)

            # Legibilidad: promedio ponderado
            _p_b = sum(lg[0] for lg in _leg_banco_list) / max(len(_leg_banco_list), 1)
            _cal_b = _leg_banco_list[0][1] if _leg_banco_list else ''
            _adv_b = [a for lg in _leg_banco_list for a in (lg[2] or [])]
            _fmt_b = _leg_banco_list[0][3] if _leg_banco_list else ''
            _cf_b  = _leg_banco_list[0][4] if _leg_banco_list else 0
            leg_banco = (_p_b, _cal_b, _adv_b, _fmt_b, _cf_b)

            # ── Parsear TODOS los auxiliares contables y unificar ────────────
            _dfs_aux, _meta_aux_list, _leg_aux_list = [], [], []
            for _af in aux_files:
                _df_j, _meta_j, _leg_j = cargar_y_parsear(_af, 'AUXILIAR', usar_ocr=usar_ocr)
                _dfs_aux.append(_df_j)
                _meta_aux_list.append(_meta_j)
                _leg_aux_list.append(_leg_j)

            df_aux = pd.concat(_dfs_aux, ignore_index=True) if _dfs_aux else pd.DataFrame()

            _ma0 = _meta_aux_list[0]
            _man = _meta_aux_list[-1]
            si_a = _ma0['SALDO_INICIAL'] if 'SALDO_INICIAL' in _ma0 else 0
            sf_a = _man['SALDO_FINAL']   if 'SALDO_FINAL'   in _man else 0
            td_a = sum(m.get('TOTAL_DEBITOS', 0) or 0 for m in _meta_aux_list)
            tc_a = sum(m.get('TOTAL_CREDITOS', 0) or 0 for m in _meta_aux_list)

            _p_a = sum(lg[0] for lg in _leg_aux_list) / max(len(_leg_aux_list), 1)
            _cal_a = _leg_aux_list[0][1] if _leg_aux_list else ''
            _adv_a = [a for lg in _leg_aux_list for a in (lg[2] or [])]
            _fmt_a = _leg_aux_list[0][3] if _leg_aux_list else ''
            _cf_a  = _leg_aux_list[0][4] if _leg_aux_list else 0
            leg_aux = (_p_a, _cal_a, _adv_a, _fmt_a, _cf_a)

        except Exception as e:
            st.error(f"❌ Error al procesar los archivos: {e}")
            import traceback; log.error(traceback.format_exc())
            st.stop()

    # ── Validación de período — 3 niveles ────────────────────────────────────
    _nombre_b = " + ".join(f.name for f in banco_files)
    _nombre_a = " + ".join(f.name for f in aux_files)
    _val_per  = validar_periodo_archivos(_nombre_b, df_banco, _nombre_a, df_aux)
    _periodo_detectado = _val_per['periodo']

    if not _val_per['ok']:
        # BLOQUEO — períodos distintos
        st.markdown(f"""
<div class='callout-danger'>
  <b>🚫 PERÍODOS DISTINTOS — Conciliación bloqueada</b><br>
  {_val_per['mensaje']}<br><br>
  Banco: <b>{_MESES_ES[_val_per['periodo_banco'][1]-1]} {_val_per['periodo_banco'][0]}</b>
  (confianza {_val_per['conf_banco']:.0f}%,
  detectado desde {('nombre + fechas' if '1+2' in (_val_per['nivel'] or '') else 'fechas del contenido')})<br>
  Auxiliar: <b>{_MESES_ES[_val_per['periodo_aux'][1]-1]} {_val_per['periodo_aux'][0]}</b>
  (confianza {_val_per['conf_aux']:.0f}%)<br><br>
  <small>Suba archivos del mismo mes y año para continuar.</small>
</div>""", unsafe_allow_html=True)
        st.session_state.run = False
        st.stop()
    elif _periodo_detectado:
        _mes_lbl  = _MESES_ES[_periodo_detectado[1]-1]
        _conf_avg = (_val_per['conf_banco'] + _val_per['conf_aux']) / 2
        _nivel_str = _val_per.get('nivel','')
        _icono_niv = "🔒" if '1+2' in _nivel_str else "✅"
        st.markdown(f"""
<div class='callout-success' style='padding:9px 16px;font-size:.88rem;'>
  {_icono_niv} <b>Período verificado: {_mes_lbl} {_periodo_detectado[0]}</b>
  &nbsp;·&nbsp; Confianza promedio: <b>{_conf_avg:.0f}%</b>
  &nbsp;·&nbsp; <small style='opacity:.7;'>Fuente: {_nivel_str}</small>
</div>""", unsafe_allow_html=True)
    else:
        st.markdown("""
<div class='callout-warning' style='padding:8px 14px;font-size:.86rem;'>
  ⚠️ No se detectó período en los archivos. Verifique que los nombres incluyan el mes y año
  (ej: <i>extracto_bancolombia_febrero_2025.pdf</i>) o que las fechas en el contenido sean legibles.
</div>""", unsafe_allow_html=True)

    # ── Comparación ──────────────────────────────────────────────────────────
    if not df_aux.empty:
        df_comp, df_solo_aux = comparar_documentos(df_banco, df_aux)
        # Garantizar columna ESTADO (fix KeyError)
        if 'ESTADO' not in df_comp.columns:
            df_comp['ESTADO'] = '❓ Sin clasificar'
        n_tot     = len(df_comp)
        rechazos  = df_comp[df_comp['ESTADO'] == '🔄 RECHAZO — CONFIRMAR']
        agrupados = df_comp[df_comp['ESTADO'].str.startswith('🔵 AGRUPADO', na=False)]
        # ── Fase D3: auto-aprendizaje NC post-reconciliacion ─────────────────
        if not df_comp.empty and 'DOC_AUXILIAR' in df_comp.columns:
            _nc_matches = df_comp[
                df_comp['DOC_AUXILIAR'].str.startswith('NC-', na=False)
            ]
            for _, _nr in _nc_matches.iterrows():
                registrar_aprendizaje_nc(
                    str(_nr.get('DESCRIPCION',   '') or ''),
                    str(_nr.get('DOC_AUXILIAR',  '') or ''),
                    str(_nr.get('CONCEPTO_AUX',  '') or ''),
                    str(_nr.get('METODO_MATCH',  'MONTO') or 'MONTO'),
                    _nr.get('VALOR_BANCO'), _nr.get('MONTO_AUXILIAR')
                )
        n_exac = (df_comp['ESTADO'] == '✅ COINCIDE EXACTO').sum()
        n_apr  = (df_comp['ESTADO'] == '🔶 COINCIDE APROX.').sum()
        n_sbco = (df_comp['ESTADO'] == '❌ SOLO EN BANCO').sum()
        n_rec  = (df_comp['ESTADO'] == '🔄 RECHAZO — CONFIRMAR').sum()
        n_agr  = df_comp['ESTADO'].str.startswith('🔵 AGRUPADO', na=False).sum()
        n_saux = len(df_solo_aux)
        pct_conc = (n_exac + n_apr) / max(n_tot, 1) * 100
        exactas = df_comp[df_comp['ESTADO'] == '✅ COINCIDE EXACTO']
        aprox   = df_comp[df_comp['ESTADO'] == '🔶 COINCIDE APROX.']
        s_banco = df_comp[df_comp['ESTADO'] == '❌ SOLO EN BANCO']
    else:
        df_comp = pd.DataFrame()
        df_solo_aux = pd.DataFrame()
        n_tot = n_exac = n_apr = n_sbco = n_saux = n_rec = n_agr = pct_conc = 0
        exactas = aprox = s_banco = rechazos = agrupados = pd.DataFrame()
    # ── Guardar análisis en historial ────────────────────────────────────────
    try:
        _per_det = _periodo_detectado
        _per_str = (f"{_MESES_ES[_per_det[1]-1]} {_per_det[0]}" if _per_det else "")
        guardar_historial({
            "fecha_hora"      : datetime.now().strftime("%Y-%m-%d %H:%M"),
            "archivo_banco"   : _nombre_b,
            "archivo_auxiliar": _nombre_a,
            "periodo"         : _per_str,
            "n_banco"         : len(df_banco),
            "n_aux"           : len(df_aux),
            "n_exactas"       : int(n_exac),
            "n_aprox"         : int(n_apr),
            "n_solo_banco"    : int(n_sbco),
            "n_solo_aux"      : int(n_saux),
            "tasa"            : float(pct_conc),
            "saldo_banco"     : float(sac  or 0),
            "saldo_aux"       : float(sf_a or 0),
            "diferencia_neta" : float((sac or 0) - (sf_a or 0)),
            "excel_path"      : "",
        })
    except Exception as e:
        log.error("[app] Error guardando historial: %s", e, exc_info=True)


    # ── Dashboard KPI ejecutivo animado ──────────────────────────────────────
    _dif_neta = float((sac or 0) - (sf_a or 0))
    _col_dif  = "#4CAF50" if abs(_dif_neta) < 1.0 else ("#FFC107" if abs(_dif_neta) < 50000 else "#F44336")
    _ico_dif  = "✅" if abs(_dif_neta) < 1.0 else ("⚠️" if abs(_dif_neta) < 50000 else "❌")
    _ico_pct, _lbl_pct, _ = _semaforo_conciliacion(pct_conc)
    _col_pct = "#4CAF50" if pct_conc >= 90 else ("#FFC107" if pct_conc >= 75 else "#F44336")

    # ── KPI valores calculados ────────────────────────────────────────────────
    _pct_fill    = min(pct_conc, 100)
    _exac_fill   = n_exac / max(n_tot, 1) * 100
    _sinc_fill   = (n_sbco + n_saux) / max(n_tot + n_saux, 1) * 100
    _cls_dif     = "rojo" if abs(_dif_neta) >= 50000 else ("ambar" if abs(_dif_neta) >= 1 else "verde")
    _cls_pct     = "verde" if pct_conc >= 90 else ("ambar" if pct_conc >= 75 else "rojo")
    _cls_sinc    = "rojo" if (n_sbco + n_saux) > 10 else ("ambar" if (n_sbco + n_saux) > 3 else "verde")
    _critico_dif = " critico" if abs(_dif_neta) >= 50000 else ""
    _critico_sinc= " critico" if (n_sbco + n_saux) > 10 else ""
    _delta_pct   = f"↑ Meta 90%" if pct_conc < 90 else "✓ Meta cumplida"
    _cls_delta   = "up" if pct_conc >= 90 else "down"
    # Gauge SVG de conciliación
    _gauge_pct   = min(pct_conc, 100)
    _gauge_angle = _gauge_pct / 100 * 180  # 0-180 grados
    _r = 38; _cx = 50; _cy = 50
    import math as _math
    _rad = _math.radians(180 - _gauge_angle)
    _gx  = _cx + _r * _math.cos(_rad)
    _gy  = _cy - _r * _math.sin(_rad)
    _arc_color = ("#22c55e" if pct_conc >= 90 else ("#f59e0b" if pct_conc >= 75 else "#ef4444"))
    _gauge_svg = f"""
<svg viewBox="0 0 100 60" width="90" height="54" style="display:block;margin:6px auto 0;">
  <path d="M12,50 A38,38 0 0,1 88,50" fill="none" stroke="rgba(255,255,255,.10)" stroke-width="8" stroke-linecap="round"/>
  <path d="M12,50 A38,38 0 0,1 {_gx:.2f},{_gy:.2f}" fill="none" stroke="{_arc_color}" stroke-width="8"
        stroke-linecap="round" style="filter:drop-shadow(0 0 4px {_arc_color}88)"/>
  <circle cx="{_gx:.2f}" cy="{_gy:.2f}" r="4" fill="{_arc_color}"/>
  <text x="50" y="54" text-anchor="middle" font-size="13" font-weight="900"
        fill="{_arc_color}" font-family="Segoe UI,sans-serif">{pct_conc:.0f}%</text>
</svg>"""

    st.markdown(f"""
<div class="kpi-grid" id="kpi-dashboard">
  <!-- KPI 1: Tasa conciliación con gauge -->
  <div class="kpi-card {_cls_pct}" style="--delay:.05s">
    <div class="kpi-icon">🎯</div>
    <div class="kpi-label">Tasa de Conciliación</div>
    {_gauge_svg}
    <div class="kpi-sub">{_ico_pct} {_lbl_pct}</div>
    <div class="kpi-delta {_cls_delta}">{_delta_pct}</div>
  </div>
  <!-- KPI 2: Exactas con barra animada -->
  <div class="kpi-card verde" style="--delay:.10s">
    <div class="kpi-icon">✅</div>
    <div class="kpi-label">Mov. Exactos</div>
    <div class="kpi-value" style="color:#4ade80;" data-count="{n_exac}">0</div>
    <div class="kpi-sub">de {n_tot} movs. banco</div>
    <div class="kpi-bar-bg">
      <div class="kpi-bar-fill" style="--fill:{_exac_fill:.1f}%;--delay:.3s"></div>
    </div>
    <div class="kpi-delta up">+{n_apr} aprox.</div>
  </div>
  <!-- KPI 3: Diferencia neta -->
  <div class="kpi-card {_cls_dif}{_critico_dif}" style="--delay:.15s">
    <div class="kpi-icon">⚖️</div>
    <div class="kpi-label">Diferencia Neta</div>
    <div class="kpi-value" style="color:{'#4ade80' if abs(_dif_neta)<1 else ('#fbbf24' if abs(_dif_neta)<50000 else '#f87171')};font-size:1.15rem;">
      {_ico_dif} $<span data-count-float="{abs(_dif_neta):.0f}">0</span>
    </div>
    <div class="kpi-sub">Banco vs Auxiliar · COP</div>
    <div class="kpi-delta {'up' if abs(_dif_neta)<1 else 'down'}">{'✓ Cuadrado' if abs(_dif_neta)<1 else 'Revisar diferencia'}</div>
  </div>
  <!-- KPI 4: Sin conciliar -->
  <div class="kpi-card morado{_critico_sinc}" style="--delay:.20s">
    <div class="kpi-icon">🔍</div>
    <div class="kpi-label">Sin Conciliar</div>
    <div class="kpi-value" style="color:#c084fc;" data-count="{n_sbco + n_saux}">0</div>
    <div class="kpi-sub">{n_sbco} solo banco · {n_saux} solo aux.</div>
    <div class="kpi-bar-bg">
      <div class="kpi-bar-fill" style="--fill:{_sinc_fill:.1f}%;--delay:.4s"></div>
    </div>
  </div>
  <!-- KPI 5: Saldo banco -->
  <div class="kpi-card azul" style="--delay:.25s">
    <div class="kpi-icon">🏦</div>
    <div class="kpi-label">Saldo Banco</div>
    <div class="kpi-value" style="color:#60a5fa;font-size:1.15rem;">${sac/1e6:,.3f}M</div>
    <div class="kpi-sub">Aux: ${sf_a/1e6:,.3f}M COP</div>
    <div class="kpi-delta {'up' if sac >= sf_a else 'down'}">
      {'↑' if sac >= sf_a else '↓'} Δ ${abs(sac - sf_a)/1e6:,.3f}M
    </div>
  </div>
</div>
<script>
(function(){{
  /* Counter animado para data-count */
  function animCount(el, target, duration, isFloat) {{
    var start = null;
    function step(ts) {{
      if (!start) start = ts;
      var prog = Math.min((ts - start) / duration, 1);
      var ease = 1 - Math.pow(1 - prog, 3);
      var val  = Math.round(ease * target);
      el.textContent = isFloat
        ? val.toLocaleString('es-CO')
        : val.toLocaleString('es-CO');
      if (prog < 1) requestAnimationFrame(step);
      else el.textContent = isFloat
        ? target.toLocaleString('es-CO')
        : target.toLocaleString('es-CO');
    }}
    requestAnimationFrame(step);
  }}
  /* Ripple en botones */
  function addRipple(btn) {{
    btn.addEventListener('click', function(e) {{
      var r = document.createElement('span');
      r.className = 'ripple';
      var rect = btn.getBoundingClientRect();
      var size = Math.max(rect.width, rect.height);
      r.style.cssText = 'width:'+size+'px;height:'+size+'px;left:'+(e.clientX-rect.left-size/2)+'px;top:'+(e.clientY-rect.top-size/2)+'px;position:absolute;border-radius:50%;background:rgba(255,255,255,.25);transform:scale(0);animation:rippleEffect .5s linear;pointer-events:none;';
      btn.appendChild(r);
      setTimeout(function(){{ r.remove(); }}, 600);
    }});
  }}
  /* Ejecutar al cargar */
  setTimeout(function() {{
    document.querySelectorAll('[data-count]').forEach(function(el) {{
      animCount(el, parseInt(el.dataset.count), 1200, false);
    }});
    document.querySelectorAll('[data-count-float]').forEach(function(el) {{
      animCount(el, parseFloat(el.dataset.countFloat), 1200, true);
    }});
    document.querySelectorAll('.stButton > button').forEach(addRipple);
  }}, 100);
}})();
</script>
""", unsafe_allow_html=True)

    # ── Función badge HTML para columna ESTADO ────────────────────────────
    def _badge_estado(estado: str) -> str:
        """Convierte texto de ESTADO en un badge HTML animado premium."""
        e = str(estado or "")
        if "EXACTO"   in e: cls, lbl = "estado-exacto",  "✅ EXACTO"
        elif "APROX"  in e: cls, lbl = "estado-aprox",   "🔶 APROX."
        elif "SOLO EN BANCO" in e: cls, lbl = "estado-solo", "❌ SOLO BANCO"
        elif "SOLO AUX" in e:      cls, lbl = "estado-solo", "❌ SOLO AUX"
        elif "AGRUPADO" in e:      cls, lbl = "estado-agrup",   "🔵 AGRUP."
        elif "RECHAZO"  in e:      cls, lbl = "estado-rechazo", "🔄 RECHAZO"
        else: cls, lbl = "estado-otros", e[:20]
        return f'<span class="estado-badge {cls}">{lbl}</span>'

    # ── Status bar flotante ───────────────────────────────────────────────
    _sb_dot  = "green" if pct_conc >= 90 else ("amber" if pct_conc >= 75 else "red")
    _sb_dot2 = "green" if abs(_dif_neta) < 1 else ("amber" if abs(_dif_neta) < 50000 else "red")
    _sb_dot3 = "green" if (n_sbco + n_saux) == 0 else ("amber" if (n_sbco + n_saux) <= 5 else "red")
    st.markdown(f"""
<div class="status-bar">
  <div class="sb-item">
    <span class="sb-dot {_sb_dot}"></span>
    Conciliación: <span>{pct_conc:.1f}%</span>
  </div>
  <div class="sb-item">
    <span class="sb-dot {_sb_dot2}"></span>
    Diferencia: <span>${abs(_dif_neta):,.0f} COP</span>
  </div>
  <div class="sb-item">
    <span class="sb-dot {_sb_dot3}"></span>
    Sin conciliar: <span>{n_sbco + n_saux} movs.</span>
  </div>
  <div class="sb-item" style="margin-left:auto;">
    <span class="sb-dot green"></span>
    <span style="color:#4ade80;">✓ Análisis completado</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Pestañas ──────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, \
    tab9, tab10, tab11, tab12, tab13, tab14 = st.tabs([
        "📊 Diagnóstico", "🏦 Extracto Banco", "📋 Auxiliar Contable",
        "🔗 Comparación", "📝 Diferencias", "⚖️ Conciliación Formal",
        "📈 Visualizaciones", "💾 Exportar",
        "🗂️ Partidas", "💸 Comisiones", "📑 DIAN / XML",
        "📚 PUC", "🤖 ML Predictor", "⚙️ Admin",
    ])

    with tab1:
        st.markdown("<div class='section-title'>🩺 Diagnóstico Contable Ejecutivo</div>", unsafe_allow_html=True)
        p_banco, cal_banco, adv_banco, fmt_banco, conf_banco = leg_banco
        p_aux,   cal_aux,   adv_aux,   fmt_aux,   conf_aux   = leg_aux

        # ── Calcular indicadores de riesgo ───────────────────────────────────
        _dif_neta_t1   = float((sac or 0) - (sf_a or 0))
        _val_sin_conc  = 0.0
        if not s_banco.empty and 'VALOR_BANCO' in s_banco.columns:
            _val_sin_conc += float(s_banco['VALOR_BANCO'].abs().sum())
        if not df_solo_aux.empty:
            _val_sin_conc += float(df_solo_aux.get('DEBITO',  pd.Series(dtype=float)).fillna(0).sum())
            _val_sin_conc += float(df_solo_aux.get('CREDITO', pd.Series(dtype=float)).fillna(0).sum())
        _dias_cierre = 0
        try:
            from datetime import datetime as _dt
            _hoy = _dt.now()
            import calendar as _cal
            _ult_dia = _cal.monthrange(_hoy.year, _hoy.month)[1]
            _dias_cierre = _ult_dia - _hoy.day
        except Exception:
            _dias_cierre = 0

        # ── Checklist semáforo contable ───────────────────────────────────────
        _checks = [
            ("🏦 Extracto bancario legible",        p_banco >= 80,   f"{p_banco:.0f}% confianza de lectura"),
            ("📋 Auxiliar contable legible",         p_aux   >= 80,   f"{p_aux:.0f}% confianza de lectura"),
            ("➕ Banco cuadra aritméticamente",      abs((sa + tab_s - tca_s) - sac) < 1,
                                                     "Saldo inicial + abonos – cargos = saldo final"),
            ("➕ Auxiliar cuadra aritméticamente",   abs((si_a + td_a - tc_a) - sf_a) < 1,
                                                     "Saldo inicial + débitos – créditos = saldo final"),
            ("🎯 Tasa conciliación ≥ 90%",           pct_conc >= 90,  f"{pct_conc:.1f}% de movimientos conciliados"),
            ("💰 Diferencia neta < $1 COP",          abs(_dif_neta_t1) < 1,
                                                     f"Diferencia banco vs auxiliar: {cop(_dif_neta_t1).strip()}"),
            ("⚠️ Valor sin conciliar < $100K",       _val_sin_conc < 100_000,
                                                     f"Partidas sin conciliar: ${_val_sin_conc:,.0f} COP"),
            ("📄 Sin advertencias de formato",       not adv_banco and not adv_aux,
                                                     "Sin alertas de parseo en ninguno de los archivos"),
        ]
        _n_ok  = sum(1 for _, v, _ in _checks if v)
        _n_tot_ch = len(_checks)
        _pct_ok = _n_ok / _n_tot_ch * 100

        # Semáforo general
        if _pct_ok >= 87.5:
            _st_color, _st_icon, _st_label = '#4CAF50', '🟢', 'LISTO PARA CIERRE'
        elif _pct_ok >= 62.5:
            _st_color, _st_icon, _st_label = '#FFC107', '🟡', 'REVISAR ANTES DE CIERRE'
        else:
            _st_color, _st_icon, _st_label = '#F44336', '🔴', 'ATENCIÓN — DIFERENCIAS CRÍTICAS'

        st.markdown(f"""
<div style='background:rgba(255,255,255,0.05);border-left:6px solid {_st_color};
     border-radius:12px;padding:16px 22px;margin-bottom:16px;'>
  <div style='font-size:1.1rem;font-weight:800;color:{_st_color};'>{_st_icon} {_st_label}</div>
  <div style='font-size:.88rem;opacity:.8;margin-top:4px;'>
    {_n_ok} de {_n_tot_ch} controles superados &nbsp;·&nbsp;
    Valor en riesgo: <b>${_val_sin_conc:,.0f} COP</b> &nbsp;·&nbsp;
    Días al cierre mensual: <b>{_dias_cierre}</b>
  </div>
</div>""", unsafe_allow_html=True)

        # Checklist visual
        st.markdown("<div class='section-title'>✅ Checklist de Control Contable</div>", unsafe_allow_html=True)
        _ch_c1, _ch_c2 = st.columns(2)
        for _i, (_label, _ok, _detalle) in enumerate(_checks):
            _ico_ch  = "✅" if _ok else "❌"
            _bg_ch   = "rgba(76,175,80,0.08)"  if _ok else "rgba(244,67,54,0.08)"
            _bdr_ch  = "#4CAF50" if _ok else "#F44336"
            _html_ch = f"""
<div style='background:{_bg_ch};border-left:3px solid {_bdr_ch};border-radius:8px;
     padding:9px 13px;margin:4px 0;'>
  <span style='font-weight:700;font-size:.87rem;'>{_ico_ch} {_label}</span><br>
  <span style='font-size:.79rem;opacity:.7;'>{_detalle}</span>
</div>"""
            if _i % 2 == 0:
                _ch_c1.markdown(_html_ch, unsafe_allow_html=True)
            else:
                _ch_c2.markdown(_html_ch, unsafe_allow_html=True)

        # ── Calidad de archivos ───────────────────────────────────────────────
        st.markdown("<div class='section-title'>📂 Calidad de Archivos Cargados</div>", unsafe_allow_html=True)
        _qa1, _qa2 = st.columns(2)
        with _qa1:
            badge_b = "badge-verde" if p_banco >= 90 else ("badge-naranja" if p_banco >= 70 else "badge-rojo")
            st.markdown(f"""
<div class='callout-info'>
  <b>🏦 Extracto Bancario</b><br>
  Formato: <b>{fmt_banco or "Desconocido"}</b> &nbsp;·&nbsp;
  Confianza: <span class='{badge_b}'>{conf_banco}%</span><br>
  Movimientos leídos: <b>{len(df_banco)}</b> &nbsp;·&nbsp;
  Legibilidad: <b>{p_banco:.0f}%</b>
  {"<br><small style='color:#F44336;'>⚠️ " + " | ".join(adv_banco[:2]) + "</small>" if adv_banco else ""}
</div>""", unsafe_allow_html=True)
        with _qa2:
            badge_a = "badge-verde" if p_aux >= 90 else ("badge-naranja" if p_aux >= 70 else "badge-rojo")
            st.markdown(f"""
<div class='callout-info'>
  <b>📋 Auxiliar Contable</b><br>
  Formato: <b>{fmt_aux or "Desconocido"}</b> &nbsp;·&nbsp;
  Confianza: <span class='{badge_a}'>{conf_aux}%</span><br>
  Asientos leídos: <b>{len(df_aux)}</b> &nbsp;·&nbsp;
  Legibilidad: <b>{p_aux:.0f}%</b>
  {"<br><small style='color:#F44336;'>⚠️ " + " | ".join(adv_aux[:2]) + "</small>" if adv_aux else ""}
</div>""", unsafe_allow_html=True)

        # ── Guía de acción según estado ───────────────────────────────────────
        if _pct_ok < 87.5:
            _acciones = []
            if abs((sa + tab_s - tca_s) - sac) >= 1:
                _acciones.append("Verificar el extracto con la entidad bancaria — puede tener una transacción incompleta.")
            if abs((si_a + td_a - tc_a) - sf_a) >= 1:
                _acciones.append("Revisar el auxiliar contable en el sistema ERP — posible asiento mal cuadrado.")
            if pct_conc < 90:
                _acciones.append(f"Conciliar las {n_sbco + n_saux} partidas pendientes (ver Tab 📝 Diferencias → clasificación automática).")
            if _val_sin_conc >= 100_000:
                _acciones.append(f"Regularizar ${_val_sin_conc:,.0f} COP sin conciliar antes del cierre mensual ({_dias_cierre} días).")
            if p_banco < 80 or p_aux < 80:
                _acciones.append("Exportar los archivos a formato CSV/Excel desde el sistema fuente para mejorar la lectura.")
            if _acciones:
                st.markdown("<div class='section-title'>🔧 Acciones Recomendadas</div>", unsafe_allow_html=True)
                for _idx, _acc in enumerate(_acciones, 1):
                    st.markdown(f"""
<div style='background:rgba(255,193,7,0.07);border-left:3px solid #FFC107;
     border-radius:8px;padding:8px 13px;margin:4px 0;font-size:.87rem;'>
  <b>{_idx}.</b> {_acc}
</div>""", unsafe_allow_html=True)

    with tab2:
        st.markdown("<div class='section-title'>🏦 Extracto Bancario</div>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Saldo Anterior", cop(sa))
        col2.metric("Total Abonos (+)", cop(tab_s))
        col3.metric("Total Cargos (−)", cop(tca_s))
        col4.metric("Saldo Final", cop(sac))

        dif_arit = (sa + tab_s - tca_s) - sac
        if abs(dif_arit) < 1:
            st.markdown(f"""
            <div class='callout-success'>
              <b>✅ El extracto cuadra aritméticamente.</b><br>
              {cop(sa)} + {cop(tab_s)} − {cop(tca_s)} = <b>{cop(sa+tab_s-tca_s)}</b>
              &nbsp;≈&nbsp; Saldo final declarado <b>{cop(sac)}</b>
              &nbsp;·&nbsp; Diferencia: <b>{cop(dif_arit)}</b>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class='callout-danger'>
              <b>⚠️ El extracto NO cuadra aritméticamente.</b><br>
              {cop(sa)} + {cop(tab_s)} − {cop(tca_s)} = <b>{cop(sa+tab_s-tca_s)}</b>
              &nbsp;vs&nbsp; Saldo declarado <b>{cop(sac)}</b>
              &nbsp;·&nbsp; <b>Diferencia: {cop(dif_arit)}</b>
            </div>""", unsafe_allow_html=True)

        # Análisis de anomalías
        st.markdown("<div class='section-title'>🔍 Análisis de Movimientos</div>", unsafe_allow_html=True)
        n_abonos = (df_banco['TIPO'] == 'ABONO').sum() if 'TIPO' in df_banco.columns else 0
        n_cargos = (df_banco['TIPO'] == 'CARGO').sum() if 'TIPO' in df_banco.columns else 0
        st.markdown(f"""
        <div class='callout-info'>
          Total movimientos: <b>{len(df_banco)}</b>
          &nbsp;·&nbsp; Abonos: <b>{n_abonos}</b>
          &nbsp;·&nbsp; Cargos: <b>{n_cargos}</b>
          &nbsp;·&nbsp; Promedio por movimiento: <b>{cop(df_banco['VALOR'].abs().mean() if not df_banco.empty else 0)}</b>
        </div>""", unsafe_allow_html=True)

        # Top 5 movimientos por valor absoluto
        if not df_banco.empty and 'VALOR' in df_banco.columns:
            top5 = df_banco.nlargest(5, df_banco['VALOR'].abs().name if hasattr(df_banco['VALOR'].abs(), 'name') else 'VALOR')
            try:
                top5 = df_banco.iloc[df_banco['VALOR'].abs().nlargest(5).index]
            except Exception:
                top5 = df_banco.head(5)
            with st.expander("📌 Top 5 movimientos de mayor valor", expanded=False):
                cols_show = [c for c in ['FECHA_RAW','DESCRIPCION','VALOR','SALDO','TIPO'] if c in df_banco.columns]
                st.dataframe(top5[cols_show], use_container_width=True)

        st.markdown("<div class='section-title'>📄 Detalle de Transacciones</div>", unsafe_allow_html=True)
        cols_banco = [c for c in ['FECHA_RAW','DESCRIPCION','VALOR','SALDO','TIPO'] if c in df_banco.columns]
        st.dataframe(df_banco[cols_banco], use_container_width=True, height=400)

    with tab3:
        st.markdown("<div class='section-title'>📋 Auxiliar Contable</div>", unsafe_allow_html=True)
        if not df_aux.empty:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Saldo Inicial", cop(si_a))
            col2.metric("Total Débitos", cop(td_a))
            col3.metric("Total Créditos", cop(tc_a))
            col4.metric("Saldo Final", cop(sf_a))

            dif_arit_aux = (si_a + td_a - tc_a) - sf_a
            if abs(dif_arit_aux) < 1:
                st.markdown(f"""
                <div class='callout-success'>
                  <b>✅ El auxiliar cuadra aritméticamente.</b><br>
                  {cop(si_a)} + {cop(td_a)} − {cop(tc_a)} = <b>{cop(si_a+td_a-tc_a)}</b>
                  &nbsp;≈&nbsp; Saldo final declarado <b>{cop(sf_a)}</b>
                  &nbsp;·&nbsp; Diferencia: <b>{cop(dif_arit_aux)}</b>
                </div>""", unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class='callout-danger'>
                  <b>⚠️ El auxiliar NO cuadra aritméticamente.</b><br>
                  Diferencia de <b>{cop(dif_arit_aux)}</b> entre saldo calculado y declarado.
                  Verifique asientos de apertura o cierres de período.
                </div>""", unsafe_allow_html=True)

            # Desglose por tipo de documento
            deb_df = df_aux[df_aux['DEBITO'].notna()]
            cre_df = df_aux[df_aux['CREDITO'].notna()]
            des_df = df_aux[df_aux['COLUMNA'] == 'DESCONOCIDO'] if 'COLUMNA' in df_aux.columns else pd.DataFrame()

            st.markdown("<div class='section-title'>📊 Desglose por Tipo de Asiento</div>", unsafe_allow_html=True)
            ca, cb, cc = st.columns(3)
            ca.metric("Asientos DÉBITO", len(deb_df), f"{cop(deb_df['DEBITO'].sum())}")
            cb.metric("Asientos CRÉDITO", len(cre_df), f"{cop(cre_df['CREDITO'].sum())}")
            cc.metric("Sin clasificar", len(des_df))

            if len(des_df) > 0:
                st.markdown(f"""
                <div class='callout-warning'>
                  <b>⚠️ {len(des_df)} asientos sin clasificar (columna DESCONOCIDO).</b><br>
                  Estos asientos no pudieron ser identificados como DÉBITO ni CRÉDITO.
                  Pueden afectar el cálculo de la conciliación.
                </div>""", unsafe_allow_html=True)

            # Desglose por tipo de comprobante (primeras 2 letras del documento)
            if 'DOCUMENTO' in df_aux.columns:
                tipo_doc = df_aux['DOCUMENTO'].str[:2].value_counts().head(8)
                if not tipo_doc.empty:
                    with st.expander("📌 Distribución por tipo de comprobante", expanded=False):
                        for prefijo, cnt in tipo_doc.items():
                            st.markdown(f"&nbsp;&nbsp;<span class='badge-azul'>{prefijo}</span> — <b>{cnt}</b> asientos", unsafe_allow_html=True)

            st.markdown("<div class='section-title'>📄 Detalle de Asientos</div>", unsafe_allow_html=True)
            cols_aux = [c for c in ['DOCUMENTO','FECHA_RAW','CONCEPTO','DEBITO','CREDITO','COLUMNA'] if c in df_aux.columns]
            st.dataframe(df_aux[cols_aux], use_container_width=True, height=400)
        else:
            st.markdown("""
            <div class='callout-danger'>
              <b>❌ No se extrajeron asientos del auxiliar contable.</b><br>
              Verifique que el archivo tiene las columnas correctas: Documento, Fecha, Concepto, Débito, Crédito.
              Si es un PDF, active OCR o exporte a CSV desde el sistema contable.
            </div>""", unsafe_allow_html=True)

    with tab4:
        st.markdown("<div class='section-title'>🔗 Comparación Banco ↔ Auxiliar</div>", unsafe_allow_html=True)
        if df_aux.empty:
            st.markdown("<div class='callout-warning'>⚠️ Sin datos del auxiliar para comparar.</div>", unsafe_allow_html=True)
        else:
            c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
            c1.metric("Total Analizados", n_tot)
            c2.metric("✅ Exactos",    n_exac, f"{n_exac/max(n_tot,1)*100:.0f}%")
            c3.metric("🔶 Aprox.",     n_apr,  f"{n_apr/max(n_tot,1)*100:.0f}%")
            c4.metric("🔵 Agrupados",  n_agr,  "N:1 NC" if n_agr else "—")
            c5.metric("🔄 Rechazos",   n_rec,  "Confirmar" if n_rec else "—")
            c6.metric("❌ Solo Banco", n_sbco, f"{n_sbco/max(n_tot,1)*100:.0f}%")
            c7.metric("📋 Solo Aux.",  n_saux)

            ico, lbl, cls = _semaforo_conciliacion(pct_conc)
            st.progress(pct_conc / 100, text=f"{ico} Tasa de conciliación: {pct_conc:.1f}% — {lbl}")
            st.markdown(f"""
            <div class='callout-{"success" if pct_conc>=90 else ("warning" if pct_conc>=75 else "danger")}'>
              <b>{ico} Conciliación {lbl} — {pct_conc:.1f}%</b><br>
              {"Más del 90% de los movimientos tienen correspondencia en el auxiliar. Excelente control contable." if pct_conc>=90
               else ("Entre 75% y 90% de los movimientos conciliados. Hay diferencias puntuales que requieren revisión." if pct_conc>=75
               else "Menos del 75% de los movimientos conciliados. Se requiere revisión detallada del auxiliar contable.")}
            </div>""", unsafe_allow_html=True)
            if n_agr > 0:
                st.markdown(f"""
                <div class='callout-info'>
                  <b>🔵 {int(n_agr)} cargo(s) bancarios agrupados y vinculados a 1 NC del auxiliar (N:1).</b><br>
                  El banco los cobró individualmente; el contador los registró como una sola Nota Contable.
                  Ver detalle en <b>📝 Diferencias</b> → sección <i>Agrupados N:1</i>.
                </div>""", unsafe_allow_html=True)
            if n_rec > 0:
                st.markdown(f"""
                <div class='callout-warning'>
                  <b>🔄 {n_rec} cargo(s) bancario(s) posiblemente vinculados a notas contables de rechazo/devolución.</b><br>
                  El sistema detectó NC- con montos similares (±3%). Revíselos en la pestaña
                  <b>📝 Diferencias</b> → sección <i>Rechazos / Devoluciones — Confirmar</i>.
                </div>""", unsafe_allow_html=True)

            # ── KPI: Valor en riesgo ─────────────────────────────────────────
            _val_riesgo_b = float(s_banco['VALOR_BANCO'].abs().sum()) if not s_banco.empty else 0.0
            _val_riesgo_a = 0.0
            if not df_solo_aux.empty:
                _deb_sa = df_solo_aux['DEBITO'].fillna(0).sum()  if 'DEBITO'  in df_solo_aux.columns else 0
                _cre_sa = df_solo_aux['CREDITO'].fillna(0).sum() if 'CREDITO' in df_solo_aux.columns else 0
                _val_riesgo_a = abs(_deb_sa) + abs(_cre_sa)
            _val_riesgo_total = _val_riesgo_b + _val_riesgo_a

            _col_riesgo = "#4CAF50" if _val_riesgo_total < 100_000 else ("#FFC107" if _val_riesgo_total < 1_000_000 else "#F44336")
            st.markdown(f"""
<div style='background:rgba(244,67,54,0.07);border-left:4px solid {_col_riesgo};
     border-radius:10px;padding:14px 20px;margin:12px 0;'>
  <b>⚠️ Valor en riesgo (sin conciliar)</b><br>
  Banco sin auxiliar: <b style='color:#F44336;'>${_val_riesgo_b:,.0f} COP</b> &nbsp;|&nbsp;
  Auxiliar sin banco: <b style='color:#2196F3;'>${_val_riesgo_a:,.0f} COP</b> &nbsp;|&nbsp;
  <b style='color:{_col_riesgo};'>Total: ${_val_riesgo_total:,.0f} COP</b><br>
  <small style='opacity:.7;'>Este monto puede impactar el estado de tesorería si no se regulariza antes del cierre.</small>
</div>""", unsafe_allow_html=True)

            # ── Tasa de conciliación por tipo de comprobante ─────────────────
            if 'DOCUMENTO' in df_aux.columns and 'ESTADO' in df_comp.columns:
                st.markdown("<div class='section-title'>📊 Tasa de Conciliación por Tipo de Comprobante</div>",
                            unsafe_allow_html=True)

                # Extraer prefijo del documento del auxiliar y cruzar con df_comp
                _aux_tipos = df_aux.copy()
                _aux_tipos['_PREFIJO'] = _aux_tipos['DOCUMENTO'].astype(str).str[:2].str.upper()
                # Mapeo de prefijo a nombre legible
                _TIPO_NOMBRE = {
                    'NC': 'NC — Nota Crédito',
                    'CE': 'CE — Comp. Egreso',
                    'CG': 'CG — Comp. Caja',
                    'CO': 'CON — Concepto',
                    'EG': 'EG — Egreso',
                    'RC': 'RC — Recibo Caja',
                    'NI': 'NI — Nota Interna',
                    'ND': 'ND — Nota Débito',
                }
                _resumen_tipos = []
                _prefijos = _aux_tipos['_PREFIJO'].unique()
                for _pref in sorted(_prefijos):
                    if not _pref or _pref == 'NA':
                        continue
                    _docs_pref = _aux_tipos[_aux_tipos['_PREFIJO'] == _pref]['DOCUMENTO'].tolist()
                    # Buscar cuántos de estos documentos están en df_comp como conciliados
                    if 'DOC_AUX' in df_comp.columns:
                        _rows_pref = df_comp[df_comp['DOC_AUX'].astype(str).str[:2].str.upper() == _pref]
                    elif 'DOCUMENTO' in df_comp.columns:
                        _rows_pref = df_comp[df_comp['DOCUMENTO'].astype(str).str[:2].str.upper() == _pref]
                    else:
                        _rows_pref = pd.DataFrame()

                    _n_pref = max(len(_docs_pref), 1)
                    if not _rows_pref.empty and 'ESTADO' in _rows_pref.columns:
                        _concil = _rows_pref['ESTADO'].str.contains('COINCIDE', na=False).sum()
                        _tasa_p = min(_concil / _n_pref * 100, 100)
                        _val_p  = 0.0
                        if 'DEBITO' in _aux_tipos.columns:
                            _val_p += _aux_tipos[_aux_tipos['_PREFIJO']==_pref]['DEBITO'].fillna(0).sum()
                        if 'CREDITO' in _aux_tipos.columns:
                            _val_p += _aux_tipos[_aux_tipos['_PREFIJO']==_pref]['CREDITO'].fillna(0).sum()
                    else:
                        _concil, _tasa_p = 0, 0.0
                        _val_p = 0.0

                    _resumen_tipos.append({
                        'Tipo': _TIPO_NOMBRE.get(_pref, f'{_pref} — Comprobante'),
                        'N Total': _n_pref,
                        'Conciliados': _concil,
                        'Tasa (%)': round(_tasa_p, 1),
                        'Valor COP': _val_p,
                    })

                if _resumen_tipos:
                    _df_tipos = pd.DataFrame(_resumen_tipos).sort_values('Tasa (%)', ascending=True)
                    # Tabla visual con barras de color
                    for _, _rt in _df_tipos.iterrows():
                        _tp = float(_rt['Tasa (%)'])
                        _bc = '#4CAF50' if _tp >= 90 else ('#FFC107' if _tp >= 70 else '#F44336')
                        _bw = max(int(_tp), 2)
                        st.markdown(f"""
<div style='display:flex;align-items:center;gap:12px;margin:4px 0;
     background:rgba(255,255,255,0.03);border-radius:8px;padding:8px 14px;'>
  <div style='width:180px;font-size:.85rem;font-weight:600;white-space:nowrap;'>{_rt['Tipo']}</div>
  <div style='flex:1;background:rgba(255,255,255,0.07);border-radius:4px;height:10px;overflow:hidden;'>
    <div style='width:{_bw}%;height:100%;background:{_bc};border-radius:4px;'></div>
  </div>
  <div style='width:55px;text-align:right;font-size:.85rem;font-weight:700;color:{_bc};'>{_tp:.0f}%</div>
  <div style='width:40px;text-align:right;font-size:.8rem;opacity:.6;'>{int(_rt["N Total"])} mov</div>
  <div style='width:130px;text-align:right;font-size:.8rem;opacity:.7;'>${abs(_rt["Valor COP"]):,.0f}</div>
</div>""", unsafe_allow_html=True)

            st.markdown("<div class='section-title'>📋 Tabla Completa de Comparación</div>", unsafe_allow_html=True)
            # Tabla con badges HTML en columna ESTADO
            _df_comp_display = df_comp.copy()
            if 'ESTADO' in _df_comp_display.columns:
                _df_comp_display['ESTADO'] = _df_comp_display['ESTADO'].apply(_badge_estado)
                _cols_comp = [c for c in ['FECHA_BANCO','DESCRIPCION','VALOR_BANCO',
                              'ESTADO','DOC_AUXILIAR','MONTO_AUXILIAR','DIFERENCIA']
                              if c in _df_comp_display.columns]
                _html_comp = _df_comp_display[_cols_comp].to_html(
                    escape=False, index=False,
                    classes="tabla-premium",
                    border=0
                )
                st.markdown(f"""
<style>
.tabla-premium {{ width:100%; border-collapse:collapse; font-size:.78rem; }}
.tabla-premium th {{
  background:rgba(59,130,246,.15); color:#93c5fd;
  font-weight:800; letter-spacing:.05em; text-transform:uppercase;
  padding:8px 12px; border-bottom:2px solid rgba(59,130,246,.3);
  font-size:.68rem;
}}
.tabla-premium td {{
  padding:7px 12px; border-bottom:1px solid rgba(255,255,255,.05);
  vertical-align:middle;
}}
.tabla-premium tr:hover td {{ background:rgba(59,130,246,.06); }}
.tabla-premium tr:nth-child(even) td {{ background:rgba(255,255,255,.02); }}
</style>
<div style="overflow-x:auto;border-radius:12px;border:1px solid rgba(59,130,246,.2);max-height:480px;overflow-y:auto;">
{_html_comp}
</div>""", unsafe_allow_html=True)
            else:
                st.dataframe(df_comp, use_container_width=True, height=450)

    with tab5:
        st.markdown("<div class='section-title'>📝 Reporte Detallado de Diferencias</div>", unsafe_allow_html=True)
        if df_aux.empty:
            st.markdown("<div class='callout-warning'>⚠️ Sin datos del auxiliar para comparar.</div>", unsafe_allow_html=True)
        else:
            # ── AUTO-CLASIFICACIÓN INTELIGENTE DE DIFERENCIAS ────────────────
            if not s_banco.empty or not df_solo_aux.empty:
                st.markdown("<div class='section-title'>🔬 Clasificación Automática de Partidas sin Conciliar</div>",
                            unsafe_allow_html=True)

                # ── Clasificar movimientos SOLO EN BANCO por tipo probable ────
                def _clasificar_mov_banco(row):
                    """Clasifica un movimiento bancario sin auxiliar por tipo probable."""
                    desc  = str(row.get('DESCRIPCION', '') or '').upper()
                    valor = float(row.get('VALOR_BANCO', 0) or 0)
                    fecha_raw = str(row.get('FECHA_BANCO', '') or row.get('FECHA_RAW', '') or '')
                    # Detectar día del mes
                    try:
                        dia = int(fecha_raw.split('/')[0]) if '/' in fecha_raw else int(fecha_raw.split('-')[2])
                    except Exception:
                        dia = 15

                    # Reglas de clasificación por prioridad
                    if any(x in desc for x in ['GMF','4X1000','4 X 1000','IMPTO GOB']):
                        return '💸 GMF / Impuesto 4×1000', '#E65100', 'Crear NC débito cuenta 5305 GMF'
                    if any(x in desc for x in ['COMISION','COMISIÓN','MANEJO','CUOTA ADMIN']):
                        return '🏦 Comisión bancaria', '#1565C0', 'Crear NC débito cuenta 5305 Comisiones bancarias'
                    if any(x in desc for x in ['INTERÉS','INTERES','RENDIMIENTO']) and valor > 0:
                        return '📈 Rendimiento financiero', '#2E7D32', 'Crear NC crédito cuenta 4205 Rendimientos'
                    if any(x in desc for x in ['INTERÉS','INTERES']) and valor < 0:
                        return '💳 Interés débito', '#C62828', 'Crear NC débito cuenta 5305 Intereses'
                    if any(x in desc for x in ['CHEQUERA','TALONARIO','CHEQUERO']):
                        return '🏦 Chequera / Talonario', '#1565C0', 'Crear NC débito cuenta 5305 Servicios bancarios'
                    if any(x in desc for x in ['SEGUROS','SEGURO','POLIZA','PÓLIZA']):
                        return '🛡️ Seguro', '#6A1B9A', 'Crear NC débito cuenta 5310 Seguros'
                    if any(x in desc for x in ['PSE','NEQUI','DAVIPLATA','CORRESPONSAL']):
                        return '📱 Comisión transacción digital', '#1565C0', 'Crear NC débito cuenta 5305 Comisiones digitales'
                    if valor > 0 and dia >= 28:
                        return '⏳ Posible depósito en tránsito', '#F57F17', 'Verificar si fue causado en el período siguiente'
                    if valor > 0:
                        return '🟢 NC crédito no contabilizada', '#2E7D32', 'Buscar la NC en el auxiliar o crear si no existe'
                    if valor < 0:
                        return '🔴 ND débito no contabilizada', '#C62828', 'Crear NC débito con el cargo bancario'
                    return '❓ Sin clasificar — revisar', '#888', 'Revisar manualmente con el extracto'

                def _clasificar_asiento_aux(row):
                    """Clasifica un asiento contable sin banco por tipo probable."""
                    doc     = str(row.get('DOCUMENTO', '') or '').upper()
                    concepto= str(row.get('CONCEPTO', '') or '').upper()
                    deb     = float(row.get('DEBITO', 0) or 0)
                    cre     = float(row.get('CREDITO', 0) or 0)
                    prefijo = doc[:2] if len(doc) >= 2 else 'XX'

                    if prefijo in ('CE','EG','CG') and deb > 0:
                        return '🔲 Cheque girado / pago pendiente de cobro', '#E65100', 'El banco aún no ha debitado. Partida en tránsito — NO crear asiento'
                    if prefijo == 'NC' and cre > 0:
                        return '📋 NC crédito pendiente banco', '#1565C0', 'El banco aún no ha abonado. Verificar en próximo extracto'
                    if prefijo in ('CE','CG') and cre > 0:
                        return '📋 CE crédito — posible ajuste', '#6A1B9A', 'Revisar si es ajuste interno o cobro de cartera'
                    if any(x in concepto for x in ['AJUSTE','RECLASIF','RECLASIFICACION']):
                        return '⚙️ Asiento de ajuste interno', '#888', 'No genera movimiento bancario — es correcto que no esté en banco'
                    if any(x in concepto for x in ['PROVISION','PROVISIÓN','DEPRECIACION','DEPRECIACIÓN']):
                        return '📊 Provisión / Depreciación', '#888', 'Asiento de ajuste contable — no genera movimiento bancario'
                    if deb > 0:
                        return '❓ Débito sin banco — revisar', '#C62828', 'Verificar si el pago fue efectuado o está pendiente'
                    return '❓ Asiento sin clasificar', '#888', 'Revisar manualmente'

                # Aplicar clasificación
                _clases_banco = []
                if not s_banco.empty:
                    for _, _row_sb in s_banco.iterrows():
                        _tipo, _color, _accion = _clasificar_mov_banco(_row_sb)
                        _clases_banco.append({
                            'Tipo': _tipo, 'Color': _color, 'Acción': _accion,
                            'Valor': float(_row_sb.get('VALOR_BANCO', 0) or 0),
                        })

                _clases_aux = []
                if not df_solo_aux.empty:
                    for _, _row_sa in df_solo_aux.iterrows():
                        _tipo, _color, _accion = _clasificar_asiento_aux(_row_sa)
                        _deb = float(_row_sa.get('DEBITO', 0) or 0)
                        _cre = float(_row_sa.get('CREDITO', 0) or 0)
                        _clases_aux.append({
                            'Tipo': _tipo, 'Color': _color, 'Acción': _accion,
                            'Valor': _deb if _deb else _cre,
                        })

                # Resumen por categoría
                _res_banco = {}
                for _cl in _clases_banco:
                    _k = _cl['Tipo']
                    if _k not in _res_banco:
                        _res_banco[_k] = {'n': 0, 'total': 0.0, 'color': _cl['Color'], 'accion': _cl['Acción']}
                    _res_banco[_k]['n'] += 1
                    _res_banco[_k]['total'] += _cl['Valor']

                _res_aux = {}
                for _cl in _clases_aux:
                    _k = _cl['Tipo']
                    if _k not in _res_aux:
                        _res_aux[_k] = {'n': 0, 'total': 0.0, 'color': _cl['Color'], 'accion': _cl['Acción']}
                    _res_aux[_k]['n'] += 1
                    _res_aux[_k]['total'] += _cl['Valor']

                # Mostrar resúmenes
                _cc1, _cc2 = st.columns(2)
                with _cc1:
                    st.markdown("**🏦 Movimientos banco sin auxiliar — clasificados**")
                    if _res_banco:
                        for _kt, _kv in _res_banco.items():
                            st.markdown(f"""
<div style='background:rgba(255,255,255,0.04);border-left:3px solid {_kv["color"]};
     border-radius:8px;padding:9px 13px;margin:5px 0;'>
  <b style='font-size:.88rem;'>{_kt}</b><br>
  <span style='font-size:.82rem;opacity:.75;'>{_kv["n"]} movimiento(s) &nbsp;·&nbsp;
  ${abs(_kv["total"]):,.0f} COP</span><br>
  <span style='font-size:.79rem;color:{_kv["color"]};'>→ {_kv["accion"]}</span>
</div>""", unsafe_allow_html=True)
                        _total_riesgo_b = sum(abs(v['total']) for v in _res_banco.values())
                        st.markdown(f"<small style='opacity:.6;'>Total sin conciliar banco: **${_total_riesgo_b:,.0f} COP**</small>",
                                    unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='callout-success'>✅ Sin movimientos bancarios pendientes.</div>",
                                    unsafe_allow_html=True)

                with _cc2:
                    st.markdown("**📋 Asientos auxiliar sin banco — clasificados**")
                    if _res_aux:
                        for _kt, _kv in _res_aux.items():
                            st.markdown(f"""
<div style='background:rgba(255,255,255,0.04);border-left:3px solid {_kv["color"]};
     border-radius:8px;padding:9px 13px;margin:5px 0;'>
  <b style='font-size:.88rem;'>{_kt}</b><br>
  <span style='font-size:.82rem;opacity:.75;'>{_kv["n"]} asiento(s) &nbsp;·&nbsp;
  ${abs(_kv["total"]):,.0f} COP</span><br>
  <span style='font-size:.79rem;color:{_kv["color"]};'>→ {_kv["accion"]}</span>
</div>""", unsafe_allow_html=True)
                        _total_riesgo_a = sum(abs(v['total']) for v in _res_aux.values())
                        st.markdown(f"<small style='opacity:.6;'>Total sin conciliar auxiliar: **${_total_riesgo_a:,.0f} COP**</small>",
                                    unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='callout-success'>✅ Todos los asientos tienen movimiento bancario.</div>",
                                    unsafe_allow_html=True)

                st.markdown("---")

            # ── Calcular desglose abonos / cargos para exactas ───────────────
            if not exactas.empty:
                _ex_ab = exactas[exactas['VALOR_BANCO'] > 0]
                _ex_ca = exactas[exactas['VALOR_BANCO'] < 0]
                _ex_bruto = exactas['VALOR_BANCO'].abs().sum()
                _ex_titulo = (f"✅ Coincidencias Exactas — {len(exactas)} mov. "
                              f"· Bruto: ${_ex_bruto:,.0f} COP")
            else:
                _ex_titulo = "✅ Coincidencias Exactas — 0 movimientos"
                _ex_ab = _ex_ca = pd.DataFrame()
                _ex_bruto = 0

            with st.expander(_ex_titulo, expanded=False):
                if exactas.empty:
                    st.markdown("<div class='callout-warning'>Sin coincidencias exactas.</div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class='callout-success'>
                      <b>{len(exactas)} movimientos conciliados exactamente</b>
                      por valor y tipo de transacción.<br><br>
                      &nbsp;&nbsp;
                      <span class='badge-verde'>Abonos (+)</span>
                      &nbsp; {len(_ex_ab)} transacciones &nbsp;·&nbsp;
                      Total: <b>${_ex_ab['VALOR_BANCO'].sum():,.0f}</b><br>
                      &nbsp;&nbsp;
                      <span class='badge-rojo'>Cargos (−)</span>
                      &nbsp; {len(_ex_ca)} transacciones &nbsp;·&nbsp;
                      Total: <b>${abs(_ex_ca['VALOR_BANCO'].sum()):,.0f}</b><br><br>
                      Valor bruto total movido: <b>${_ex_bruto:,.0f} COP</b><br>
                      <small style='opacity:.7;'>
                        El valor bruto es la suma de valores absolutos (abonos + cargos).
                        El neto algebraico (abonos − cargos) es
                        ${exactas['VALOR_BANCO'].sum():,.0f} —
                        es normal que sea negativo si los cargos superan los abonos en el período.
                      </small>
                    </div>""", unsafe_allow_html=True)
                    cols_e = [c for c in ['FECHA_BANCO','TIPO_MOV','VALOR_BANCO','DOC_AUXILIAR','MONTO_AUXILIAR'] if c in exactas.columns]
                    st.dataframe(exactas[cols_e].head(100), use_container_width=True)
                    if len(exactas) > 100:
                        st.caption(f"Mostrando primeros 100 de {len(exactas)}. Descargue el Excel para ver todos.")

            with st.expander(f"🔶 Coincidencias Aproximadas — {len(aprox)} movimientos", expanded=False):
                if aprox.empty:
                    st.markdown("<div class='callout-success'>Sin diferencias aproximadas.</div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class='callout-warning'>
                      <b>{len(aprox)} movimientos con diferencias menores</b> (mismo período, valor cercano).<br>
                      Revise si existen redondeos, diferencias de centavos o asientos de ajuste.
                    </div>""", unsafe_allow_html=True)
                    cols_a = [c for c in ['FECHA_BANCO','TIPO_MOV','VALOR_BANCO','MONTO_AUXILIAR','DIFERENCIA','DOC_AUXILIAR'] if c in aprox.columns]
                    st.dataframe(aprox[cols_a], use_container_width=True)

            # ── SECCIÓN AGRUPADOS N:1 ────────────────────────────────────────
            _bruto_agr = agrupados['VALOR_BANCO'].abs().sum() if not agrupados.empty else 0
            with st.expander(
                f"🔵 Cargos Agrupados N:1 — {int(n_agr)} cargos bancarios → NC únicas · ${_bruto_agr:,.0f} COP",
                expanded=bool(n_agr > 0)
            ):
                if agrupados.empty:
                    st.markdown("<div class='callout-success'>✅ Sin cargos agrupados este período.</div>",
                                unsafe_allow_html=True)
                else:
                    # Resumir por NC vinculada
                    _agr_grupos = agrupados.groupby('DOC_AUXILIAR').agg(
                        N_cargos=('VALOR_BANCO', 'count'),
                        Suma_banco=('VALOR_BANCO', lambda x: x.abs().sum()),
                        NC_concepto=('CONCEPTO_AUX', 'first'),
                        Fecha_NC=('FECHA_AUXILIAR', 'first'),
                    ).reset_index()
                    st.markdown(f"""
                    <div class='callout-info'>
                      <b>🔵 {int(n_agr)} cargo(s) bancarios corresponden a {len(_agr_grupos)} NC del auxiliar.</b><br>
                      El banco cobra individualmente (por cada transacción) y el contador registra
                      una sola NC por el total. El sistema los vinculó automáticamente con tolerancia ±1%.<br><br>
                      <b>¿Requieren acción?</b> No — ya están registrados como nota contable.
                      Solo verifique que la NC del auxiliar esté correctamente fechada.
                    </div>""", unsafe_allow_html=True)
                    st.markdown("**Resumen por NC:**")
                    st.dataframe(_agr_grupos.rename(columns={
                        'DOC_AUXILIAR': 'NC Auxiliar',
                        'N_cargos'    : 'N cargos banco',
                        'Suma_banco'  : 'Total banco ($)',
                        'NC_concepto' : 'Concepto NC',
                        'Fecha_NC'    : 'Fecha NC',
                    }), use_container_width=True)
                    st.markdown("**Detalle de cargos individuales:**")
                    _cols_agr = [c for c in ['FECHA_BANCO','DESCRIPCION','VALOR_BANCO',
                                             'DOC_AUXILIAR','CONCEPTO_AUX','CONFIANZA','ESTADO']
                                 if c in agrupados.columns]
                    st.dataframe(agrupados[_cols_agr], use_container_width=True)

            # ── SECCIÓN RECHAZOS / DEVOLUCIONES ──────────────────────────────
            _bruto_rec = rechazos['VALOR_BANCO'].abs().sum() if not rechazos.empty else 0
            with st.expander(
                f"🔄 Rechazos / Devoluciones — CONFIRMAR — {int(n_rec)} trans. · ${_bruto_rec:,.0f} COP",
                expanded=bool(n_rec > 0)
            ):
                if rechazos.empty:
                    st.markdown("<div class='callout-success'>✅ Sin cargos rechazados detectados este período.</div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class='callout-warning'>
                      <b>🔄 {n_rec} cargo(s) bancario(s) posiblemente vinculados a NC por rechazo o devolución.</b><br>
                      El sistema los detectó con tolerancia de monto ±3% (para cubrir comisiones bancarias por rechazo).
                      <b>Requieren verificación manual</b> antes de darlos por conciliados.<br><br>
                      <b>Qué hacer:</b> Compare cada fila — si el cargo bancario y la NC corresponden al mismo evento,
                      el contador confirma que ya está registrado. Si no corresponden, cree la NC faltante en el sistema contable.
                    </div>""", unsafe_allow_html=True)
                    _cols_rec = [c for c in ['FECHA_BANCO','DESCRIPCION','VALOR_BANCO',
                                             'DOC_AUXILIAR','FECHA_AUXILIAR','CONCEPTO_AUX',
                                             'MONTO_AUXILIAR','DIFERENCIA','CONFIANZA'] if c in rechazos.columns]
                    st.dataframe(rechazos[_cols_rec], use_container_width=True)
                    st.markdown("""
                    <div class='callout-accion'>
                      <b>📌 ACCIÓN:</b> Si confirma la coincidencia → el cargo queda conciliado (no crear nueva NC).
                      Si NO corresponde → crear NC en SIIGO con el valor exacto del cargo bancario.
                    </div>""", unsafe_allow_html=True)

            # ── SECCIÓN CRÍTICA: Movimientos sin registro contable ──────────
            n_sb = len(s_banco)
            bruto_sb = s_banco['VALOR_BANCO'].abs().sum() if not s_banco.empty else 0
            with st.expander(f"❌ Movimientos Bancarios SIN Registro Contable — {int(n_sb)} trans. · ${bruto_sb:,.0f} COP", expanded=bool(n_sb > 0)):
                if s_banco.empty:
                    st.markdown("<div class='callout-success'>✅ Todos los movimientos tienen asiento contable.</div>", unsafe_allow_html=True)
                else:
                    abonos_sb = s_banco[s_banco['VALOR_BANCO'] > 0]
                    cargos_sb = s_banco[s_banco['VALOR_BANCO'] < 0]
                    st.markdown(f"""
                    <div class='callout-danger'>
                      <b>❌ {n_sb} movimientos bancarios no tienen asiento en el auxiliar.</b><br>
                      Abonos sin asiento: <b>{_cop_limpio(abonos_sb['VALOR_BANCO'].sum())}</b> ({len(abonos_sb)} trans.)
                      &nbsp;·&nbsp;
                      Cargos sin asiento: <b>{_cop_limpio(cargos_sb['VALOR_BANCO'].sum())}</b> ({len(cargos_sb)} trans.)<br>
                      <b>Valor bruto no registrado: ${bruto_sb:,.0f} COP</b>
                    </div>""", unsafe_allow_html=True)

                    # ── DIAGNÓSTICO POR GRUPO ──────────────────────────────
                    st.markdown("<div class='section-title'>📊 Diagnóstico por Grupo de Transacciones</div>", unsafe_allow_html=True)
                    st.caption("Cada grupo de transacciones con la misma descripción se analiza en conjunto. "
                               "La columna 'Candidato en Auxiliar' muestra el asiento más cercano encontrado.")

                    def _badge_grupo(diff_pct, razon):
                        if diff_pct is None or str(razon) == 'Sin asiento con monto similar — registrar NC':
                            return "🔴 Sin NC — Crear asiento"
                        elif diff_pct <= 0.5:
                            return "🟡 Monto OK — Revisar concepto/doc"
                        elif diff_pct <= 5:
                            return "🟠 Diferencia monto leve"
                        else:
                            return "🔴 Sin NC similar"

                    # Agrupar por descripción
                    _sb_grp = s_banco.copy()
                    _sb_grp['_desc_key'] = _sb_grp['DESCRIPCION'].str[:50]
                    grupos_diag = []
                    for _dk, _dg in _sb_grp.groupby('_desc_key', sort=False):
                        _total   = _dg['VALOR_BANCO'].abs().sum()
                        _cnt     = len(_dg)
                        _min_diff = _dg['CANDIDATO_DIFF_PCT'].min() if 'CANDIDATO_DIFF_PCT' in _dg.columns else None
                        _cand_doc  = _dg['CANDIDATO_DOC'].iloc[0] if 'CANDIDATO_DOC' in _dg.columns else ''
                        _cand_conc = _dg['CANDIDATO_CONCEPTO'].iloc[0] if 'CANDIDATO_CONCEPTO' in _dg.columns else ''
                        _cand_mon  = _dg['CANDIDATO_MONTO'].iloc[0] if 'CANDIDATO_MONTO' in _dg.columns else None
                        _razon     = _dg['RAZON_NO_MATCH'].iloc[0] if 'RAZON_NO_MATCH' in _dg.columns else ''
                        grupos_diag.append({
                            'Descripción Banco': str(_dk)[:45],
                            'Cant.': int(_cnt),
                            'Total ($)': round(_total, 0),
                            'Candidato Auxiliar': str(_cand_doc)[:15],
                            'Concepto Candidato': str(_cand_conc)[:40],
                            'Monto Candidato': round(float(_cand_mon), 0) if _cand_mon and str(_cand_mon) not in ('nan','') else None,
                            'Diff %': f"{_min_diff:.1f}%" if _min_diff is not None and str(_min_diff) != 'nan' else '—',
                            'Diagnóstico': _badge_grupo(_min_diff, _razon),
                        })
                    df_diag = pd.DataFrame(grupos_diag)
                    # Colorear filas según diagnóstico
                    def _color_diag(val):
                        if '🔴' in str(val): return 'color:#c62828;font-weight:700'
                        if '🟠' in str(val): return 'color:#e65100;font-weight:600'
                        if '🟡' in str(val): return 'color:#f57f17;font-weight:600'
                        return ''
                    st.dataframe(
                        df_diag.style.applymap(_color_diag, subset=['Diagnóstico']),
                        use_container_width=True, hide_index=True
                    )

                    st.markdown("""
                    <div class='callout-accion'>
                      <b>📌 ¿QUÉ HACER?</b>
                      <ul style='margin:4px 0 0 16px;'>
                        <li><b>🔴 Sin NC</b>: Crear Nota Contable en SIIGO con el valor exacto del cargo bancario.</li>
                        <li><b>🟠 Diferencia leve</b>: Verificar si la NC existe con monto diferente y ajustar.</li>
                        <li><b>🟡 Monto OK</b>: Revisar el número de documento o concepto en el auxiliar.</li>
                      </ul>
                    </div>""", unsafe_allow_html=True)

                    # ── DETALLE INDIVIDUAL ─────────────────────────────────
                    st.markdown("<div class='section-title'>📋 Detalle por Movimiento — Guía de Acción</div>", unsafe_allow_html=True)
                    for _, row in s_banco.iterrows():
                        st.markdown(_guia_banco_sin_aux(row), unsafe_allow_html=True)

            # ── SECCIÓN CRÍTICA: Asientos sin transacción bancaria ─────────
            n_sa = len(df_solo_aux)
            deb_sa = df_solo_aux['DEBITO'].sum()  if not df_solo_aux.empty else 0
            cre_sa = df_solo_aux['CREDITO'].sum() if not df_solo_aux.empty else 0
            with st.expander(f"📋 Asientos Auxiliar SIN Transacción Bancaria — {int(n_sa)} asientos", expanded=bool(n_sa > 0)):
                if df_solo_aux.empty:
                    st.markdown("<div class='callout-success'>✅ Todos los asientos tienen transacción bancaria.</div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class='callout-warning'>
                      <b>📋 {n_sa} asientos contables no tienen transacción bancaria correspondiente.</b><br>
                      Débitos sin banco: <b>{_cop_limpio(deb_sa)}</b>
                      &nbsp;·&nbsp;
                      Créditos sin banco: <b>{_cop_limpio(cre_sa)}</b>
                    </div>""", unsafe_allow_html=True)

                    st.markdown("""
                    <div class='callout-accion'>
                      <b>📌 ¿QUÉ HACER?</b> Para cada fila de abajo: busque si el movimiento bancario existe
                      en el extracto (puede ser de otro período). Si no existe, puede ser un asiento de ajuste
                      interno, un pago en efectivo, o requiere anulación.
                    </div>""", unsafe_allow_html=True)

                    st.markdown("<div class='section-title'>Guía de acción por asiento</div>", unsafe_allow_html=True)
                    for _, row in df_solo_aux.iterrows():
                        st.markdown(_guia_aux_sin_banco(row), unsafe_allow_html=True)

    with tab6:
        st.markdown("<div class='section-title'>⚖️ Conciliación Bancaria Formal</div>", unsafe_allow_html=True)

        # ── Variables de conciliación formal ──────────────────────────────────
        # Partidas conciliatorias: clasificadas desde df_solo_banco y df_solo_aux
        _banco_sin_aux = s_banco.copy() if not s_banco.empty else pd.DataFrame()
        _aux_sin_banco = df_solo_aux.copy() if not df_solo_aux.empty else pd.DataFrame()

        # Depósitos en tránsito: abonos del banco en los últimos 3 días del mes (sin auxiliar)
        _dep_transito = 0.0
        if not _banco_sin_aux.empty and 'VALOR_BANCO' in _banco_sin_aux.columns:
            _abonos_sb = _banco_sin_aux[_banco_sin_aux.get('TIPO_MOV', pd.Series(dtype=str)).str.upper().isin(['ABONO','CREDITO']) | (_banco_sin_aux['VALOR_BANCO'] > 0)]
            _dep_transito = float(_abonos_sb['VALOR_BANCO'].sum())

        # Cheques girados pendientes de cobro: CE/CG en auxiliar sin banco (débitos)
        _cheques_pend = 0.0
        if not _aux_sin_banco.empty and 'DEBITO' in _aux_sin_banco.columns:
            _ce_cg = _aux_sin_banco[
                _aux_sin_banco.get('DOCUMENTO', pd.Series(dtype=str)).str[:2].isin(['CE','CG','EG'])
            ]
            _cheques_pend = float(_ce_cg['DEBITO'].fillna(0).sum())

        # ND bancarias no contabilizadas (cargos banco sin auxiliar que NO son abonos)
        _nd_no_cont = 0.0
        if not _banco_sin_aux.empty and 'VALOR_BANCO' in _banco_sin_aux.columns:
            _cargos_sb = _banco_sin_aux[_banco_sin_aux['VALOR_BANCO'] < 0]
            _nd_no_cont = float(abs(_cargos_sb['VALOR_BANCO'].sum()))

        # NC bancarias no contabilizadas (créditos auxiliar sin banco)
        _nc_no_cont = 0.0
        if not _aux_sin_banco.empty and 'CREDITO' in _aux_sin_banco.columns:
            _nc_no_cont = float(_aux_sin_banco['CREDITO'].fillna(0).sum())

        # Saldos conciliados (ambos deben coincidir)
        _saldo_banco_conc = sac + _dep_transito - _cheques_pend
        _saldo_aux_conc   = sf_a + _nc_no_cont - _nd_no_cont
        _dif_conc         = _saldo_banco_conc - _saldo_aux_conc

        TL  = 'padding:5px 4px;font-size:.88rem;'
        TR  = 'padding:5px 4px;font-size:.88rem;text-align:right;'
        TLB = 'padding:7px 4px;font-size:.9rem;font-weight:700;'
        TRB = 'padding:7px 4px;font-size:.9rem;font-weight:800;text-align:right;'

        def _card(ok):
            if ok: return ('rgba(102,187,106,0.13)', '#66bb6a', 'rgba(102,187,106,0.35)')
            return ('rgba(239,83,80,0.13)', '#ef5350', 'rgba(239,83,80,0.35)')

        # ── I. Verificación aritmética de archivos ────────────────────────────
        with st.expander("📋 I. Verificación aritmética interna de archivos", expanded=False):
            calc_banco = sa + tab_s - tca_s
            dif_b      = calc_banco - sac
            calc_aux   = si_a + td_a - tc_a
            dif_a      = calc_aux - sf_a
            _c1, _c2   = st.columns(2)
            with _c1:
                bg_b, br_b, sp_b = _card(abs(dif_b) < 1)
                st.markdown(f"""
<div style='background:{bg_b};border-left:4px solid {br_b};border-radius:10px;padding:16px 20px;'>
  <div style='font-size:.9rem;font-weight:800;margin-bottom:10px;'>Extracto Bancario</div>
  <table style='width:100%;border-collapse:collapse;'>
    <tr><td style='{TL}'>Saldo anterior</td><td style='{TR}'>{cop(sa)}</td></tr>
    <tr><td style='{TL}'>(+) Total abonos</td><td style='{TR}'>{cop(tab_s)}</td></tr>
    <tr><td style='{TL}'>(-) Total cargos</td><td style='{TR}'>{cop(tca_s)}</td></tr>
    <tr><td colspan='2' style='border-top:1px solid {sp_b};padding:2px 0;'></td></tr>
    <tr><td style='{TLB}'>(=) Saldo calculado</td><td style='{TRB}'>{cop(calc_banco)}</td></tr>
    <tr><td style='{TL}'>Saldo declarado</td><td style='{TR}'>{cop(sac)}</td></tr>
    <tr><td style='{TLB}'>Diferencia</td>
        <td style='{TRB}color:{br_b};'>{"✅ CUADRA" if abs(dif_b)<1 else f"❌ {cop(dif_b)}"}</td></tr>
  </table>
</div>""", unsafe_allow_html=True)
            with _c2:
                bg_a, br_a, sp_a = _card(abs(dif_a) < 1)
                st.markdown(f"""
<div style='background:{bg_a};border-left:4px solid {br_a};border-radius:10px;padding:16px 20px;'>
  <div style='font-size:.9rem;font-weight:800;margin-bottom:10px;'>Auxiliar Contable</div>
  <table style='width:100%;border-collapse:collapse;'>
    <tr><td style='{TL}'>Saldo inicial</td><td style='{TR}'>{cop(si_a)}</td></tr>
    <tr><td style='{TL}'>(+) Total débitos</td><td style='{TR}'>{cop(td_a)}</td></tr>
    <tr><td style='{TL}'>(-) Total créditos</td><td style='{TR}'>{cop(tc_a)}</td></tr>
    <tr><td colspan='2' style='border-top:1px solid {sp_a};padding:2px 0;'></td></tr>
    <tr><td style='{TLB}'>(=) Saldo calculado</td><td style='{TRB}'>{cop(calc_aux)}</td></tr>
    <tr><td style='{TL}'>Saldo final declarado</td><td style='{TR}'>{cop(sf_a)}</td></tr>
    <tr><td style='{TLB}'>Diferencia</td>
        <td style='{TRB}color:{br_a};'>{"✅ CUADRA" if abs(dif_a)<1 else f"❌ {cop(dif_a)}"}</td></tr>
  </table>
</div>""", unsafe_allow_html=True)

        st.markdown('<br>', unsafe_allow_html=True)

        # ── II. CONCILIACIÓN FORMAL (2 columnas: banco → libros) ─────────────
        st.markdown("<div class='section-title'>📄 II. Conciliación Bancaria Formal</div>",
                    unsafe_allow_html=True)
        bg_conc, br_conc, _ = _card(abs(_dif_conc) < 1)
        _banco_ppal = WL.get("empresa_banco_principal","Banco") if _WL_DISPONIBLE else "Banco"
        _cta_ppal   = WL.get("empresa_cuenta_bancaria","Cuenta corriente") if _WL_DISPONIBLE else "Cuenta corriente"

        fc1, fc2 = st.columns(2)
        with fc1:
            st.markdown(f"""
<div style='background:rgba(66,165,245,0.07);border-left:4px solid #42a5f5;
            border-radius:10px;padding:18px 22px;height:100%;'>
  <div style='font-size:.92rem;font-weight:800;margin-bottom:14px;color:#42a5f5;'>
    🏦 BANCO — {_banco_ppal}</div>
  <table style='width:100%;border-collapse:collapse;'>
    <tr><td style='{TLB}'>Saldo según extracto bancario</td>
        <td style='{TRB}'>{cop(sac)}</td></tr>
    <tr><td colspan='2' style='border-top:0.5px solid rgba(66,165,245,0.3);padding:3px 0;'></td></tr>
    <tr><td style='{TL}color:#4CAF50;'>(+) Depósitos en tránsito</td>
        <td style='{TR}color:#4CAF50;'>{cop(_dep_transito)}</td></tr>
    <tr><td style='{TL}' colspan='2'>
      <small style='opacity:.6;font-size:.78rem;'>Abonos bancarios del período sin registro
      contable aún. Se registrarán en el mes siguiente.</small></td></tr>
    <tr><td colspan='2' style='padding:3px 0;'></td></tr>
    <tr><td style='{TL}color:#F44336;'>(-) Cheques girados pendientes de cobro</td>
        <td style='{TR}color:#F44336;'>{cop(_cheques_pend)}</td></tr>
    <tr><td style='{TL}' colspan='2'>
      <small style='opacity:.6;font-size:.78rem;'>CE/CG en auxiliar que el banco aún no ha
      debitado. Partida temporal hasta que el cheque sea cobrado.</small></td></tr>
    <tr><td colspan='2' style='border-top:1px solid rgba(66,165,245,0.4);padding:3px 0;'></td></tr>
    <tr><td style='{TLB}'>SALDO BANCARIO CONCILIADO</td>
        <td style='{TRB}color:#42a5f5;'>{cop(_saldo_banco_conc)}</td></tr>
  </table>
</div>""", unsafe_allow_html=True)

        with fc2:
            st.markdown(f"""
<div style='background:rgba(76,175,80,0.07);border-left:4px solid #4CAF50;
            border-radius:10px;padding:18px 22px;height:100%;'>
  <div style='font-size:.92rem;font-weight:800;margin-bottom:14px;color:#4CAF50;'>
    📋 LIBROS — {_cta_ppal}</div>
  <table style='width:100%;border-collapse:collapse;'>
    <tr><td style='{TLB}'>Saldo según auxiliar contable</td>
        <td style='{TRB}'>{cop(sf_a)}</td></tr>
    <tr><td colspan='2' style='border-top:0.5px solid rgba(76,175,80,0.3);padding:3px 0;'></td></tr>
    <tr><td style='{TL}color:#4CAF50;'>(+) Notas crédito banco no contabilizadas</td>
        <td style='{TR}color:#4CAF50;'>{cop(_nc_no_cont)}</td></tr>
    <tr><td style='{TL}' colspan='2'>
      <small style='opacity:.6;font-size:.78rem;'>Abonos que el banco ya registró pero
      el contador aún no ha causado en el auxiliar.</small></td></tr>
    <tr><td colspan='2' style='padding:3px 0;'></td></tr>
    <tr><td style='{TL}color:#F44336;'>(-) Notas débito banco no contabilizadas</td>
        <td style='{TR}color:#F44336;'>{cop(_nd_no_cont)}</td></tr>
    <tr><td style='{TL}' colspan='2'>
      <small style='opacity:.6;font-size:.78rem;'>Cargos que el banco ya aplicó (GMF,
      comisiones) pero el contador aún no ha registrado.</small></td></tr>
    <tr><td colspan='2' style='border-top:1px solid rgba(76,175,80,0.4);padding:3px 0;'></td></tr>
    <tr><td style='{TLB}'>SALDO LIBROS CONCILIADO</td>
        <td style='{TRB}color:#4CAF50;'>{cop(_saldo_aux_conc)}</td></tr>
  </table>
</div>""", unsafe_allow_html=True)

        st.markdown('<br>', unsafe_allow_html=True)

        # ── Verificación final de la ecuación ─────────────────────────────────
        bg_eq, br_eq, _ = _card(abs(_dif_conc) < 1)
        _eq_ok = abs(_dif_conc) < 1
        st.markdown(f"""
<div style='background:{bg_eq};border-left:4px solid {br_eq};border-radius:10px;padding:16px 22px;'>
  <table style='width:100%;border-collapse:collapse;'>
    <tr>
      <td style='{TLB}'>{"✅ CONCILIACIÓN CUADRA" if _eq_ok else "❌ CONCILIACIÓN NO CUADRA — REVISAR PARTIDAS"}</td>
      <td style='{TRB}color:{br_eq};font-size:1.1rem;'>
        Diferencia: {cop(_dif_conc)}</td>
    </tr>
    <tr><td colspan='2' style='border-top:0.5px solid {br_eq}44;padding:3px 0;'></td></tr>
    <tr><td style='{TL}'>Saldo bancario conciliado</td>
        <td style='{TR}'>{cop(_saldo_banco_conc)}</td></tr>
    <tr><td style='{TL}'>Saldo libros conciliado</td>
        <td style='{TR}'>{cop(_saldo_aux_conc)}</td></tr>
  </table>
  <div style='margin-top:10px;font-size:.82rem;opacity:.8;'>
    {"Ambas columnas coinciden. La conciliación está completa." if _eq_ok
     else f"Queda una diferencia de {cop(abs(_dif_conc)).strip()} sin explicar. Revise si existen partidas no clasificadas en las pestañas Diferencias y Partidas Conciliatorias."}
  </div>
</div>""", unsafe_allow_html=True)

        # ── III. Conclusión auditora dinámica ─────────────────────────────────
        if not df_aux.empty:
            ico2, lbl2, _ = _semaforo_conciliacion(pct_conc)
            b2 = 'badge-verde' if pct_conc >= 90 else 'badge-naranja'
            _per_str_t6 = ""
            try:
                _pd6 = _periodo_detectado if _periodo_detectado else (_extraer_periodo(_nombre_b) if banco_files else None)
                if _pd6:
                    _per_str_t6 = f"{_MESES_ES[_pd6[1]-1]} de {_pd6[0]}"
            except Exception:
                pass
            _nit_t6    = WL.get("empresa_nit",  "") if _WL_DISPONIBLE else ""
            _ciudad_t6 = WL.get("empresa_ciudad","") if _WL_DISPONIBLE else ""
            _fecha_hoy = datetime.now().strftime("%d de %B de %Y").replace(
                "January","enero").replace("February","febrero").replace("March","marzo"
                ).replace("April","abril").replace("May","mayo").replace("June","junio"
                ).replace("July","julio").replace("August","agosto").replace("September","septiembre"
                ).replace("October","octubre").replace("November","noviembre").replace("December","diciembre")
            st.markdown('<br>', unsafe_allow_html=True)
            st.markdown(f"""
<div class='callout-info'>
  <b>III. Conclusión Auditora</b><br><br>
  La presente conciliación bancaria fue elaborada con base en el extracto bancario y el
  auxiliar contable de la cuenta <b>{_cta_ppal}</b>, correspondiente al período
  <b>{_per_str_t6 or "indicado en los archivos"}</b>.<br><br>
  Empresa: <b>{_empresa}</b>{(" · NIT: " + _nit_t6) if _nit_t6 else ""}
  {(" · " + _ciudad_t6) if _ciudad_t6 else ""}.<br>
  Tasa de conciliación: <span class='{b2}'>{ico2} {pct_conc:.1f}% — {lbl2}</span><br><br>
  <b>{"✅ Los saldos cuadran. Los registros contables y bancarios están en orden." if _eq_ok
     else f"⚠️ Existen {n_sbco + n_saux} partidas sin conciliar por un valor total de {cop(abs(_dep_transito)+abs(_nd_no_cont)+abs(_nc_no_cont)).strip()} COP. Se deben revisar antes del cierre contable."}</b><br><br>
  <small style='opacity:.7;'>Elaborada el {_fecha_hoy} por el sistema CREDIEXPRESS.</small>
</div>""", unsafe_allow_html=True)

        # ── IV. Detalle de partidas sin conciliar ──────────────────────────────
        if (n_sbco + n_saux) > 0 and not df_aux.empty:
            st.markdown('<br>', unsafe_allow_html=True)
            _tot_sin_conc = abs(_dep_transito) + abs(_nd_no_cont) + abs(_nc_no_cont)
            with st.expander(
                f"\U0001f50d IV. Detalle de partidas sin conciliar "
                f"\u2014 {n_sbco + n_saux} items \u00b7 {cop(_tot_sin_conc).strip()} COP",
                expanded=True,
            ):
                # ---- A: Solo en banco ----------------------------------------
                if not _banco_sin_aux.empty:
                    st.markdown("##### \U0001f3e6 A. Transacciones en Banco sin registrar en Auxiliar")
                    st.caption(
                        "Aparecen en el extracto bancario pero NO tienen asiento contable en el auxiliar. "
                        "Origen probable: dep\u00f3sitos en tr\u00e1nsito, ND bancarias, GMF o comisiones no causadas."
                    )
                    _cols_sb = [c for c in ["FECHA_RAW","DESCRIPCION","VALOR_BANCO","SALDO","TIPO_MOV"]
                                if c in _banco_sin_aux.columns]
                    _df_sb = _banco_sin_aux[_cols_sb].copy().rename(columns={
                        "FECHA_RAW": "Fecha", "DESCRIPCION": "Descripci\u00f3n",
                        "VALOR_BANCO": "Valor ($)", "SALDO": "Saldo", "TIPO_MOV": "Tipo",
                    })

                    def _causa_banco(row):
                        val  = float(row.get("Valor ($)", 0) or 0)
                        desc = str(row.get("Descripci\u00f3n", "")).upper()
                        if val > 0:
                            return "\U0001f535 Dep\u00f3sito en tr\u00e1nsito"
                        if any(k in desc for k in ("GMF","4X1000","IMPUESTO")):
                            return "\U0001f4cb GMF \u2014 Cta. 5305"
                        if any(k in desc for k in ("COMISION","CUOTA","MANEJO","ADMON")):
                            return "\U0001f4cb Comisi\u00f3n bancaria"
                        if any(k in desc for k in ("CREDITO","DESEMBOLSO","ABONO")):
                            return "\U0001f4cb NC bancaria no causada"
                        return "\u2753 ND no contabilizada \u2014 revisar"

                    _df_sb["Causa probable"] = _df_sb.apply(_causa_banco, axis=1)
                    # Pre-formatear columnas monetarias como texto (formato cop)
                    if "Valor ($)" in _df_sb.columns:
                        _df_sb["Valor ($)"] = _df_sb["Valor ($)"].apply(
                            lambda v: cop(v) if pd.notna(v) else ""
                        )
                    if "Saldo" in _df_sb.columns:
                        _df_sb["Saldo"] = _df_sb["Saldo"].apply(
                            lambda v: cop(v) if pd.notna(v) else ""
                        )
                    st.dataframe(_df_sb, use_container_width=True, hide_index=True)

                    if "Fecha" in _df_sb.columns:
                        _grp_b = (_df_sb.groupby("Fecha")["Valor ($)"]
                                  .agg(N="count", Total="sum").reset_index()
                                  .rename(columns={"N": "N\u00b0 items", "Total": "Total ($)"})
                                  .sort_values("Fecha"))
                        st.markdown("**\U0001f4c5 Resumen por fecha \u2014 Solo banco:**")
                        _grp_b["Total ($)"] = _grp_b["Total ($)"].apply(
                            lambda v: cop(v) if pd.notna(v) else ""
                        )
                        st.dataframe(_grp_b, use_container_width=True, hide_index=True)

                st.markdown("---")

                # ---- B: Solo en auxiliar -------------------------------------
                if not _aux_sin_banco.empty:
                    st.markdown("##### \U0001f4cb B. Asientos en Auxiliar sin movimiento en Banco")
                    st.caption(
                        "Existen en el auxiliar pero NO tienen transacci\u00f3n bancaria correspondiente. "
                        "Origen probable: cheques pendientes de cobro, pagos ACH no liquidados, o errores de registro."
                    )
                    _cols_ax = [c for c in ["FECHA","DOCUMENTO","CUENTA","DEBITO","CREDITO"]
                                if c in _aux_sin_banco.columns]
                    _df_ax = _aux_sin_banco[_cols_ax].copy().rename(columns={
                        "FECHA": "Fecha", "DOCUMENTO": "Comprobante",
                        "CUENTA": "Cuenta PUC", "DEBITO": "D\u00e9bito ($)", "CREDITO": "Cr\u00e9dito ($)",
                    })

                    def _causa_aux(row):
                        comp = str(row.get("Comprobante", "")).upper()[:2]
                        deb  = float(row.get("D\u00e9bito ($)", 0) or 0)
                        if comp in ("CE","CG","EG"):
                            return "\U0001f534 Cheque/egreso pendiente de cobro"
                        if comp == "NC":
                            return "\U0001f7e2 NC no reflejada a\u00fan en banco"
                        if deb > 0:
                            return "\u2753 D\u00e9bito sin contrapartida bancaria"
                        return "\u2753 Cr\u00e9dito sin contrapartida bancaria"

                    _df_ax["Causa probable"] = _df_ax.apply(_causa_aux, axis=1)
                    # Calcular resumen por fecha ANTES de formatear (necesita numericos)
                    _grp_a = None
                    if "Fecha" in _df_ax.columns:
                        _val_c = "D\u00e9bito ($)" if "D\u00e9bito ($)" in _df_ax.columns else "Cr\u00e9dito ($)"
                        try:
                            _grp_a = (_df_ax.groupby("Fecha")[_val_c]
                                      .agg(N="count", Total="sum").reset_index()
                                      .rename(columns={"N": "N\u00b0 asientos", "Total": "Total ($)"})
                                      .sort_values("Fecha"))
                            _grp_a["Total ($)"] = _grp_a["Total ($)"].apply(
                                lambda v: cop(v) if pd.notna(v) else ""
                            )
                            _grp_a["Fecha"] = pd.to_datetime(
                                _grp_a["Fecha"], errors="coerce"
                            ).dt.strftime("%d/%m/%Y").fillna("")
                        except Exception:
                            _grp_a = None
                    # Formatear fecha y monetarios para display
                    if "Fecha" in _df_ax.columns:
                        _df_ax["Fecha"] = pd.to_datetime(
                            _df_ax["Fecha"], errors="coerce"
                        ).dt.strftime("%d/%m/%Y").fillna("")
                    if "D\u00e9bito ($)" in _df_ax.columns:
                        _df_ax["D\u00e9bito ($)"] = _df_ax["D\u00e9bito ($)"].fillna(0).apply(
                            lambda v: cop(v) if v else ""
                        )
                    if "Cr\u00e9dito ($)" in _df_ax.columns:
                        _df_ax["Cr\u00e9dito ($)"] = _df_ax["Cr\u00e9dito ($)"].fillna(0).apply(
                            lambda v: cop(v) if v else ""
                        )
                    st.dataframe(_df_ax, use_container_width=True, hide_index=True)

                    if _grp_a is not None:
                        st.markdown("**\U0001f4c5 Resumen por fecha \u2014 Solo auxiliar:**")
                        st.dataframe(_grp_a, use_container_width=True, hide_index=True)

                st.markdown("---")
                st.markdown("""
<div style='font-size:.82rem;opacity:.75;padding:8px 0;'>
  <b>Acci\u00f3n requerida:</b> Antes del cierre contable revise cada partida listada.<br>
  \u2022 <b>Grupo A</b>: causer las en el auxiliar o confirmar que son dep\u00f3sitos en tr\u00e1nsito leg\u00edtimos.<br>
  \u2022 <b>Grupo B</b>: cotejar con comprobantes f\u00edsicos para verificar si el banco ya los proces\u00f3.
</div>""", unsafe_allow_html=True)

        # ── V. Firmantes (dinámicos desde White Label) ───────────────────────
        _rep_legal   = WL.get("empresa_rep_legal","")    if _WL_DISPONIBLE else ""
        _rep_legal_cc= WL.get("empresa_rep_legal_cc","") if _WL_DISPONIBLE else ""
        _contador    = WL.get("empresa_contador","")      if _WL_DISPONIBLE else ""
        _tp_contador = WL.get("empresa_tp_contador","")   if _WL_DISPONIBLE else ""

        if _rep_legal or _contador:
            st.markdown('<br>', unsafe_allow_html=True)
            _firma_rl = f"""
  <div style='flex:1;text-align:center;
              border-right:1px solid rgba(66,165,245,0.2);padding-right:24px;'>
    <div style='font-weight:800;font-size:.95rem;color:#42a5f5;'>{_rep_legal.upper() or "REPRESENTANTE LEGAL"}</div>
    <div style='font-size:.82rem;opacity:.7;margin-top:3px;'>REPRESENTANTE LEGAL</div>
    {f"<div style='font-size:.78rem;opacity:.5;margin-top:2px;'>C.C. {_rep_legal_cc}</div>" if _rep_legal_cc else ""}
  </div>"""
            _firma_ct = f"""
  <div style='flex:1;text-align:center;padding-left:24px;'>
    <div style='font-weight:800;font-size:.95rem;color:#42a5f5;'>{_contador.upper() or "CONTADOR PÚBLICO"}</div>
    <div style='font-size:.82rem;opacity:.7;margin-top:3px;'>CONTADOR PÚBLICO</div>
    {f"<div style='font-size:.78rem;opacity:.5;margin-top:2px;'>T.P. {_tp_contador}</div>" if _tp_contador else ""}
  </div>"""
            st.markdown(f"""
<div style='display:flex;gap:32px;padding:22px 28px;
            background:rgba(66,165,245,0.06);
            border:1px solid rgba(66,165,245,0.2);
            border-radius:12px;'>
  {_firma_rl}{_firma_ct}
</div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
<div class='callout-warning' style='font-size:.85rem;'>
  <b>⚙️ Configura los firmantes</b> en <b>Tab 14 → Admin → White Label → Firmantes</b> para que
  aparezcan aquí: Representante Legal y Contador Público con sus datos completos.
</div>""", unsafe_allow_html=True)

    with tab7:
        st.markdown("<div class='section-title'>📈 Visualizaciones Interactivas</div>", unsafe_allow_html=True)

        if _PLOTLY_DISPONIBLE:
            _plotly_layout = dict(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Segoe UI, system-ui, sans-serif", size=12),
                margin=dict(l=10, r=10, t=40, b=10),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )

            # ── Fila 1: Saldo + Pastel ───────────────────────────────────────
            col_g1, col_g2 = st.columns([3, 2])

            with col_g1:
                st.markdown("**📉 Evolución del Saldo Bancario**")
                df_s = df_banco[df_banco["SALDO"].notna()].copy() if not df_banco.empty else pd.DataFrame()
                fig1 = go.Figure()
                if not df_s.empty:
                    _x1 = list(range(len(df_s)))
                    _y1 = (df_s["SALDO"] / 1e6).tolist()
                    fig1.add_trace(go.Scatter(
                        x=_x1, y=_y1,
                        mode="lines",
                        line=dict(color="#42a5f5", width=2.5),
                        fill="tozeroy",
                        fillcolor="rgba(66,165,245,0.10)",
                        name="Saldo (M COP)",
                        hovertemplate="Mov. %{x}<br>Saldo: $%{y:.2f}M<extra></extra>",
                    ))
                    _min_s, _max_s = min(_y1), max(_y1)
                    fig1.add_hline(y=(_min_s+_max_s)/2, line_dash="dot",
                                   line_color="rgba(255,183,77,0.5)",
                                   annotation_text="Promedio", annotation_position="right")
                fig1.update_layout(**_plotly_layout, height=280,
                    xaxis=dict(title="Movimiento #", gridcolor="rgba(255,255,255,0.06)"),
                    yaxis=dict(title="Millones COP", gridcolor="rgba(255,255,255,0.06)",
                               tickprefix="$", ticksuffix="M"),
                )
                st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})

            with col_g2:
                st.markdown("**🥧 Estado de Conciliación**")
                fig2 = go.Figure()
                if n_tot > 0:
                    _cont = df_comp["ESTADO"].value_counts()
                    _cmap = {"COINCIDE EXACTO": "#4CAF50", "COINCIDE APROX.": "#FFC107",
                             "SOLO EN BANCO": "#F44336", "SOLO EN AUXILIAR": "#2196F3"}
                    _cols2 = [next((v for k, v in _cmap.items() if k in e), "#9E9E9E")
                              for e in _cont.index]
                    _lbl2 = [e.replace("✅ ", "").replace("🔶 ", "").replace("❌ ", "")
                              .replace(" — CONFIRMAR", "") for e in _cont.index]
                    fig2.add_trace(go.Pie(
                        labels=_lbl2, values=_cont.values.tolist(),
                        marker_colors=_cols2,
                        hole=0.45,
                        textinfo="percent+label",
                        hovertemplate="%{label}<br>%{value} mov. (%{percent})<extra></extra>",
                        textfont_size=11,
                    ))
                    fig2.add_annotation(text=f"<b>{pct_conc:.0f}%</b><br><span style='font-size:10px'>conciliado</span>",
                                        x=0.5, y=0.5, showarrow=False, font=dict(size=14, color="#42a5f5"))
                fig2.update_layout(**_plotly_layout, height=280, showlegend=False)
                st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

            # ── Fila 2: Movimientos por día + Barras banco vs auxiliar ───────
            col_g3, col_g4 = st.columns(2)

            with col_g3:
                st.markdown("**📅 Movimientos por Día del Mes**")
                fig3 = go.Figure()
                if not df_banco.empty:
                    _df_d = df_banco.copy()
                    _df_d["DIA"] = _df_d["FECHA_RAW"].apply(
                        lambda x: int(str(x).split("/")[0]) if "/" in str(x) else 0)
                    _por_dia = _df_d.groupby(["DIA", "TIPO"])["VALOR"].sum().unstack(fill_value=0)
                    if "ABONO" in _por_dia.columns:
                        fig3.add_trace(go.Bar(
                            x=_por_dia.index.tolist(),
                            y=(_por_dia["ABONO"] / 1e6).tolist(),
                            name="Abonos", marker_color="#4CAF50",
                            hovertemplate="Día %{x}<br>Abonos: $%{y:.2f}M<extra></extra>",
                        ))
                    if "CARGO" in _por_dia.columns:
                        fig3.add_trace(go.Bar(
                            x=_por_dia.index.tolist(),
                            y=(_por_dia["CARGO"].abs() / 1e6).tolist(),
                            name="Cargos", marker_color="#F44336",
                            hovertemplate="Día %{x}<br>Cargos: $%{y:.2f}M<extra></extra>",
                        ))
                fig3.update_layout(**_plotly_layout, height=280, barmode="group",
                    xaxis=dict(title="Día", gridcolor="rgba(255,255,255,0.06)"),
                    yaxis=dict(title="Millones COP", gridcolor="rgba(255,255,255,0.06)",
                               tickprefix="$", ticksuffix="M"),
                )
                st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

            with col_g4:
                st.markdown("**⚖️ Banco vs Auxiliar — Totales**")
                fig4 = go.Figure()
                _cats4 = ["Entradas Banco", "Débitos Aux.", "Salidas Banco", "Créditos Aux."]
                _vals4 = [tab_s/1e6, td_a/1e6, tca_s/1e6, tc_a/1e6]
                _cols4 = ["#2196F3", "#4CAF50", "#F44336", "#FF9800"]
                fig4.add_trace(go.Bar(
                    x=_cats4, y=_vals4,
                    marker_color=_cols4,
                    text=[f"${v:.1f}M" for v in _vals4],
                    textposition="outside",
                    hovertemplate="%{x}<br>$%{y:.2f}M<extra></extra>",
                ))
                fig4.update_layout(**_plotly_layout, height=280, showlegend=False,
                    xaxis=dict(gridcolor="rgba(255,255,255,0.06)"),
                    yaxis=dict(title="Millones COP", gridcolor="rgba(255,255,255,0.06)",
                               tickprefix="$", ticksuffix="M"),
                )
                st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar": False})

            # ── Fila 3: Tipo comprobante + Valor por estado ──────────────────
            col_g5, col_g6 = st.columns(2)

            with col_g5:
                st.markdown("**📋 Asientos por Tipo de Comprobante**")
                fig5 = go.Figure()
                if not df_aux.empty and "DOCUMENTO" in df_aux.columns:
                    _tc = df_aux["DOCUMENTO"].str[:2].value_counts()
                    _cols5 = ["#4CAF50","#2196F3","#FF9800","#9C27B0","#F44336","#00BCD4"]
                    fig5.add_trace(go.Bar(
                        x=_tc.index.tolist(), y=_tc.values.tolist(),
                        marker_color=_cols5[:len(_tc)],
                        text=_tc.values.tolist(),
                        textposition="outside",
                        hovertemplate="Tipo: %{x}<br>%{y} asientos<extra></extra>",
                    ))
                else:
                    fig5.add_annotation(text="Sin datos de auxiliar",
                                        x=0.5, y=0.5, showarrow=False,
                                        font=dict(size=14, color="#90CAF9"))
                fig5.update_layout(**_plotly_layout, height=280, showlegend=False,
                    xaxis=dict(gridcolor="rgba(255,255,255,0.06)"),
                    yaxis=dict(title="N° asientos", gridcolor="rgba(255,255,255,0.06)"),
                )
                st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar": False})

            with col_g6:
                st.markdown("**💰 Valor Monetario por Estado**")
                fig6 = go.Figure()
                if n_tot > 0:
                    _ve = exactas["VALOR_BANCO"].abs().sum()/1e6 if not exactas.empty else 0
                    _va = aprox["VALOR_BANCO"].abs().sum()/1e6   if not aprox.empty   else 0
                    _vs = s_banco["VALOR_BANCO"].abs().sum()/1e6 if not s_banco.empty else 0
                    _vx = (df_solo_aux["DEBITO"].fillna(0).sum() +
                           df_solo_aux["CREDITO"].fillna(0).sum())/1e6 if not df_solo_aux.empty else 0
                    _lbl6 = ["Exacto", "Aprox.", "Solo banco", "Solo auxiliar"]
                    _val6 = [_ve, _va, _vs, _vx]
                    fig6.add_trace(go.Bar(
                        x=_lbl6, y=_val6,
                        marker_color=["#4CAF50", "#FFC107", "#F44336", "#2196F3"],
                        text=[f"${v:.1f}M" for v in _val6],
                        textposition="outside",
                        hovertemplate="%{x}<br>$%{y:.2f}M<extra></extra>",
                    ))
                fig6.update_layout(**_plotly_layout, height=280, showlegend=False,
                    xaxis=dict(gridcolor="rgba(255,255,255,0.06)"),
                    yaxis=dict(title="Millones COP", gridcolor="rgba(255,255,255,0.06)",
                               tickprefix="$", ticksuffix="M"),
                )
                st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar": False})

        else:
            # Fallback matplotlib si Plotly no está instalado
            import matplotlib.pyplot as plt
            import matplotlib.ticker as mticker
            plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 110})
            fig, axes = plt.subplots(2, 3, figsize=(22, 12))
            fig.suptitle("CREDIEXPRESS POPAYAN SAS — Conciliacion Bancaria",
                         fontsize=14, fontweight="bold", y=1.01)
            ax1, ax2, ax3 = axes[0]
            ax4, ax5, ax6 = axes[1]
            df_s = df_banco[df_banco["SALDO"].notna()].copy() if not df_banco.empty else pd.DataFrame()
            if not df_s.empty:
                ax1.plot(range(len(df_s)), df_s["SALDO"]/1e6, color="#1565C0", lw=1.2)
                ax1.fill_between(range(len(df_s)), df_s["SALDO"]/1e6, alpha=0.12, color="#1565C0")
            ax1.set_title("Evolucion del Saldo Bancario", fontweight="bold")
            ax1.set_ylabel("Millones COP")
            ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}M"))
            ax1.grid(True, alpha=0.3)
            if n_tot > 0:
                cont = df_comp["ESTADO"].value_counts()
                _cmap = {"COINCIDE EXACTO":"#4CAF50","COINCIDE APROX.":"#FFC107","SOLO EN BANCO":"#F44336"}
                cs = [next((v for k, v in _cmap.items() if k in e), "#9E9E9E") for e in cont.index]
                ax2.pie(cont.values, labels=cont.index, colors=cs, autopct="%1.0f%%", startangle=90)
            ax2.set_title("Estado Conciliacion", fontweight="bold")
            _cats = ["Entradas\nBanco", "Debitos\nAux.", "Salidas\nBanco", "Creditos\nAux."]
            _vals = [tab_s/1e6, td_a/1e6, tca_s/1e6, tc_a/1e6]
            ax3.bar(_cats, _vals, color=["#2196F3","#4CAF50","#F44336","#FF9800"], alpha=0.85)
            ax3.set_title("Totales: Banco vs Auxiliar", fontweight="bold")
            ax3.set_ylabel("Millones COP")
            ax3.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}M"))
            ax3.grid(True, axis="y", alpha=0.3)
            plt.tight_layout()
            st.pyplot(fig)
            st.info("💡 Instala plotly para gráficas interactivas: `pip install plotly`")

        # Interpretacion
        st.markdown("<div class='section-title'>📖 Interpretación de los Gráficos</div>", unsafe_allow_html=True)
        ico_v, lbl_v, _ = _semaforo_conciliacion(pct_conc)
        st.markdown(f"""
        <div class='callout-info'>
          <b>G1 — Evolución del Saldo:</b> Muestra cómo varió el saldo bancario durante el período.
          Una línea estable indica flujo predecible; caídas o picos bruscos requieren revisión.<br><br>
          <b>G2 — Estado de Conciliación:</b> {ico_v} <b>{pct_conc:.1f}%</b> de los movimientos están conciliados ({lbl_v}).
          El <b>{100-pct_conc:.1f}%</b> restante está pendiente de revisar.<br><br>
          <b>G3 — Movimientos por Día:</b> Identifica días de mayor actividad bancaria.
          Picos de débitos o créditos en fechas específicas pueden indicar pagos masivos o recaudos.<br><br>
          <b>G4 — Banco vs Auxiliar:</b> Compara los totales entre el extracto y el auxiliar.
          Barras parejas indican registros completos; barras desiguales apuntan a diferencias.<br><br>
          <b>G5 — Tipo de Comprobante:</b> Distribución de asientos por prefijo de documento
          (NC = notas crédito, CE = comprobantes egreso, CG = comprobantes de caja, etc.).<br><br>
          <b>G6 — Valor por Estado:</b> Valor monetario en cada categoría. El mayor valor debe estar
          en "Exacto"; valores altos en "Solo banco" o "Solo auxiliar" requieren atención inmediata.
        </div>""", unsafe_allow_html=True)

    with tab8:
        st.markdown("<div class='section-title'>Exportar a Excel</div>", unsafe_allow_html=True)
        st.markdown("""
        <div class='callout-info'>
          El archivo Excel contiene <b>8 hojas</b> con toda la informacion del analisis:
          Comparacion completa, Coincidencias exactas, Aproximadas, Solo en banco, Solo en auxiliar,
          Datos completos del extracto, Datos completos del auxiliar y Resumen ejecutivo.
        </div>""", unsafe_allow_html=True)

        _per_det = st.session_state.get('_periodo_detectado') or _periodo_detectado if '_periodo_detectado' in dir() else None
        _per_lbl = (f"{_MESES_ES[_per_det[1]-1]}_{_per_det[0]}" if _per_det else "")
        _nombre_base = f"Conciliacion_{_per_lbl}" if _per_lbl else "Conciliacion"
        # Truncar a 60 chars para no superar límites del OS en nombres de archivo
        nombre_salida = f"{_nombre_base[:60]}.xlsx"

        output, nombre_salida = generar_excel(
            df_comp=df_comp, df_banco=df_banco, df_aux=df_aux, df_solo_aux=df_solo_aux,
            banco_nombre=_nombre_b, aux_nombre=_nombre_a,
            sa=sa, sac=sac, tab_s=tab_s, tca_s=tca_s,
            si_a=si_a, sf_a=sf_a, td_a=td_a, tc_a=tc_a,
            n_tot=n_tot, n_exac=n_exac, n_apr=n_apr, n_agr=n_agr,
            n_rec=n_rec, n_sbco=n_sbco, n_saux=n_saux, pct_conc=pct_conc,
            nombre_salida=nombre_salida,
        )
        output.seek(0)
        # ── Auto-guardar Excel localmente (solo offline) ─────────────────────
        if OFFLINE_MODE:
            _excel_local = _auto_guardar_excel(output.getvalue(), nombre_salida)
            if _excel_local:
                st.markdown(f"""
<div class='callout-success' style='margin-bottom:10px;font-size:.85rem;'>
  💾 Excel guardado automaticamente en:<br>
  <code style='font-size:.78rem;word-break:break-all;'>{_excel_local}</code>
</div>""", unsafe_allow_html=True)


        # ── PDF firmado digitalmente ──────────────────────────────────────
        if _PDF_FIRMADO_DISPONIBLE:
            if st.button("📄 Generar PDF firmado (SHA-256)", use_container_width=True):
                with st.spinner("Generando PDF con firma digital..."):
                    try:
                        _pdf_b, _pdf_n, _pdf_h = generar_pdf_conciliacion(
                            df_comp=df_comp, df_banco=df_banco, df_aux=df_aux,
                            banco_nombre=_nombre_b, aux_nombre=_nombre_a,
                            periodo=nombre_salida[:10],
                            pct_conc=pct_conc,
                            saldo_banco=float(sac or 0), saldo_aux=float(sf_a or 0),
                            diferencia_neta=float((sac or 0)-(sf_a or 0)),
                            usuario=_usuario_actual.get("username","admin"),
                            n_exactas=n_exac, n_aprox=n_apr,
                            n_agrupadas=n_agr, n_rechazos=n_rec,
                        )
                        if _pdf_b:
                            st.success(f"✅ PDF generado · SHA-256: `{_pdf_h[:32]}…`")
                            st.download_button("⬇️ Descargar PDF firmado", data=_pdf_b,
                                               file_name=_pdf_n, mime="application/pdf",
                                               use_container_width=True)
                            if _NOTIF_DISPONIBLE:
                                encolar_notificacion(
                                    "CONCILIACION_LISTA", "LOG",
                                    _usuario_actual.get("username","admin"),
                                    f"Conciliación PDF lista — {_pdf_n}",
                                    f"SHA-256: {_pdf_h}"
                                )
                        else:
                            st.warning("PDF no generado. Instale reportlab o fpdf2: "
                                       "`pip install reportlab`")
                    except Exception as _epdf:
                        st.error(f"Error PDF: {_epdf}")

        col_dl, col_info = st.columns([1, 2])
        with col_dl:
            st.download_button(
                label="Descargar Excel Premium",
                data=output,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_info:
            st.markdown(f"""
            <div class='callout-success'>
              Archivo listo: <b>{nombre_salida}</b><br>
              Movimientos banco: <b>{len(df_banco)}</b> &nbsp;|&nbsp;
              Asientos auxiliar: <b>{len(df_aux)}</b> &nbsp;|&nbsp;
              Conciliacion: <b>{pct_conc:.1f}%</b>
            </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div class='callout-success' style='margin-top:20px;'>
      <b>Analisis completado exitosamente.</b>
      Revise las pestanas para el detalle completo. Descargue el Excel para el archivo oficial.
    </div>""", unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 9 — PARTIDAS CONCILIATORIAS
    # ══════════════════════════════════════════════════════════════════════════
    with tab9:
        st.markdown("<div class='section-title'>🗂️ Partidas Conciliatorias</div>",
                    unsafe_allow_html=True)
        if not _PARTIDAS_DISPONIBLE:
            st.warning("Módulo de partidas no disponible.")
        else:
            periodo_actual = ""
            if banco_file:
                _per = _periodo_detectado or _extraer_periodo(_nombre_b)
                if _per:
                    periodo_actual = f"{_per[0]}-{_per[1]:02d}"

            # Auto-detección si hay datos procesados
            if st.session_state.get('run') and 'df_solo_banco' in dir():
                try:
                    n_auto = detectar_partidas_automaticas(
                        df_solo_banco, df_solo_aux,
                        periodo=periodo_actual or "SIN-PERIODO",
                        banco=_nombre_b if banco_files else "",
                        usuario=_usuario_actual.get("username", "admin")
                    )
                    if n_auto:
                        st.success(f"✅ {n_auto} partidas auto-detectadas desde movimientos sin conciliar")
                except Exception as _ep:
                    pass

            # Resumen
            _res_part = resumen_partidas()
            c1, c2, c3, c4 = st.columns(4)
            for _col, _est, _label, _color in [
                (c1, "PENDIENTE",  "⏳ Pendientes",  "#F9A825"),
                (c2, "EN_PROCESO", "🔄 En proceso",  "#1565C0"),
                (c3, "CONCILIADA", "✅ Conciliadas", "#2E7D32"),
                (c4, "ANULADA",    "❌ Anuladas",    "#C62828"),
            ]:
                cnt, tot = _res_part.get(_est, (0, 0))
                with _col:
                    st.markdown(f"""<div style='background:{_bg_card};border-left:4px solid {_color};
                    border-radius:8px;padding:14px 16px;'>
                    <div style='font-size:1.5rem;font-weight:800;color:{_color};'>{cnt}</div>
                    <div style='font-size:.8rem;opacity:.8;'>{_label}</div>
                    <div style='font-size:.75rem;opacity:.65;'>${tot:,.0f} COP</div>
                    </div>""", unsafe_allow_html=True)

            st.markdown("---")

            # Registrar nueva partida
            if _tiene_permiso("partidas"):
                with st.expander("➕ Registrar nueva partida conciliatoria", expanded=False):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        _tipo_p  = st.selectbox("Tipo", list(TIPOS_PARTIDA.keys()),
                                                 format_func=lambda x: TIPOS_PARTIDA[x])
                        _desc_p  = st.text_input("Descripción")
                        _valor_p = st.number_input("Valor (COP)", min_value=0.0, step=1000.0)
                    with col_b:
                        _per_p   = st.text_input("Período origen (YYYY-MM)", value=periodo_actual)
                        _banco_p = st.text_input("Banco", value=_nombre_b if banco_files else "")
                        _obs_p   = st.text_area("Observaciones", height=68)
                    if st.button("💾 Registrar partida", type="primary"):
                        ok, msg = registrar_partida(
                            _per_p, _tipo_p, _desc_p, _valor_p,
                            banco=_banco_p, observaciones=_obs_p,
                            usuario=_usuario_actual.get("username", "admin")
                        )
                        if ok:
                            st.success(f"✅ Partida registrada (UUID: {msg})")
                        else:
                            st.error(f"❌ {msg}")

            # Lista partidas
            _estado_f = st.selectbox("Filtrar por estado",
                                      ["TODOS", "PENDIENTE", "EN_PROCESO", "CONCILIADA", "ANULADA"])
            _partidas = listar_partidas(estado=None if _estado_f == "TODOS" else _estado_f)

            if _partidas:
                _df_part = pd.DataFrame(_partidas)[
                    ["periodo_origen","tipo_label","descripcion","valor","banco","estado_label","fecha"]
                ].rename(columns={
                    "periodo_origen": "Período", "tipo_label": "Tipo",
                    "descripcion": "Descripción", "valor": "Valor COP",
                    "banco": "Banco", "estado_label": "Estado", "fecha": "Registrada",
                })
                _df_part["Valor COP"] = _df_part["Valor COP"].apply(lambda x: f"${x:,.0f}")
                st.dataframe(_df_part, use_container_width=True, hide_index=True)

                # Conciliar seleccionada
                if _tiene_permiso("partidas"):
                    _pend = [p for p in _partidas if p["estado"] == "PENDIENTE"]
                    if _pend:
                        _opciones = {p["uuid"]: f"{p['descripcion'][:40]} | ${p['valor']:,.0f}" for p in _pend}
                        _sel_uuid = st.selectbox("Marcar como conciliada:",
                                                  list(_opciones.keys()),
                                                  format_func=lambda x: _opciones[x])
                        _per_cierre = st.text_input("Período de cierre (YYYY-MM)")
                        if st.button("✅ Conciliar partida seleccionada"):
                            ok, msg = conciliar_partida(_sel_uuid, _per_cierre,
                                                         usuario=_usuario_actual.get("username","admin"))
                            if ok:
                                st.success("✅ Partida conciliada")
                            else:
                                st.error(f"❌ {msg}")
            else:
                st.info("No hay partidas registradas para el filtro seleccionado.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 10 — COMISIONES BANCARIAS
    # ══════════════════════════════════════════════════════════════════════════
    with tab10:
        st.markdown("<div class='section-title'>💸 Comisiones Bancarias Detectadas</div>",
                    unsafe_allow_html=True)
        if not _COMISIONES_DISPONIBLE:
            st.warning("Módulo de comisiones no disponible.")
        else:
            _comisiones_lista = []
            _alerta_gmf = None
            if st.session_state.get('run') and 'df_banco' in dir() and df_banco is not None:
                _comisiones_lista = detectar_comisiones(
                    df_banco,
                    banco=_nombre_b if banco_files else "",
                    periodo=periodo_actual if 'periodo_actual' in dir() else ""
                )
                if _comisiones_lista:
                    guardar_comisiones(_comisiones_lista)
                _alerta_gmf = alertas_gmf(df_banco)

            if _alerta_gmf:
                st.warning(_alerta_gmf)

            # Resumen por tipo
            _res_com = resumen_comisiones()
            _por_tipo = _res_com.get("por_tipo", {})
            if _por_tipo:
                cols_com = st.columns(min(len(_por_tipo), 4))
                _colores_com = {"GMF":"#C62828","MANEJO":"#E65100","INTERES":"#F9A825",
                                "ACH":"#1565C0","DATAFONO":"#6A1B9A","SEGURO":"#2E7D32"}
                for i, (tipo, datos) in enumerate(_por_tipo.items()):
                    with cols_com[i % 4]:
                        _cc = _colores_com.get(tipo, "#607D8B")
                        st.markdown(f"""<div style='background:{_bg_card};border-left:4px solid {_cc};
                        border-radius:8px;padding:12px 14px;margin-bottom:8px;'>
                        <div style='font-size:1.3rem;font-weight:800;color:{_cc};'>
                          ${datos["total"]:,.0f}</div>
                        <div style='font-size:.78rem;opacity:.8;'>{tipo}</div>
                        <div style='font-size:.72rem;opacity:.6;'>{datos["count"]} movimiento(s)</div>
                        </div>""", unsafe_allow_html=True)

                _total_per = _res_com.get("total_periodo", 0)
                st.markdown(f"**💰 Total comisiones período: `${_total_per:,.0f} COP`**")

                for _alerta in _res_com.get("alertas", []):
                    st.warning(_alerta)
            else:
                st.info("No se han detectado comisiones aún. Cargue y analice un extracto bancario.")

            # Tabla detalle
            if _comisiones_lista:
                st.markdown("#### Detalle comisiones detectadas en este análisis")
                _df_com = pd.DataFrame(_comisiones_lista)[
                    ["tipo_comision","descripcion","descripcion_banco","valor","fecha_transaccion"]
                ].rename(columns={
                    "tipo_comision":"Tipo","descripcion":"Concepto DIAN",
                    "descripcion_banco":"Descripción banco",
                    "valor":"Valor COP","fecha_transaccion":"Fecha",
                })
                _df_com["Valor COP"] = _df_com["Valor COP"].apply(lambda x: f"${x:,.0f}")
                st.dataframe(_df_com, use_container_width=True, hide_index=True)

            elif not st.session_state.get('run'):
                st.info("📂 Cargue un extracto bancario y ejecute el análisis para detectar comisiones.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 11 — DIAN / XML MEDIOS MAGNÉTICOS
    # ══════════════════════════════════════════════════════════════════════════
    with tab11:
        st.markdown("<div class='section-title'>📑 DIAN — Medios Magnéticos XML</div>",
                    unsafe_allow_html=True)
        if not (_DIAN_DISPONIBLE and _tiene_permiso("dian")):
            st.warning("Módulo DIAN no disponible o sin permisos.")
        else:
            # Calendario fiscal en sidebar de esta tab
            if _ROI_DISPONIBLE:
                _cal = calendario_fiscal_colombia()
                _proximos = [v for v in _cal if v.get("proximo") and not v.get("vencido")]
                _vencidos  = [v for v in _cal if v.get("vencido") and
                              abs(v.get("dias_restantes", 999)) < 30]
                if _proximos:
                    st.warning(f"⏰ **{len(_proximos)} vencimiento(s) próximo(s) en 30 días**")
                    for v in _proximos[:3]:
                        _urgente_ico = "🔴" if v.get("urgente") else "🟡"
                        st.markdown(f"{_urgente_ico} **{v['tipo']}** — {v['fecha']} "
                                    f"({v['dias_restantes']} días) · {v['descripcion']}")
                if _vencidos:
                    st.error(f"❌ {len(_vencidos)} obligación(es) vencida(s) recientemente")

            st.markdown("---")

            col_d1, col_d2 = st.columns(2)
            with col_d1:
                _formato_dian = st.selectbox("Formato DIAN", [
                    "1008 — Pagos o abonos en cuenta",
                    "1007 — Retenciones practicadas",
                    "XML genérico (todos los movimientos)",
                ])
                _periodo_dian = st.text_input("Período (YYYY-MM)",
                                               value=periodo_actual if 'periodo_actual' in dir() else "")
            with col_d2:
                st.markdown("**Información empresa:**")
                _nit_emp = st.text_input("NIT",
                    value=WL.get("empresa_nit","") if _WL_DISPONIBLE else "")
                _nom_emp = st.text_input("Razón social",
                    value=_empresa)

            if st.button("🏛️ Generar XML DIAN", type="primary",
                         disabled=not st.session_state.get("run")):
                if not st.session_state.get("run") or 'df_banco' not in dir():
                    st.error("Primero ejecute el análisis con un extracto bancario.")
                else:
                    try:
                        if "1008" in _formato_dian:
                            _xml_b, _xml_n, _xml_h = generar_formato_1008_pagos(
                                df_banco, _periodo_dian,
                                usuario=_usuario_actual.get("username","admin"))
                        elif "1007" in _formato_dian:
                            _xml_b, _xml_n, _xml_h = generar_formato_1007_retenciones(
                                df_banco, _periodo_dian,
                                usuario=_usuario_actual.get("username","admin"))
                        else:
                            _xml_b, _xml_n, _xml_h = generar_xml_medios_magneticos(
                                df_banco, _periodo_dian,
                                usuario=_usuario_actual.get("username","admin"))

                        if _xml_b:
                            st.success(f"✅ XML generado: **{_xml_n}**")
                            st.code(f"SHA-256: {_xml_h}", language="text")
                            st.download_button("⬇️ Descargar XML DIAN", data=_xml_b,
                                               file_name=_xml_n, mime="application/xml",
                                               use_container_width=True)
                            if _NOTIF_DISPONIBLE:
                                encolar_notificacion(
                                    "DIAN_XML_LISTO", "LOG",
                                    _usuario_actual.get("username","admin"),
                                    f"XML DIAN {_xml_n} generado",
                                    f"Período {_periodo_dian} · SHA-256: {_xml_h[:16]}…"
                                )
                        else:
                            st.error("Error generando XML. Verifique los datos del extracto.")
                    except Exception as _ex:
                        st.error(f"Error: {_ex}")

            # Historial exportaciones
            st.markdown("---")
            st.markdown("#### 📋 Historial de exportaciones DIAN")
            _hist_dian = listar_exportaciones_dian(20)
            if _hist_dian:
                _df_dian = pd.DataFrame(_hist_dian)
                st.dataframe(_df_dian, use_container_width=True, hide_index=True)
            else:
                st.info("No hay exportaciones DIAN registradas.")

            # Calendario fiscal completo
            if _ROI_DISPONIBLE:
                with st.expander("📅 Calendario Fiscal Colombia (año actual)", expanded=False):
                    for v in _cal:
                        _ico = ("✅" if not v.get("vencido") and not v.get("proximo")
                                else ("❌" if v.get("vencido") else "⚠️"))
                        _urgente = " 🔴" if v.get("urgente") else ""
                        st.markdown(
                            f"{_ico} **{v['fecha']}** · **{v['tipo']}**{_urgente} — "
                            f"{v['descripcion']}"
                        )

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 12 — PUC (PLAN ÚNICO DE CUENTAS)
    # ══════════════════════════════════════════════════════════════════════════
    with tab12:
        st.markdown("<div class='section-title'>📚 Plan Único de Cuentas (PUC)</div>",
                    unsafe_allow_html=True)
        if not _PUC_DISPONIBLE:
            st.warning("Módulo PUC no disponible.")
        else:
            if st.session_state.get("run") and 'df_banco' in dir() and df_banco is not None:
                _df_puc = enriquecer_dataframe_con_puc(df_banco)
                _cols_puc = [c for c in ["FECHA_RAW","DESCRIPCION","VALOR","TIPO",
                                          "PUC","CUENTA_PUC","NATURALEZA_PUC"]
                             if c in _df_puc.columns]
                st.markdown("#### Extracto bancario con clasificación PUC automática")
                st.dataframe(_df_puc[_cols_puc].head(100),
                             use_container_width=True, hide_index=True)

                # Botón descargar con PUC
                _buf_puc = io.BytesIO()
                _df_puc.to_excel(_buf_puc, index=False, engine='openpyxl')
                _buf_puc.seek(0)
                st.download_button("⬇️ Descargar Excel con PUC",
                                   data=_buf_puc,
                                   file_name=f"extracto_puc_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # Aprendizaje manual
                if _tiene_permiso("config"):
                    st.markdown("---")
                    st.markdown("#### ✏️ Corregir clasificación PUC")
                    col_pa, col_pb = st.columns(2)
                    with col_pa:
                        _desc_corr = st.text_input("Descripción bancaria a reclasificar")
                    with col_pb:
                        _cat_puc = listar_catalogo_puc()
                        _cod_opts = [f"{c['codigo']} — {c['nombre']}" for c in _cat_puc]
                        _cod_sel  = st.selectbox("Cuenta PUC correcta", _cod_opts)
                    if st.button("💾 Guardar clasificación"):
                        _cod = _cod_sel.split(" — ")[0] if _cod_sel else ""
                        if aprender_clasificacion(_desc_corr, _cod,
                                                   usuario=_usuario_actual.get("username","admin")):
                            st.success(f"✅ Clasificación guardada: '{_desc_corr}' → {_cod}")
                        else:
                            st.error("Error guardando clasificación")
            else:
                st.info("📂 Cargue y analice un extracto para ver la clasificación PUC automática.")

            # Catálogo PUC completo
            with st.expander("📖 Ver catálogo PUC completo", expanded=False):
                _df_cat = pd.DataFrame(listar_catalogo_puc())
                st.dataframe(_df_cat, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 13 — ML PREDICTOR
    # ══════════════════════════════════════════════════════════════════════════
    with tab13:
        st.markdown("<div class='section-title'>🤖 ML — Predictor de Partidas</div>",
                    unsafe_allow_html=True)
        if not _ML_DISPONIBLE:
            st.warning("Módulo ML no disponible.")
        else:
            _acc = accuracy_modelo()
            cm1, cm2, cm3 = st.columns(3)
            with cm1:
                st.metric("Predicciones totales", _acc.get("total", 0))
            with cm2:
                st.metric("Confirmadas", _acc.get("confirmadas", 0))
            with cm3:
                st.metric("Precisión del modelo", f"{_acc.get('accuracy', 0):.1f}%")

            st.markdown("---")

            col_ml1, col_ml2 = st.columns(2)
            with col_ml1:
                _base_per  = st.text_input("Período base (YYYY-MM)",
                                            value=periodo_actual if 'periodo_actual' in dir() else "")
            with col_ml2:
                _pred_per  = st.text_input("Período a predecir (YYYY-MM)")

            if st.button("🔮 Generar predicciones", type="primary"):
                if not _base_per or not _pred_per:
                    st.error("Ingrese ambos períodos")
                else:
                    with st.spinner("Analizando historial de partidas..."):
                        _preds = predecir_partidas_proximas(_base_per, _pred_per)
                    if _preds:
                        st.success(f"✅ {len(_preds)} partidas predichas para {_pred_per}")
                        _df_ml = pd.DataFrame(_preds)[
                            ["tipo","descripcion","valor_estimado","confianza","apariciones"]
                        ].rename(columns={
                            "tipo":"Tipo","descripcion":"Descripción",
                            "valor_estimado":"Valor estimado COP",
                            "confianza":"Confianza","apariciones":"Veces visto",
                        })
                        _df_ml["Valor estimado COP"] = _df_ml["Valor estimado COP"].apply(
                            lambda x: f"${x:,.0f}")
                        _df_ml["Confianza"] = _df_ml["Confianza"].apply(
                            lambda x: f"{x*100:.0f}%")
                        st.dataframe(_df_ml, use_container_width=True, hide_index=True)
                    else:
                        st.info("No hay suficiente historial para generar predicciones. "
                                "Registre más partidas conciliatorias en los meses anteriores.")

            # Predicciones guardadas
            st.markdown("---")
            st.markdown("#### 📋 Predicciones guardadas")
            _pred_guard = listar_predicciones(_pred_per if '_pred_per' in dir() else None)
            if _pred_guard:
                for _pg in _pred_guard:
                    _conf_color = ("#2E7D32" if _pg["confianza"] >= 0.7
                                   else "#F9A825" if _pg["confianza"] >= 0.4 else "#C62828")
                    _conf_pct   = f"{_pg['confianza']*100:.0f}%"
                    _confirm_ico = "✅" if _pg.get("confirmado") else "⏳"
                    st.markdown(
                        f"{_confirm_ico} **{_pg['descripcion'][:50]}** — "
                        f"${_pg['valor']:,.0f} | "
                        f"<span style='color:{_conf_color};font-weight:700;'>{_conf_pct} confianza</span>",
                        unsafe_allow_html=True
                    )
            else:
                st.info("No hay predicciones guardadas para el período seleccionado.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 14 — ADMINISTRACIÓN
    # ══════════════════════════════════════════════════════════════════════════
    with tab14:
        st.markdown("<div class='section-title'>⚙️ Panel de Administración</div>",
                    unsafe_allow_html=True)
        if not _tiene_permiso("usuarios"):
            st.warning("🔒 Acceso restringido — se requiere rol Administrador.")
        else:
            _admin_tab = st.tabs(["👥 Usuarios", "🔍 Auditoría", "🏢 White Label",
                                   "💾 Backup", "📈 ROI del Sistema"])

            # ── Sub-tab Usuarios ───────────────────────────────────────────────
            with _admin_tab[0]:
                if not _AUTH_DISPONIBLE:
                    st.warning("Módulo auth no disponible.")
                else:
                    st.markdown("##### Usuarios del sistema")
                    _df_usr = pd.DataFrame(listar_usuarios())
                    if not _df_usr.empty:
                        _df_usr["activo"] = _df_usr["activo"].map({True:"✅ Activo", False:"❌ Inactivo"})
                        st.dataframe(_df_usr[["username","nombre","rol","activo",
                                               "creado","ultimo_acceso"]],
                                     use_container_width=True, hide_index=True)

                    st.markdown("---")
                    st.markdown("##### ➕ Crear nuevo usuario")
                    cu1, cu2 = st.columns(2)
                    with cu1:
                        _nu_user = st.text_input("Username", key="nu_user")
                        _nu_pwd  = st.text_input("Contraseña", type="password", key="nu_pwd")
                        _nu_nom  = st.text_input("Nombre completo", key="nu_nom")
                    with cu2:
                        _nu_rol  = st.selectbox("Rol", list(ROLES.keys()),
                                                 format_func=lambda x: ROLES[x]["label"])
                        _nu_email = st.text_input("Email", key="nu_email")
                    if st.button("💾 Crear usuario", type="primary"):
                        ok, msg = crear_usuario(
                            _nu_user, _nu_pwd, _nu_rol, _nu_nom, _nu_email,
                            creado_por=_usuario_actual.get("username","admin")
                        )
                        if ok:
                            st.success(f"✅ {msg}")
                            registrar_auditoria(_usuario_actual.get("username","admin"),
                                                "CREAR_USUARIO", "admin",
                                                f"Usuario {_nu_user} ({_nu_rol}) creado")
                        else:
                            st.error(f"❌ {msg}")

                    st.markdown("---")
                    st.markdown("##### 🔑 Cambiar contraseña")
                    cp1, cp2 = st.columns(2)
                    with cp1:
                        _cp_user = st.text_input("Usuario a modificar", key="cp_user")
                    with cp2:
                        _cp_pwd  = st.text_input("Nueva contraseña", type="password", key="cp_pwd")
                    if st.button("🔑 Actualizar contraseña"):
                        ok, msg = cambiar_password(_cp_user, _cp_pwd,
                                                    por=_usuario_actual.get("username","admin"))
                        st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")

            # ── Sub-tab Auditoría ──────────────────────────────────────────────
            with _admin_tab[1]:
                if not _AUTH_DISPONIBLE:
                    st.warning("Módulo auth no disponible.")
                else:
                    _audit_usr = st.text_input("Filtrar por usuario (vacío = todos)")
                    _audit_rows = listar_auditoria(100, _audit_usr or None)
                    if _audit_rows:
                        _df_audit = pd.DataFrame(_audit_rows)
                        st.dataframe(_df_audit, use_container_width=True, hide_index=True)
                    else:
                        st.info("No hay registros de auditoría.")

            # ── Sub-tab White Label ────────────────────────────────────────────
            with _admin_tab[2]:
                if not (_WL_DISPONIBLE and _tiene_permiso("config")):
                    st.warning("Sin permisos para configurar white label.")
                else:
                    st.markdown("##### 🎨 Personalización de marca")
                    _all_cfg = WL.get_all()
                    wl1, wl2 = st.columns(2)
                    with wl1:
                        _wl_nom  = st.text_input("Nombre empresa",
                                                   value=_all_cfg.get("empresa_nombre",""),
                                                   key="wl_empresa_nombre")
                        _wl_nit  = st.text_input("NIT",
                                                   value=_all_cfg.get("empresa_nit",""),
                                                   key="wl_empresa_nit")
                        _wl_c1   = st.color_picker("Color corporativo primario",
                                                     value=_all_cfg.get("empresa_color_primario","#1F4E79"),
                                                     key="wl_color_primario")
                    with wl2:
                        _wl_ciudad = st.text_input("Ciudad",
                                                    value=_all_cfg.get("empresa_ciudad",""),
                                                    key="wl_empresa_ciudad")
                        _wl_plan   = st.selectbox("Plan activo",
                                                   ["starter","profesional","empresarial"],
                                                   index=["starter","profesional","empresarial"].index(
                                                       _all_cfg.get("plan_activo","profesional")),
                                                   key="wl_plan_activo")
                        _wl_tema   = st.selectbox("Tema UI",["oscuro","claro"],
                                                   index=0 if _all_cfg.get("tema_default","oscuro")=="oscuro" else 1,
                                                   key="wl_tema_ui")
                        _wl_c2     = st.color_picker("Color acento",
                                                      value=_all_cfg.get("empresa_color_secundario","#C9A227"),
                                                      key="wl_color_secundario")
                    if st.button("💾 Guardar configuración de marca", type="primary", key="wl_marca"):
                        _campos_m = {
                            "empresa_nombre": _wl_nom, "empresa_nit": _wl_nit,
                            "empresa_ciudad": _wl_ciudad,
                            "empresa_color_primario": _wl_c1,
                            "empresa_color_secundario": _wl_c2,
                            "plan_activo": _wl_plan, "tema_default": _wl_tema,
                        }
                        _ok_all = all(WL.set_config(k, v,
                                       usuario=_usuario_actual.get("username","admin"))
                                      for k, v in _campos_m.items())
                        if _ok_all:
                            st.success("✅ Configuración guardada — recargue la app")
                            registrar_auditoria(_usuario_actual.get("username","admin"),
                                                "CONFIG_WHITE_LABEL", "admin",
                                                f"Empresa: {_wl_nom} | Tema: {_wl_tema}")
                        else:
                            st.error("Error guardando configuración")

                    st.markdown("---")
                    st.markdown("##### 🖊️ Firmantes — Conciliación Formal y PDF")
                    _f1, _f2 = st.columns(2)
                    with _f1:
                        _wl_rl   = st.text_input("Nombre Representante Legal",
                                                   value=_all_cfg.get("empresa_rep_legal",""),
                                                   key="wl_rep_legal")
                        _wl_rl_cc= st.text_input("C.C. Representante Legal",
                                                   value=_all_cfg.get("empresa_rep_legal_cc",""),
                                                   key="wl_rep_legal_cc")
                    with _f2:
                        _wl_ct   = st.text_input("Nombre Contador Público",
                                                   value=_all_cfg.get("empresa_contador",""),
                                                   key="wl_contador")
                        _wl_tp   = st.text_input("T.P. Contador",
                                                   value=_all_cfg.get("empresa_tp_contador",""),
                                                   key="wl_tp_contador")
                    st.markdown("---")
                    st.markdown("##### 🏦 Cuenta bancaria principal")
                    _b1, _b2 = st.columns(2)
                    with _b1:
                        _wl_banco= st.text_input("Banco principal",
                                                   value=_all_cfg.get("empresa_banco_principal",""),
                                                   placeholder="Ej: Bancolombia",
                                                   key="wl_banco_principal")
                    with _b2:
                        _wl_cta  = st.text_input("Cuenta bancaria",
                                                   value=_all_cfg.get("empresa_cuenta_bancaria",""),
                                                   placeholder="Ej: 1105-01 Caja",
                                                   key="wl_cuenta_bancaria")
                    if st.button("💾 Guardar firmantes y banco", type="primary", key="wl_firmantes"):
                        _campos_f = {
                            "empresa_rep_legal":       _wl_rl,
                            "empresa_rep_legal_cc":    _wl_rl_cc,
                            "empresa_contador":        _wl_ct,
                            "empresa_tp_contador":     _wl_tp,
                            "empresa_banco_principal": _wl_banco,
                            "empresa_cuenta_bancaria": _wl_cta,
                        }
                        _ok_f = all(WL.set_config(k, v,
                                     usuario=_usuario_actual.get("username","admin"))
                                    for k, v in _campos_f.items())
                        if _ok_f:
                            st.success("✅ Firmantes guardados — aparecen en Tab 6 y PDF firmado")
                            registrar_auditoria(_usuario_actual.get("username","admin"),
                                                "CONFIG_FIRMANTES", "admin",
                                                f"RL: {_wl_rl} | CT: {_wl_ct}")
                        else:
                            st.error("Error guardando firmantes")
                    st.caption("💡 Los firmantes aparecen en la Conciliación Formal (Tab 6) y en el PDF firmado.")

            # ── Sub-tab Backup ─────────────────────────────────────────────────
            with _admin_tab[3]:
                st.markdown("##### 💾 Backup automático")
                _carpeta_bk = WL.get("backup_carpeta","./backups") if _WL_DISPONIBLE else "./backups"
                st.info(f"Carpeta de backup: `{_carpeta_bk}`")
                c_bk1, c_bk2 = st.columns(2)
                with c_bk1:
                    if st.button("▶️ Ejecutar backup ahora", use_container_width=True):
                        if _NOTIF_DISPONIBLE:
                            import os as _os
                            _db_path = DB_PATH
                            encolar_backup(_db_path, _carpeta_bk, tipo="LOCAL")
                            _res_bk = procesar_cola_backup(_carpeta_bk)
                            st.success(f"✅ Backup: {_res_bk.get('copiados',0)} archivos copiados")
                        else:
                            st.error("Módulo de backup no disponible")
                with c_bk2:
                    if st.button("📤 Procesar cola notificaciones", use_container_width=True):
                        if _NOTIF_DISPONIBLE:
                            _pend_n = listar_notificaciones_pendientes()
                            _res_n = procesar_cola_notificaciones()
                            st.success(f"✅ {_res_n.get('enviadas',0)} enviadas | "
                                       f"{_res_n.get('fallidas',0)} fallidas")
                        else:
                            st.error("Módulo de notificaciones no disponible")

                if _NOTIF_DISPONIBLE:
                    _pend_n2 = listar_notificaciones_pendientes()
                    if _pend_n2:
                        st.markdown(f"**📬 {len(_pend_n2)} notificación(es) en cola:**")
                        for _nn in _pend_n2[:5]:
                            st.caption(f"[{_nn['canal']}] {_nn['tipo']} → {_nn['dest']} "
                                       f"({_nn['intentos']} intentos)")

            # ── Sub-tab ROI ────────────────────────────────────────────────────
            with _admin_tab[4]:
                if not _ROI_DISPONIBLE:
                    st.warning("Módulo ROI no disponible.")
                else:
                    st.markdown("##### 📈 ROI del Sistema — Ahorro mensual")
                    _roi_mes = roi_acumulado_mes()
                    if _roi_mes:
                        ra1, ra2, ra3, ra4 = st.columns(4)
                        with ra1:
                            st.metric("Conciliaciones este mes",
                                      _roi_mes.get("n_conciliaciones_mes", 0))
                        with ra2:
                            st.metric("Horas ahorradas",
                                      f"{_roi_mes.get('horas_ahorradas',0):.1f}h")
                        with ra3:
                            st.metric("Ahorro estimado",
                                      f"${_roi_mes.get('pesos_ahorrados',0):,.0f}")
                        with ra4:
                            st.metric("Automatización",
                                      f"{_roi_mes.get('pct_automatizacion',0):.1f}%")

                        st.info(_roi_mes.get("mensaje", ""))

                    # ROI del análisis actual
                    if st.session_state.get("run") and 'df_banco' in dir():
                        st.markdown("---")
                        st.markdown("##### ROI de este análisis")
                        _roi_now = calcular_roi(
                            n_movimientos_banco=len(df_banco) if df_banco is not None else 0,
                            n_movimientos_aux=len(df_aux) if df_aux is not None else 0,
                            n_conciliados=n_exac + n_apr + n_agr if 'n_exac' in dir() else 0,
                            n_rechazos=n_rec if 'n_rec' in dir() else 0,
                        )
                        rc1, rc2, rc3 = st.columns(3)
                        with rc1:
                            st.metric("Tiempo manual estimado",
                                      f"{_roi_now['horas_manual_est']:.1f}h")
                        with rc2:
                            st.metric("Tiempo real del sistema",
                                      f"{_roi_now['horas_sistema']*60:.0f} min")
                        with rc3:
                            st.metric("Ahorro",
                                      f"${_roi_now['pesos_ahorrados']:,.0f} COP")
                        st.info(_roi_now.get("mensaje", ""))

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 12 — PUC (PLAN ÚNICO DE CUENTAS)
    # ══════════════════════════════════════════════════════════════════════════
    with tab12:
        st.markdown("<div class='section-title'>📚 Plan Único de Cuentas (PUC)</div>",
                    unsafe_allow_html=True)
        if not _PUC_DISPONIBLE:
            st.warning("Módulo PUC no disponible.")
        else:
            if st.session_state.get("run") and 'df_banco' in dir() and df_banco is not None:
                _df_puc = enriquecer_dataframe_con_puc(df_banco)
                _cols_puc = [c for c in ["FECHA_RAW","DESCRIPCION","VALOR","TIPO",
                                          "PUC","CUENTA_PUC","NATURALEZA_PUC"]
                             if c in _df_puc.columns]
                st.markdown("#### Extracto bancario con clasificación PUC automática")
                st.dataframe(_df_puc[_cols_puc].head(100),
                             use_container_width=True, hide_index=True)

                # Botón descargar con PUC
                _buf_puc = io.BytesIO()
                _df_puc.to_excel(_buf_puc, index=False, engine='openpyxl')
                _buf_puc.seek(0)
                st.download_button("⬇️ Descargar Excel con PUC",
                                   data=_buf_puc,
                                   file_name=f"extracto_puc_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # Aprendizaje manual
                if _tiene_permiso("config"):
                    st.markdown("---")
                    st.markdown("#### ✏️ Corregir clasificación PUC")
                    col_pa, col_pb = st.columns(2)
                    with col_pa:
                        _desc_corr = st.text_input("Descripción bancaria a reclasificar")
                    with col_pb:
                        _cat_puc = listar_catalogo_puc()
                        _cod_opts = [f"{c['codigo']} — {c['nombre']}" for c in _cat_puc]
                        _cod_sel  = st.selectbox("Cuenta PUC correcta", _cod_opts)
                    if st.button("💾 Guardar clasificación"):
                        _cod = _cod_sel.split(" — ")[0] if _cod_sel else ""
                        if aprender_clasificacion(_desc_corr, _cod,
                                                   usuario=_usuario_actual.get("username","admin")):
                            st.success(f"✅ Clasificación guardada: '{_desc_corr}' → {_cod}")
                        else:
                            st.error("Error guardando clasificación")
            else:
                st.info("📂 Cargue y analice un extracto para ver la clasificación PUC automática.")

            # Catálogo PUC completo
            with st.expander("📖 Ver catálogo PUC completo", expanded=False):
                _df_cat = pd.DataFrame(listar_catalogo_puc())
                st.dataframe(_df_cat, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 13 — ML PREDICTOR
    # ══════════════════════════════════════════════════════════════════════════
    with tab13:
        st.markdown("<div class='section-title'>🤖 ML — Predictor de Partidas</div>",
                    unsafe_allow_html=True)
        if not _ML_DISPONIBLE:
            st.warning("Módulo ML no disponible.")
        else:
            _acc = accuracy_modelo()
            cm1, cm2, cm3 = st.columns(3)
            with cm1:
                st.metric("Predicciones totales", _acc.get("total", 0))
            with cm2:
                st.metric("Confirmadas", _acc.get("confirmadas", 0))
            with cm3:
                st.metric("Precisión del modelo", f"{_acc.get('accuracy', 0):.1f}%")

            st.markdown("---")

            col_ml1, col_ml2 = st.columns(2)
            with col_ml1:
                _base_per  = st.text_input("Período base (YYYY-MM)",
                                            value=periodo_actual if 'periodo_actual' in dir() else "")
            with col_ml2:
                _pred_per  = st.text_input("Período a predecir (YYYY-MM)")

            if st.button("🔮 Generar predicciones", type="primary"):
                if not _base_per or not _pred_per:
                    st.error("Ingrese ambos períodos")
                else:
                    with st.spinner("Analizando historial de partidas..."):
                        _preds = predecir_partidas_proximas(_base_per, _pred_per)
                    if _preds:
                        st.success(f"✅ {len(_preds)} partidas predichas para {_pred_per}")
                        _df_ml = pd.DataFrame(_preds)[
                            ["tipo","descripcion","valor_estimado","confianza","apariciones"]
                        ].rename(columns={
                            "tipo":"Tipo","descripcion":"Descripción",
                            "valor_estimado":"Valor estimado COP",
                            "confianza":"Confianza","apariciones":"Veces visto",
                        })
                        _df_ml["Valor estimado COP"] = _df_ml["Valor estimado COP"].apply(
                            lambda x: f"${x:,.0f}")
                        _df_ml["Confianza"] = _df_ml["Confianza"].apply(
                            lambda x: f"{x*100:.0f}%")
                        st.dataframe(_df_ml, use_container_width=True, hide_index=True)
                    else:
                        st.info("No hay suficiente historial para generar predicciones. "
                                "Registre más partidas conciliatorias en los meses anteriores.")

            # Predicciones guardadas
            st.markdown("---")
            st.markdown("#### 📋 Predicciones guardadas")
            _pred_guard = listar_predicciones(_pred_per if '_pred_per' in dir() else None)
            if _pred_guard:
                for _pg in _pred_guard:
                    _conf_color = ("#2E7D32" if _pg["confianza"] >= 0.7
                                   else "#F9A825" if _pg["confianza"] >= 0.4 else "#C62828")
                    _conf_pct   = f"{_pg['confianza']*100:.0f}%"
                    _confirm_ico = "✅" if _pg.get("confirmado") else "⏳"
                    st.markdown(
                        f"{_confirm_ico} **{_pg['descripcion'][:50]}** — "
                        f"${_pg['valor']:,.0f} | "
                        f"<span style='color:{_conf_color};font-weight:700;'>{_conf_pct} confianza</span>",
                        unsafe_allow_html=True
                    )
            else:
                st.info("No hay predicciones guardadas para el período seleccionado.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 14 — ADMINISTRACIÓN
    # ══════════════════════════════════════════════════════════════════════════
    with tab14:
        st.markdown("<div class='section-title'>⚙️ Panel de Administración</div>",
                    unsafe_allow_html=True)
        if not _tiene_permiso("usuarios"):
            st.warning("🔒 Acceso restringido — se requiere rol Administrador.")
        else:
            _admin_tab = st.tabs(["👥 Usuarios", "🔍 Auditoría", "🏢 White Label",
                                   "💾 Backup", "📈 ROI del Sistema"])

            # ── Sub-tab Usuarios ───────────────────────────────────────────────
            with _admin_tab[0]:
                if not _AUTH_DISPONIBLE:
                    st.warning("Módulo auth no disponible.")
                else:
                    st.markdown("##### Usuarios del sistema")
                    _df_usr = pd.DataFrame(listar_usuarios())
                    if not _df_usr.empty:
                        _df_usr["activo"] = _df_usr["activo"].map({True:"✅ Activo", False:"❌ Inactivo"})
                        st.dataframe(_df_usr[["username","nombre","rol","activo",
                                               "creado","ultimo_acceso"]],
                                     use_container_width=True, hide_index=True)

                    st.markdown("---")
                    st.markdown("##### ➕ Crear nuevo usuario")
                    cu1, cu2 = st.columns(2)
                    with cu1:
                        _nu_user = st.text_input("Username", key="nu_user")
                        _nu_pwd  = st.text_input("Contraseña", type="password", key="nu_pwd")
                        _nu_nom  = st.text_input("Nombre completo", key="nu_nom")
                    with cu2:
                        _nu_rol  = st.selectbox("Rol", list(ROLES.keys()),
                                                 format_func=lambda x: ROLES[x]["label"])
                        _nu_email = st.text_input("Email", key="nu_email")
                    if st.button("💾 Crear usuario", type="primary"):
                        ok, msg = crear_usuario(
                            _nu_user, _nu_pwd, _nu_rol, _nu_nom, _nu_email,
                            creado_por=_usuario_actual.get("username","admin")
                        )
                        if ok:
                            st.success(f"✅ {msg}")
                            registrar_auditoria(_usuario_actual.get("username","admin"),
                                                "CREAR_USUARIO", "admin",
                                                f"Usuario {_nu_user} ({_nu_rol}) creado")
                        else:
                            st.error(f"❌ {msg}")

                    st.markdown("---")
                    st.markdown("##### 🔑 Cambiar contraseña")
                    cp1, cp2 = st.columns(2)
                    with cp1:
                        _cp_user = st.text_input("Usuario a modificar", key="cp_user")
                    with cp2:
                        _cp_pwd  = st.text_input("Nueva contraseña", type="password", key="cp_pwd")
                    if st.button("🔑 Actualizar contraseña"):
                        ok, msg = cambiar_password(_cp_user, _cp_pwd,
                                                    por=_usuario_actual.get("username","admin"))
                        st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")

            # ── Sub-tab Auditoría ──────────────────────────────────────────────
            with _admin_tab[1]:
                if not _AUTH_DISPONIBLE:
                    st.warning("Módulo auth no disponible.")
                else:
                    _audit_usr = st.text_input("Filtrar por usuario (vacío = todos)")
                    _audit_rows = listar_auditoria(100, _audit_usr or None)
                    if _audit_rows:
                        _df_audit = pd.DataFrame(_audit_rows)
                        st.dataframe(_df_audit, use_container_width=True, hide_index=True)
                    else:
                        st.info("No hay registros de auditoría.")

            # ── Sub-tab White Label ────────────────────────────────────────────
            with _admin_tab[2]:
                if not (_WL_DISPONIBLE and _tiene_permiso("config")):
                    st.warning("Sin permisos para configurar white label.")
                else:
                    st.markdown("##### 🎨 Personalización de marca")
                    _all_cfg = WL.get_all()
                    wl1, wl2 = st.columns(2)
                    with wl1:
                        _wl_nom  = st.text_input("Nombre empresa",
                                                   value=_all_cfg.get("empresa_nombre",""),
                                                   key="wl_empresa_nombre")
                        _wl_nit  = st.text_input("NIT",
                                                   value=_all_cfg.get("empresa_nit",""),
                                                   key="wl_empresa_nit")
                        _wl_c1   = st.color_picker("Color corporativo primario",
                                                     value=_all_cfg.get("empresa_color_primario","#1F4E79"),
                                                     key="wl_color_primario")
                    with wl2:
                        _wl_ciudad = st.text_input("Ciudad",
                                                    value=_all_cfg.get("empresa_ciudad",""),
                                                    key="wl_empresa_ciudad")
                        _wl_plan   = st.selectbox("Plan activo",
                                                   ["starter","profesional","empresarial"],
                                                   index=["starter","profesional","empresarial"].index(
                                                       _all_cfg.get("plan_activo","profesional")),
                                                   key="wl_plan_activo")
                        _wl_tema   = st.selectbox("Tema UI",["oscuro","claro"],
                                                   index=0 if _all_cfg.get("tema_default","oscuro")=="oscuro" else 1,
                                                   key="wl_tema_ui")
                        _wl_c2     = st.color_picker("Color acento",
                                                      value=_all_cfg.get("empresa_color_secundario","#C9A227"),
                                                      key="wl_color_secundario")
                    if st.button("💾 Guardar configuración de marca", type="primary", key="wl_marca"):
                        _campos_m = {
                            "empresa_nombre": _wl_nom, "empresa_nit": _wl_nit,
                            "empresa_ciudad": _wl_ciudad,
                            "empresa_color_primario": _wl_c1,
                            "empresa_color_secundario": _wl_c2,
                            "plan_activo": _wl_plan, "tema_default": _wl_tema,
                        }
                        _ok_all = all(WL.set_config(k, v,
                                       usuario=_usuario_actual.get("username","admin"))
                                      for k, v in _campos_m.items())
                        if _ok_all:
                            st.success("✅ Configuración guardada — recargue la app")
                            registrar_auditoria(_usuario_actual.get("username","admin"),
                                                "CONFIG_WHITE_LABEL", "admin",
                                                f"Empresa: {_wl_nom} | Tema: {_wl_tema}")
                        else:
                            st.error("Error guardando configuración")

                    st.markdown("---")
                    st.markdown("##### 🖊️ Firmantes — Conciliación Formal y PDF")
                    _f1, _f2 = st.columns(2)
                    with _f1:
                        _wl_rl   = st.text_input("Nombre Representante Legal",
                                                   value=_all_cfg.get("empresa_rep_legal",""),
                                                   key="wl_rep_legal")
                        _wl_rl_cc= st.text_input("C.C. Representante Legal",
                                                   value=_all_cfg.get("empresa_rep_legal_cc",""),
                                                   key="wl_rep_legal_cc")
                    with _f2:
                        _wl_ct   = st.text_input("Nombre Contador Público",
                                                   value=_all_cfg.get("empresa_contador",""),
                                                   key="wl_contador")
                        _wl_tp   = st.text_input("T.P. Contador",
                                                   value=_all_cfg.get("empresa_tp_contador",""),
                                                   key="wl_tp_contador")
