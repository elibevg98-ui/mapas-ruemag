# ============================================================
# MAPAS RUEMAG
# Aplicación local en Python, Anaconda y Spyder
# ============================================================

import os
import subprocess
import sys


# ============================================================
# ABRIR AUTOMÁTICAMENTE DESDE SPYDER
# ============================================================

if os.environ.get("RUEMAG_STREAMLIT") != "ACTIVO":

    archivo_actual = os.path.abspath(__file__)

    entorno = os.environ.copy()
    entorno["RUEMAG_STREAMLIT"] = "ACTIVO"

    subprocess.Popen(
        [
            sys.executable,
            "-m",
            "streamlit",
            "run",
            archivo_actual
        ],
        env=entorno
    )

    raise SystemExit


# ============================================================
# IMPORTACIONES
# ============================================================

import io
import math
import textwrap
import tempfile
import unicodedata
import zipfile
import datetime 
from pathlib import Path
from xml.sax.saxutils import escape

import matplotlib.image as mpimg
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

from matplotlib.colors import LinearSegmentedColormap, to_hex, to_rgb
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from matplotlib.path import Path as RutaMatplotlib
from matplotlib.ticker import FormatStrFormatter 


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Mapas RUEMAG",
    page_icon="☁️",
    layout="wide"
)


RUTA_EXCEL = r"G:\RUEMAG2\ruemag2base.xlsx"

RUTA_LIMITE = r"G:\datos25\limite_gdl.gpkg"

RUTA_COLONIAS = r"C:\Users\bcomi\Documents\SHAPEFILES\colonias_2022"

RUTA_LOGO = r"G:\RUEMAG2\logo_inteligencia.png"

RUTA_ROSA = r"G:\RUEMAG2\rosa_vientos.png"

RUTA_CLIMATOLOGIA = (
    r"G:\RUEMAG2\paginamapas\gdl_completo_precip.csv"
)

RUTA_RACHAS_HISTORICAS = (
    r"G:\RUEMAG2\paginamapas\seriecompletaiam+ruemagcentro.xlsx"
)

# Carpeta de Eventos y ruta directa del excel general
RUTA_EVENTOS = r"G:\RUEMAG2\TEMPORAL 2026\ACTUALIZACION_JULIO"
RUTA_HISTORICO_EVENTOS = r"G:\RUEMAG2\TEMPORAL 2026\eventos.xlsx"


# ============================================================
# VARIABLES
# ============================================================

VARIABLES = {

    "T_Avg": {
        "nombre": "Temperatura media",
        "unidad": "°C",
        "colores": [
            "#e8f2ff",
            "#83b7ee",
            "#f2b1b5",
            "#e85a3f"
        ]
    },

    "T_Max": {
        "nombre": "Temperatura máxima",
        "unidad": "°C",
        "colores": [
            "#fff1e9",
            "#ffad79",
            "#ef643f",
            "#a92d1f"
        ]
    },

    "T_Min": {
        "nombre": "Temperatura mínima",
        "unidad": "°C",
        "colores": [
            "#eaf3ff",
            "#a6cdf5",
            "#4d93e6",
            "#1757a6"
        ]
    },

    "RH_Avg": {
        "nombre": "Humedad relativa media",
        "unidad": "%",
        "colores": [
            "#f1f7ec",
            "#a8d9c0",
            "#48a7a1",
            "#19596c"
        ]
    },

    "P_Avg": {
        "nombre": "Presión media",
        "unidad": "hPa",
        "colores": [
            "#f4f0f7",
            "#c8aad5",
            "#8353a3",
            "#41215e"
        ]
    },

    "Rain_Total": {
        "nombre": "Precipitación acumulada",
        "unidad": "mm",
        "colores": [
            "#eff3ff",
            "#9ecae1",
            "#3182bd",
            "#08306b"
        ]
    },

    "Wind_Avg": {
        "nombre": "Viento promedio y dirección predominante",
        "unidad": "km/h",
        "colores": [
            "#f1eff7",
            "#c8c3df",
            "#9489bd",
            "#724fa5",
            "#4d187c"
        ]
    },

    "Wind_Max": {
        "nombre": "Viento máximo y dirección predominante",
        "unidad": "km/h",
        "colores": [
            "#f1eff7",
            "#c8c3df",
            "#9489bd",
            "#724fa5",
            "#4d187c"
        ]
    },

    "Gust_Max": {
        "nombre": "Ráfaga máxima",
        "unidad": "km/h",
        "colores": [
            "#e1f4ef",
            "#8dcac5",
            "#3f949e",
            "#245761"
        ]
    }
}


COLUMNAS_METEOROLOGICAS = (
    list(VARIABLES.keys())
    +
    ["Wind_Dir"]
)


NOMBRES_ESTACIONES = {

    "archivo": "Expo-GDL",
    "base1": "Agua Azul",
    "base2": "Tránsito",
    "base3": "Zona Industrial",
    "base4": "San Andrés",
    "base5": "Periférico Norte",
    "colegio": "Providencia-Acueducto",
    "colmena": "Oblatos",
    "crit": "CRIT",
    "dean": "DEAN",
    "estadio": "Estadio",
    "mercado": "Centro"

}


# ============================================================
# ESTILOS DE LA PÁGINA
# ============================================================



# ============================================================
# TARJETAS DE ESTADÍSTICAS
# ============================================================

st.markdown('<style id="ruemag-profesional">\n@import url(\'https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&display=swap\');\n:root {\n  color-scheme: light;\n  --ru-blanco: #ffffff; --ru-fondo: #f2f5f8; --ru-suave: #f7f9fb;\n  --ru-tinta: #172b40; --ru-secundario: #526476; --ru-linea: #e0e6ed; --ru-control: #7e8c9b;\n  --ru-vino: #84273f; --ru-vino-fuerte: #641c31; --ru-vino-suave: #f8f0f2;\n  --ru-exito: #1b6648; --ru-aviso: #855300; --ru-error: #b12838;\n  --ru-fuente: \'Source Sans 3\', \'Segoe UI\', sans-serif;\n  --ru-texto: 1rem; --ru-pequeno: .875rem; --ru-h1: clamp(1.5rem, 2.2vw, 2rem);\n  --ru-h2: 1.375rem; --ru-h3: 1.125rem; --ru-cifra: clamp(1.625rem, 2.5vw, 2rem);\n  --ru-espacio-1: .25rem; --ru-espacio-2: .5rem; --ru-espacio-3: .75rem;\n  --ru-espacio-4: 1rem; --ru-espacio-6: 1.5rem; --ru-espacio-8: 2rem;\n  --ru-radio-control: 6px; --ru-radio-panel: 10px; --ru-radio-hoja: 16px;\n  --ru-tacto: 44px; --ru-sombra: none; --ru-sombra-flotante: 0 8px 24px rgb(23 43 64 / .12);\n}\nhtml, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] { background: var(--ru-fondo); color: var(--ru-tinta); }\n.stApp {\n  --text-color: var(--ru-tinta); --background-color: var(--ru-blanco);\n  --secondary-background-color: var(--ru-suave); --primary-color: var(--ru-vino);\n  font-family: var(--ru-fuente); font-size: var(--ru-texto); line-height: 1.5;\n}\n.stApp p, .stApp label, .stApp li, .stApp input, .stApp textarea, .stApp button, .stApp [data-baseweb="select"] { font-family: var(--ru-fuente); font-size: var(--ru-texto); }\n.stApp p, .stApp label, .stApp li { color: var(--ru-tinta); }\n.stApp h1, .stApp h2, .stApp h3, .stApp h4 { color: var(--ru-tinta); font-family: var(--ru-fuente); line-height: 1.25; font-weight: 600; letter-spacing: -.015em; }\n.stApp h1 { font-size: var(--ru-h1); }\n.stApp h2 { font-size: var(--ru-h2); }\n.stApp h3 { font-size: var(--ru-h3); }\n.stApp [data-testid="stCaptionContainer"] p { font-size: var(--ru-pequeno); color: var(--ru-secundario); }\n.stApp .block-container {\n  width: calc(100% - 16px); max-width: 1600px; margin: 4rem auto var(--ru-espacio-2);\n  padding: var(--ru-espacio-6); background: var(--ru-blanco); border: 1px solid var(--ru-linea); border-radius: var(--ru-radio-hoja);\n}\n.stApp [data-testid="stHorizontalBlock"], .stApp [data-testid="stHorizontalBlock"] > div { min-width: 0; }\n/* Keep native column-gap arithmetic; 4 equal KPI columns retain their original shared edges. */\n.stApp section[data-testid="stSidebar"] { background: var(--ru-fondo); }\n.stApp section[data-testid="stSidebar"] h2 { font-size: var(--ru-h3); }\n.stApp section[data-testid="stSidebar"] input { background: var(--ru-blanco); color: var(--ru-tinta); }\n.stApp section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: var(--ru-espacio-4); }\n.ru-cabecera { display: flex; align-items: center; gap: var(--ru-espacio-6); padding-bottom: var(--ru-espacio-6); border-bottom: 1px solid var(--ru-linea); }\n.ru-logo { width: 142px; height: auto; flex-shrink: 0; }\n.ru-identidad { min-width: 0; flex: 1; padding-left: var(--ru-espacio-6); border-left: 1px solid var(--ru-linea); }\n.ru-cabecera h1 { padding: 0; margin: 0; font-weight: 700; color: var(--ru-vino); }\n.stApp .ru-submarca { font-size: var(--ru-h3); margin: var(--ru-espacio-1) 0; font-weight: 500; }\n.stApp .ru-geografia { font-size: var(--ru-pequeno); color: var(--ru-secundario); margin: 0; }\n.ru-sello { display: flex; align-items: center; gap: var(--ru-espacio-3); color: var(--ru-secundario); font-size: var(--ru-pequeno); }\n.ru-icono { width: 20px; height: 20px; flex-shrink: 0; vertical-align: middle; }\n.ru-sello .ru-icono { width: 30px; height: 30px; color: var(--ru-vino); }\n.ru-sidebar-titulo { display: flex; gap: var(--ru-espacio-2); align-items: center; padding: var(--ru-espacio-2) 0; }\n.ru-sidebar-titulo .ru-icono { color: var(--ru-vino); }\n.ru-sidebar-titulo h2 { margin: 0; padding: 0; }\n.ru-estado-datos { display: flex; flex-wrap: wrap; gap: var(--ru-espacio-2) var(--ru-espacio-6); font-size: var(--ru-pequeno); color: var(--ru-secundario); padding: var(--ru-espacio-3) 0; border-bottom: 1px solid var(--ru-linea); }\n.ru-estado-listo { color: var(--ru-exito); font-weight: 600; }\n.ru-estado-listo::before { content: \'\'; display: inline-block; width: 6px; height: 6px; border-radius: 50%; background: var(--ru-exito); margin-right: var(--ru-espacio-2); }\n.ru-titulo-vista { margin: var(--ru-espacio-2) 0; }\n.ru-titulo-vista h2 { margin: 0; padding: 0; }\n.stApp .ru-titulo-vista p { color: var(--ru-secundario); font-size: var(--ru-pequeno); margin: var(--ru-espacio-1) 0 0; }\n.stApp [data-testid="stRadio"] [role="radiogroup"] { gap: var(--ru-espacio-2); }\n.stApp [data-testid="stRadio"] label { min-height: var(--ru-tacto); align-items: center; padding: var(--ru-espacio-2) var(--ru-espacio-3); margin: 0; border-radius: var(--ru-radio-control); }\n.stApp [data-testid="stRadio"] label:has(input:checked) { background: var(--ru-suave); color: var(--ru-vino); }\n/* The anchor styles only the module navigation, leaving the period and numbering radios intact. */\n.stApp .element-container:has(.ru-nav-anchor) { display: none; }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container [role="radiogroup"] { display: flex; flex-wrap: nowrap; gap: 0; border-bottom: 1px solid var(--ru-linea); }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container label { flex: 1 1 0; min-width: 0; min-height: 54px; padding: var(--ru-espacio-3) var(--ru-espacio-4); justify-content: center; border-radius: 0; border-bottom: 3px solid transparent; }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container label > div:first-of-type { display: none; }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container label > div:last-of-type { flex: 1; min-width: 0; margin: 0 !important; }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container label p { font-weight: 600; color: var(--ru-secundario); text-align: center; }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container label:has(input:checked) { border-bottom-color: var(--ru-vino); background: var(--ru-vino-suave); }\n.stApp .element-container:has(.ru-nav-anchor) + .element-container label:has(input:checked) p { color: var(--ru-vino); }\n.stApp [data-baseweb="select"] > div, .stApp [data-baseweb="input"], .stApp [data-baseweb="base-input"], .stApp textarea {\n  background: var(--ru-blanco); border-color: var(--ru-control); border-radius: var(--ru-radio-control); min-height: var(--ru-tacto); color: var(--ru-tinta);\n}\n.stApp input, .stApp textarea { color: var(--ru-tinta); -webkit-text-fill-color: var(--ru-tinta); font-size: var(--ru-texto); }\n.stApp [data-testid="stWidgetLabel"] p { font-size: var(--ru-pequeno); font-weight: 600; }\n.stApp [data-testid="stDownloadButton"] button, .stApp [data-testid="stButton"] button {\n  background: var(--ru-blanco); border: 1px solid var(--ru-control); color: var(--ru-vino); border-radius: var(--ru-radio-control);\n  min-height: var(--ru-tacto); min-width: var(--ru-tacto); padding: var(--ru-espacio-2) var(--ru-espacio-4); max-width: 100%;\n}\n.stApp button p { color: inherit; }\n.stApp [data-testid="stDownloadButton"] button:hover, .stApp [data-testid="stButton"] button:hover { background: var(--ru-vino-suave); border-color: var(--ru-vino); }\n.stApp button:disabled { color: var(--ru-secundario); background: var(--ru-suave); cursor: not-allowed; }\n.tarjeta-metrica { padding: var(--ru-espacio-4) var(--ru-espacio-4) var(--ru-espacio-4) 0; border-bottom: 1px solid var(--ru-linea); min-height: 112px; }\n.titulo-metrica { color: var(--ru-secundario); font-size: var(--ru-pequeno); font-weight: 500; display: flex; gap: var(--ru-espacio-2); align-items: center; margin-bottom: var(--ru-espacio-2); }\n.titulo-metrica .ru-icono { width: 16px; height: 16px; color: var(--ru-secundario); }\n.valor-metrica { font-size: var(--ru-cifra); color: var(--ru-tinta); line-height: 1.15; font-weight: 600; font-variant-numeric: tabular-nums; overflow-wrap: anywhere; }\n.texto-neutral, .texto-maximo, .texto-minimo { color: var(--ru-secundario); font-size: var(--ru-pequeno); margin-top: var(--ru-espacio-2); line-height: 1.35; overflow-wrap: anywhere; }\n.contenedor-top { max-width: 100%; width: 100%; overflow-x: auto; border: 1px solid var(--ru-linea); border-radius: var(--ru-radio-panel); background: var(--ru-blanco); }\n.tabla-top { width: 100%; border-collapse: collapse; font-family: var(--ru-fuente); font-size: var(--ru-pequeno); font-variant-numeric: tabular-nums; }\n.tabla-top th { font-weight: 600; color: var(--ru-secundario); background: var(--ru-suave); text-align: left; }\n.tabla-top td, .tabla-top th { padding: 12px 10px; border-bottom: 1px solid var(--ru-linea); vertical-align: middle; line-height: 1.35; }\n.tabla-top td { color: var(--ru-tinta); }\n.tabla-top td:first-child { width: 40px; text-align: center; color: var(--ru-secundario); }\n.tabla-top tbody tr:nth-child(-n+3) td:first-child { color: var(--ru-vino); font-weight: 700; }\n.tabla-top td:nth-child(3), .tabla-top th:nth-child(3) { text-align: right; white-space: nowrap; }\n.tabla-top tr:last-child td { border-bottom: 0; }\n.tabla-top tbody tr:hover { background: var(--ru-suave); }\n.ru-estacion-detalle { display: grid; grid-template-columns: 1fr auto; gap: var(--ru-espacio-2); padding: var(--ru-espacio-4); background: var(--ru-suave); border-radius: var(--ru-radio-control); }\n.ru-estacion-detalle h3 { padding: 0; margin: 0; }\n.stApp .ru-estacion-detalle p { color: var(--ru-secundario); font-size: var(--ru-pequeno); margin: 0; }\n.ru-estacion-detalle strong { font-size: var(--ru-h2); font-variant-numeric: tabular-nums; }\n.ru-estacion-detalle > p { grid-column: 1 / -1; }\n.stApp [data-testid="stImage"] { min-width: 0; max-width: 100%; }\n.stApp [data-testid="stImage"] img { max-width: 100%; height: auto; border-radius: var(--ru-radio-control); }\n.stApp [data-testid="stVegaLiteChart"], .stApp [data-testid="stDataFrame"], .stApp [data-testid="stDeckGlJsonChart"] { min-width: 0; max-width: 100%; overflow-x: auto; }\n.stApp [data-testid="stExpander"] { border: 1px solid var(--ru-linea); border-radius: var(--ru-radio-control); background: var(--ru-blanco); }\n.stApp [data-testid="stExpander"] summary, .stApp [data-testid="stCheckbox"] label { min-height: var(--ru-tacto); }\n.stApp [data-baseweb="tab-list"] { gap: var(--ru-espacio-4); overflow-x: auto; border-bottom: 1px solid var(--ru-linea); scrollbar-width: thin; }\n.stApp button[data-baseweb="tab"] { color: var(--ru-secundario); min-height: 48px; padding: var(--ru-espacio-2) var(--ru-espacio-3); flex-shrink: 0; }\n.stApp button[data-baseweb="tab"][aria-selected="true"] { color: var(--ru-vino); font-weight: 600; }\n.stApp [data-baseweb="tab-highlight"] { background: var(--ru-vino); }\n.stApp [data-baseweb="tab-border"] { background: transparent; }\n.stApp [data-baseweb="tab-panel"] { padding-top: var(--ru-espacio-6); min-width: 0; }\n.stApp [data-testid="stAlert"] { border-radius: var(--ru-radio-control); }\n.ev-header { padding: var(--ru-espacio-3) 0; }\n.ev-header h2 { color: var(--ru-tinta); margin: 0; padding: 0; }\n.ev-header p { color: var(--ru-secundario); font-size: var(--ru-pequeno); margin: var(--ru-espacio-1) 0 0; }\n.stApp hr { border-color: var(--ru-linea); }\n[data-baseweb="popover"] > div, [data-baseweb="calendar"], [role="listbox"] { background: var(--ru-blanco); color: var(--ru-tinta); font-family: var(--ru-fuente); }\n[data-baseweb="calendar"] { border: 1px solid var(--ru-linea); border-radius: var(--ru-radio-panel); box-shadow: var(--ru-sombra-flotante); }\n[data-baseweb="calendar"] button { min-width: var(--ru-tacto); min-height: var(--ru-tacto); }\n/* DatePicker lives outside .stApp; override native dark text on its white surface. */\n[data-baseweb="calendar"], [data-baseweb="calendar"] * {\n  color: var(--ru-tinta) !important; -webkit-text-fill-color: var(--ru-tinta) !important;\n  font-family: var(--ru-fuente) !important;\n}\n[data-baseweb="calendar"], [data-baseweb="calendar"] > div,\n[data-baseweb="calendar"] [role="presentation"],\n[data-baseweb="calendar"] > div > div > div:first-child { background: var(--ru-blanco) !important; }\n[data-baseweb="calendar"] { width: min(326px, calc(100vw - 24px)); padding: 8px; box-sizing: border-box; }\n[data-baseweb="calendar"] > div, [data-baseweb="calendar"] > div > div,\n[data-baseweb="calendar"] [role="grid"], [data-baseweb="calendar"] [role="grid"] > div {\n  width: 100% !important; padding: 0 !important; margin: 0 !important; box-sizing: border-box;\n}\n[data-baseweb="calendar"] [role="gridcell"], [data-baseweb="calendar"] [role="gridcell"] * { font-size: var(--ru-pequeno) !important; }\n[data-baseweb="calendar"] [role="row"], [data-baseweb="calendar"] [role="presentation"] {\n  display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); width: 100%;\n}\n[data-baseweb="calendar"] [role="gridcell"] {\n  width: 100%; min-width: 0; height: 44px; display: flex; align-items: center; justify-content: center;\n  border-radius: 6px; background: transparent !important;\n}\n[data-baseweb="calendar"] [role="gridcell"]::before,\n[data-baseweb="calendar"] [role="gridcell"]::after { background: transparent !important; border-color: transparent !important; }\n[data-baseweb="calendar"] [role="gridcell"] > div { background: transparent !important; }\n[data-baseweb="calendar"] [role="presentation"] > div { width: 100%; color: var(--ru-secundario) !important; }\n[data-baseweb="calendar"] [role="gridcell"][aria-label^="Not available"],\n[data-baseweb="calendar"] [role="gridcell"][aria-disabled="true"] {\n  color: var(--ru-secundario) !important; -webkit-text-fill-color: var(--ru-secundario) !important; background: var(--ru-suave) !important;\n}\n[data-baseweb="calendar"] [role="gridcell"][aria-label^="Not available"] *,\n[data-baseweb="calendar"] [role="gridcell"][aria-disabled="true"] * {\n  color: var(--ru-secundario) !important; -webkit-text-fill-color: var(--ru-secundario) !important;\n}\n[data-baseweb="calendar"] [role="gridcell"]:hover { outline: 1px solid var(--ru-vino); outline-offset: -2px; }\n[data-baseweb="calendar"] [aria-selected="true"],\n[data-baseweb="calendar"] [role="gridcell"][aria-label^="Selected"] { background: var(--ru-vino) !important; }\n[data-baseweb="calendar"] [aria-selected="true"], [data-baseweb="calendar"] [aria-selected="true"] *,\n[data-baseweb="calendar"] [role="gridcell"][aria-label^="Selected"],\n[data-baseweb="calendar"] [role="gridcell"][aria-label^="Selected"] * {\n  color: var(--ru-blanco) !important; -webkit-text-fill-color: var(--ru-blanco) !important;\n}\n[data-baseweb="calendar"] svg { color: var(--ru-tinta) !important; fill: currentColor; }\n[data-baseweb="calendar"] :focus-visible { outline: 3px solid var(--ru-vino); outline-offset: 1px; }\n[role="listbox"], [role="option"], [role="option"] *, [data-baseweb="tooltip"] {\n  color: var(--ru-tinta) !important; -webkit-text-fill-color: var(--ru-tinta) !important; background-color: var(--ru-blanco) !important;\n}\n[role="option"][aria-selected="true"] { background-color: var(--ru-vino-suave) !important; }\n.stApp [data-baseweb="select"] svg, .stApp [data-baseweb="input"] svg { color: var(--ru-secundario); fill: currentColor; }\n.stApp [data-testid="stSpinner"] { color: var(--ru-tinta); }\n.stApp button:focus-visible, .stApp input:focus-visible, .stApp summary:focus-visible, .stApp a:focus-visible,\n.stApp label[data-baseweb="radio"]:focus-within, .contenedor-top:focus-visible { outline: 3px solid var(--ru-vino); outline-offset: 3px; }\n[data-testid="collapsedControl"] button, [data-testid="stSidebar"] button { min-height: var(--ru-tacto); min-width: var(--ru-tacto); }\n@media (max-width: 1100px) {\n  .ru-sello { display: none; }\n  .ru-logo { width: 112px; }\n  .stApp .block-container { padding: var(--ru-espacio-4); }\n}\n@media (max-width: 850px) {\n  .stApp [data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }\n  .stApp [data-testid="stHorizontalBlock"] > [data-testid="column"],\n  .stApp [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { width: 100% !important; flex: 1 1 100% !important; }\n  .stApp [data-testid="stHorizontalBlock"]:has(.tarjeta-metrica) > div { width: calc(50% - .5rem) !important; flex: 1 1 calc(50% - .5rem) !important; }\n  .stApp [data-testid="stHorizontalBlock"]:has(.tarjeta-metrica) > div:last-child:nth-child(odd) { flex-basis: 100% !important; }\n  .stApp [data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]) > div { width: calc(50% - .5rem) !important; flex: 1 1 calc(50% - .5rem) !important; }\n  .tarjeta-metrica { min-height: 104px; }\n  .tabla-top td, .tabla-top th { padding: 10px 8px; }\n}\n@media (max-width: 600px) {\n  .stApp .block-container { width: 100%; margin-top: 3.5rem; padding: var(--ru-espacio-4); border: 0; border-radius: 0; }\n  .ru-cabecera { gap: var(--ru-espacio-3); align-items: flex-start; padding-bottom: var(--ru-espacio-4); }\n  .ru-logo { width: 74px; }\n  .ru-identidad { padding-left: var(--ru-espacio-3); }\n  .stApp .ru-submarca { font-size: var(--ru-pequeno); }\n  .stApp .ru-geografia { font-size: .8125rem; }\n  .ru-estado-datos { gap: var(--ru-espacio-1) var(--ru-espacio-3); font-size: .8125rem; }\n  .ru-estado-datos > span { width: 100%; }\n  .ru-estado-datos > span:last-child { font-size: .8125rem; }\n  .stApp .element-container:has(.ru-nav-anchor) + .element-container label { padding: var(--ru-espacio-2) 4px; min-height: 58px; }\n  .stApp .element-container:has(.ru-nav-anchor) + .element-container label p { font-size: .8125rem; text-align: center; overflow-wrap: normal !important; word-break: keep-all !important; line-height: 1.3; }\n  .stApp [data-testid="stRadio"] label { padding: var(--ru-espacio-2); }\n  .stApp [data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]) > div:first-child,\n  .stApp [data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]) > div:last-child { width: 100% !important; flex: 1 1 100% !important; }\n  .ru-estacion-detalle { grid-template-columns: 1fr; }\n  .titulo-metrica .ru-icono { display: none; }\n}\n@media (prefers-reduced-motion: reduce) {\n  *, *::before, *::after { animation: none !important; transition: none !important; scroll-behavior: auto !important; }\n}\n\n/* Compact contextual exports. Keep 44px touch targets without full-width cards. */\n.stApp div[data-testid="stDownloadButton"] button,\n.stApp div.stDownloadButton button {\n    width: auto !important; max-width: 100% !important;\n    min-height: var(--ru-tacto, 44px) !important;\n    padding: 8px 12px !important;\n    border: 1px solid var(--ru-linea, #e0e6ed) !important;\n    border-radius: var(--ru-radio-control, 6px);\n    background: var(--ru-blanco, #ffffff); text-align: left;\n}\n.stApp div[data-testid="stDownloadButton"] button p,\n.stApp div.stDownloadButton button p {\n    color: var(--ru-vino, #84273f) !important;\n    -webkit-text-fill-color: var(--ru-vino, #84273f) !important;\n    font: 600 var(--ru-pequeno, .875rem)/1.3 var(--ru-fuente, sans-serif) !important;\n    white-space: normal; overflow-wrap: normal;\n}\n.stApp div[data-testid="stDownloadButton"] button:disabled p {\n    color: var(--ru-secundario, #526476) !important;\n    -webkit-text-fill-color: var(--ru-secundario, #526476) !important;\n}\n.stApp .ru-export-contexto { color: var(--ru-secundario); font-size: var(--ru-pequeno); margin: 0 0 8px; }\n.stApp [data-testid="stExpander"]:has(.ru-export-contexto) [data-testid="stVerticalBlock"] { gap: 8px; }\n@media (max-width: 850px) {\n    .stApp [data-testid="stExpander"]:has(.ru-export-contexto) [data-testid="stHorizontalBlock"] { gap: 8px; }\n    .stApp [data-testid="stExpander"]:has(.ru-export-contexto) [data-testid="stHorizontalBlock"] > [data-testid="column"],\n    .stApp [data-testid="stExpander"]:has(.ru-export-contexto) [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {\n        width: calc(50% - 4px) !important; flex: 1 1 calc(50% - 4px) !important;\n    }\n}\n\n</style>', unsafe_allow_html=True)

import base64
import json

ICONOS_RUEMAG = {
    'red': '<circle cx="6" cy="6" r="2"/><circle cx="18" cy="6" r="2"/><circle cx="12" cy="18" r="2"/><path d="m8 7 3 9m5-9-3 9M8 6h8"/>',
    'candado': '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3m-4 5v2"/>',
    'promedio': '<path d="M4 12h16M6 7h12M6 17h12"/>',
    'maximo': '<path d="m6 10 6-6 6 6M12 4v16"/>',
    'minimo': '<path d="m6 14 6 6 6-6M12 4v16"/>',
    'estacion': '<path d="M12 21s7-6 7-12a7 7 0 0 0-14 0c0 6 7 12 7 12Z"/><circle cx="12" cy="9" r="2"/>',
}


def icono_ruemag(nombre):
    contenido = ICONOS_RUEMAG.get(nombre, ICONOS_RUEMAG['red'])
    return '<svg class="ru-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + contenido + '</svg>'


def recordar_cambio_ruemag(clave, callback, argumentos, parametros):
    st.session_state.setdefault('_ruemag_filtros', {})[clave] = st.session_state[clave]
    if callback is not None:
        callback(*(argumentos or ()), **(parametros or {}))


def control_ruemag(tipo, label, *args, **opciones):
    """Memoria de filtros por módulo, independiente de la limpieza de widgets."""
    memoria = st.session_state.setdefault('_ruemag_filtros', {})
    iniciales = st.session_state.setdefault('_ruemag_iniciales', {})
    clave = opciones.get('key')
    argumentos = list(args)
    if clave:
        generacion = st.session_state.get('_ruemag_generacion', 0)
        anterior = iniciales.get(clave, {})
        restaurar = not anterior or (clave != 'ruemag_modulo' and anterior.get('generacion') != generacion)
        if tipo in (st.selectbox, st.radio):
            alternativas = list(argumentos[0] if argumentos else opciones['options'])
            indice = argumentos[1] if len(argumentos) > 1 else opciones.get('index', 0)
            defecto = alternativas[indice] if alternativas and indice is not None else None
            guardado = memoria.get(clave, defecto) if restaurar else anterior.get('valor', defecto)
            inicial = guardado if guardado in alternativas else defecto
            indice_inicial = alternativas.index(inicial) if inicial in alternativas else None
            if len(argumentos) > 1: argumentos[1] = indice_inicial
            else: opciones['index'] = indice_inicial
        elif tipo == st.date_input:
            defecto = argumentos[0] if argumentos else opciones.get('value')
            inicial = memoria.get(clave, defecto) if restaurar else anterior.get('valor', defecto)
            minimo, maximo = opciones.get('min_value'), opciones.get('max_value')
            if isinstance(inicial, datetime.date):
                if minimo is not None: inicial = max(inicial, minimo)
                if maximo is not None: inicial = min(inicial, maximo)
            if argumentos: argumentos[0] = inicial
            else: opciones['value'] = inicial
        elif tipo == st.checkbox:
            defecto = argumentos[0] if argumentos else opciones.get('value', False)
            inicial = bool(memoria.get(clave, defecto) if restaurar else anterior.get('valor', defecto))
            if argumentos: argumentos[0] = inicial
            else: opciones['value'] = inicial
        iniciales[clave] = {'valor': inicial, 'generacion': generacion}
        callback = opciones.pop('on_change', None)
        argumentos_callback = opciones.pop('args', None)
        parametros_callback = opciones.pop('kwargs', None)
        opciones['on_change'] = recordar_cambio_ruemag
        opciones['args'] = (clave, callback, argumentos_callback, parametros_callback)
    valor = tipo(label, *argumentos, **opciones)
    if tipo == st.date_input and valor is None:
        st.warning('Selecciona una fecha válida para continuar con la consulta.')
        st.stop()
    if clave:
        memoria[clave] = valor
        if clave == 'ruemag_modulo' and st.session_state.get('_ruemag_modulo_visto') != valor:
            st.session_state['_ruemag_generacion'] = st.session_state.get('_ruemag_generacion', 0) + 1
            st.session_state['_ruemag_modulo_visto'] = valor
    return valor


def marcar_periodo_personalizado(clave):
    st.session_state[clave] = 'Personalizado'
    st.session_state.setdefault('_ruemag_filtros', {})[clave] = 'Personalizado'


def aplicar_periodo_ruemag(clave_periodo, clave_inicio, clave_fin, minimo, maximo):
    seleccion = st.session_state[clave_periodo]
    st.session_state.setdefault('_ruemag_filtros', {})[clave_periodo] = seleccion
    if seleccion == 'Personalizado':
        return
    dias = {'Último día': 1, '7 días': 7, '30 días': 30}[seleccion]
    inicio = max(minimo, maximo - datetime.timedelta(days=dias - 1))
    memoria = st.session_state.setdefault('_ruemag_filtros', {})
    for clave, valor in [(clave_inicio, inicio), (clave_fin, maximo)]:
        st.session_state[clave] = valor
        memoria[clave] = valor


def mostrar_periodos_ruemag(clave, clave_inicio, clave_fin, minimo, maximo):
    control_ruemag(
        st.selectbox, 'Periodo rápido', ['Personalizado', 'Último día', '7 días', '30 días'],
        key=clave,
        help=f'Los accesos rápidos terminan en el último día con datos: {maximo:%d/%m/%Y}. Los calendarios permanecen disponibles.',
        on_change=aplicar_periodo_ruemag,
        args=(clave, clave_inicio, clave_fin, minimo, maximo)
    )


def encabezado_ruemag():
    logo = ''
    if Path(RUTA_LOGO).exists():
        contenido = base64.b64encode(Path(RUTA_LOGO).read_bytes()).decode('ascii')
        logo = f'<img class="ru-logo" src="data:image/png;base64,{contenido}" alt="Oficialía de Inteligencia">'
    else:
        st.warning('No se encontró el logo.')
    st.markdown(
        '<header class="ru-cabecera">' + logo + '<div class="ru-identidad">'
        '<h1>MAPAS RUEMAG</h1><p class="ru-submarca">Red Urbana de Estaciones Meteorológicas</p>'
        '<p class="ru-geografia">Guadalajara · Monitoreo, análisis espacial y eventos de tormenta</p>'
        '</div><div class="ru-sello">' + icono_ruemag('red') + '<span>Observación<br>meteorológica urbana</span></div></header>',
        unsafe_allow_html=True
    )


def estado_datos_ruemag(datos_base):
    primero, ultimo = datos_base['Day'].min(), datos_base['Day'].max()
    consulta = pd.Timestamp.now(tz='America/Mexico_City')
    st.markdown(
        '<div class="ru-estado-datos" role="status">'
        '<span class="ru-estado-listo">Datos cargados</span>'
        f'<span>Último dato disponible <strong>{ultimo:%d/%m/%Y}</strong></span>'
        f'<span>Histórico desde {primero:%d/%m/%Y}</span>'
        f'<span>Consulta {consulta:%d/%m/%Y · %H:%M} · UTC−6</span></div>', unsafe_allow_html=True
    )


def titulo_vista_ruemag(titulo, descripcion):
    st.markdown('<div class="ru-titulo-vista"><h2>' + escape(titulo) + '</h2><p>' + escape(descripcion) + '</p></div>', unsafe_allow_html=True)


def mostrar_estacion_ruemag(puntos, catalogo, variable, inicio, fin, clave):
    """Detalle y ubicación reales; la cartografía IDW original permanece visible."""
    catalogo = catalogo.sort_values('nombre_estacion').drop_duplicates('estacion')
    nombres = dict(zip(catalogo.estacion, catalogo.nombre_estacion))
    elegida = control_ruemag(st.selectbox, 'Consultar una estación', [''] + catalogo.estacion.tolist(),
        format_func=lambda codigo: nombres.get(codigo, 'Selecciona una estación'), key=clave,
        help='Consulta el valor del periodo y su ubicación real. S/D indica que no hay mediciones disponibles.')
    if not elegida:
        return
    ubicacion = catalogo.loc[catalogo.estacion == elegida].iloc[0]
    observada = puntos.loc[puntos.estacion == elegida]
    valor = 'S/D' if observada.empty else f"{observada.iloc[0][variable]:.1f} {VARIABLES[variable]['unidad']}"
    direccion = ''
    if not observada.empty and variable in ['Wind_Avg', 'Wind_Max']:
        direccion = f" · Dirección: {observada.iloc[0].get('direccion_cardinal', 'S/D')}"
    st.markdown(
        '<section class="ru-estacion-detalle"><div><h3>' + escape(str(ubicacion.nombre_estacion)) + '</h3>'
        '<p>' + escape(VARIABLES[variable]['nombre']) + '</p></div><strong>' + escape(valor) + '</strong>'
        f'<p class="ru-estacion-periodo">{inicio:%d/%m/%Y} al {fin:%d/%m/%Y}' + escape(direccion) + '</p>'
        f'<p>Latitud {ubicacion.lat:.5f} · Longitud {ubicacion.lon:.5f}</p></section>', unsafe_allow_html=True
    )
    if observada.empty:
        st.info('Esta estación no tiene datos de la variable en el periodo elegido. Su ubicación se conserva.')
    st.map(pd.DataFrame({'lat': [float(ubicacion.lat)], 'lon': [float(ubicacion.lon)]}), zoom=13)
    st.caption('Ubicación de la estación seleccionada. El mapa IDW mantiene todas las estaciones y sus escalas originales.')


def contexto_cartografico_ruemag():
    """Invalida la caché ante cambios de código, identidad o escalas científicas."""
    archivos = [Path(__file__), Path(RUTA_LOGO), Path(RUTA_ROSA)]
    firmas = tuple((str(ruta), ruta.stat().st_mtime_ns if ruta.exists() else None) for ruta in archivos)
    return firmas, json.dumps(VARIABLES, sort_keys=True, ensure_ascii=False)


@st.cache_data(show_spinner=False, max_entries=8, ttl=900, hash_funcs={
    'geopandas.geodataframe.GeoDataFrame': lambda tabla: (str(tabla.crs), tabla.geometry.to_wkb(hex=True).tolist())
})
def mapa_ruemag_cache(contexto, *args, **kwargs):
    return crear_mapa(*args, **kwargs)


@st.cache_data(show_spinner=False, max_entries=8, ttl=900, hash_funcs={
    'geopandas.geodataframe.GeoDataFrame': lambda tabla: (str(tabla.crs), tabla.geometry.to_wkb(hex=True).tolist())
})
def kmz_ruemag_cache(contexto, *args, **kwargs):
    return crear_kmz(*args, **kwargs)


@st.cache_data(show_spinner=False, max_entries=8, ttl=900, hash_funcs={
    'geopandas.geodataframe.GeoDataFrame': lambda tabla: (str(tabla.crs), tabla.geometry.to_wkb(hex=True).tolist())
})
def shp_ruemag_cache(contexto, *args, **kwargs):
    return crear_shapefile_idw(*args, **kwargs)


def mostrar_tarjeta(titulo, valor, texto_inferior="", tipo="neutral"):
    clase = {"maximo": "texto-maximo", "minimo": "texto-minimo"}.get(tipo, "texto-neutral")
    flecha = {"maximo": "↑ ", "minimo": "↓ "}.get(tipo, "")
    simbolo = "estacion" if "estaciones" in titulo.lower() else (tipo if tipo in ["maximo", "minimo"] else "promedio")
    st.markdown(
        '<div class="tarjeta-metrica">'
        '<div class="titulo-metrica">' + icono_ruemag(simbolo) + escape(str(titulo)) + '</div>'
        '<div class="valor-metrica">' + escape(str(valor)) + '</div>'
        f'<div class="{clase}">' + escape(flecha + str(texto_inferior)) + '</div></div>',
        unsafe_allow_html=True
    )


# ============================================================
# LEER EXCEL
# ============================================================


def descarga_ruemag(label, data, file_name, mime, *, titulo, contexto, formato, **opciones):
    """Un control compacto; descripción completa en la ayuda, bytes originales."""
    contenido = data.getvalue() if hasattr(data, 'getvalue') else data
    cantidad = len(contenido) if contenido is not None else 0
    peso = f'{cantidad / 1048576:.1f} MB' if cantidad >= 1048576 else f'{cantidad / 1024:.0f} KB'
    detalle = f'{titulo} · {contexto} · {formato} · {peso}\n\nArchivo: {file_name}'
    ayuda = opciones.pop('help', None)
    opciones['help'] = detalle + (f'\n\n{ayuda}' if ayuda else '')
    opciones['use_container_width'] = False
    texto = str(titulo).lower()
    if formato == 'PNG':
        if 'mapa' in texto:
            nombre = 'Mapa de anomalías' if 'anomal' in texto else 'Mapa'
        elif 'evolución' in texto:
            nombre = 'Gráfica por estación'
        elif 'categorías' in texto:
            nombre = 'Categorías de lluvia'
        elif 'duración' in texto:
            nombre = 'Duración de eventos'
        elif 'máxima por evento' in texto:
            nombre = 'Lluvia por evento'
        else:
            nombre = 'Gráfica por estación'
    else:
        nombre = {'CSV': 'Datos por estación', 'SHP': 'Capa IDW', 'KMZ': 'Google Earth', 'TXT': 'Reporte del evento'}.get(formato, titulo)
    return st.download_button(label=f'↓ {nombre} · {formato}', data=data, file_name=file_name, mime=mime, **opciones)


@st.cache_data(show_spinner=False)
def leer_excel(origen):

    if isinstance(origen, bytes):
        libro = pd.ExcelFile(io.BytesIO(origen))
    else:
        libro = pd.ExcelFile(origen)

    if "UBICACION" not in libro.sheet_names:
        raise ValueError("El Excel no contiene la hoja UBICACION.")

    ubicaciones = pd.read_excel(libro, sheet_name="UBICACION")

    ubicaciones.columns = [
        str(columna).strip().lower()
        for columna in ubicaciones.columns
    ]

    columnas_necesarias = {"estacion", "lon", "lat"}

    if not columnas_necesarias.issubset(ubicaciones.columns):
        raise ValueError("La hoja UBICACION debe contener estacion, lon y lat.")

    ubicaciones["estacion"] = (
        ubicaciones["estacion"].astype(str).str.strip().str.lower()
    )

    coordenadas = ubicaciones.set_index("estacion")[["lon", "lat"]]

    lista_datos = []

    for hoja in libro.sheet_names:

        estacion = hoja.strip().lower()

        if estacion == "ubicacion":
            continue
        if estacion not in coordenadas.index:
            continue

        datos = pd.read_excel(libro, sheet_name=hoja)
        datos.columns = [str(columna).strip() for columna in datos.columns]

        if "Day" not in datos.columns:
            continue

        datos["Day"] = pd.to_datetime(datos["Day"], errors="coerce")
        datos = datos.dropna(subset=["Day"])
        datos["estacion"] = estacion
        datos["nombre_estacion"] = NOMBRES_ESTACIONES.get(estacion, hoja)
        datos["lon"] = float(coordenadas.loc[estacion, "lon"])
        datos["lat"] = float(coordenadas.loc[estacion, "lat"])

        for variable in COLUMNAS_METEOROLOGICAS:
            if variable in datos.columns:
                datos[variable] = pd.to_numeric(datos[variable], errors="coerce")

        lista_datos.append(datos)

    if not lista_datos:
        raise ValueError("No se encontraron hojas de estaciones.")

    resultado = pd.concat(lista_datos, ignore_index=True)
    resultado.attrs["ubicaciones"] = ubicaciones[["estacion", "lon", "lat"]].to_dict("records")
    return resultado


# ============================================================
# LEER EXCEL DE EVENTOS (FORMATO ANCHO)
# ============================================================

def normalizar_texto_evento(texto):

    texto = unicodedata.normalize("NFKD", str(texto).strip().lower())
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return " ".join(texto.replace("_", " ").replace("-", " ").split())


def leer_historico_temporales(origen):
    """Lee hojas anuales; no requiere ID ni coordenadas de estaciones."""
    temporales = {}
    with pd.ExcelFile(origen) as libro:
        for hoja in libro.sheet_names:
            anio = str(hoja).strip()
            if len(anio) != 4 or not anio.isdigit():
                continue
            tabla = pd.read_excel(libro, sheet_name=hoja)
            tabla.columns = [normalizar_texto_evento(c) for c in tabla.columns]
            if not {'evento', 'fecha'}.issubset(tabla.columns):
                continue
            tabla = tabla.rename(columns={'evento': 'Evento', 'fecha': 'Fecha'})
            tabla['Evento'] = pd.to_numeric(tabla['Evento'], errors='coerce')
            tabla['Fecha'] = pd.to_datetime(tabla['Fecha'], errors='coerce', dayfirst=True)
            tabla = tabla.dropna(subset=['Evento', 'Fecha']).copy()
            if tabla.empty:
                continue
            tabla['Evento'] = tabla['Evento'].astype(int)
            tabla['Anio'] = int(anio)
            for columna in ['precip', 'viento']:
                tabla[columna] = pd.to_numeric(tabla.get(columna, float('nan')), errors='coerce')
            for columna in ['zona precip', 'zona viento']:
                if columna not in tabla.columns:
                    tabla[columna] = 'Sin dato'
                tabla[columna] = tabla[columna].fillna('Sin dato')
            temporales[anio] = tabla.drop_duplicates(subset=['Fecha', 'Evento']).sort_values(['Fecha', 'Evento'])
    if not temporales:
        raise ValueError('No hay hojas anuales con eventos y fechas válidos.')
    consecutivo = 1
    for anio in sorted(temporales, key=int):
        tabla = temporales[anio]
        tabla['Consecutivo'] = range(consecutivo, consecutivo + len(tabla))
        consecutivo += len(tabla)
    return temporales


@st.cache_data(show_spinner=False, ttl=5)
def localizar_archivos_eventos(ruta_carpeta):

    carpeta = Path(ruta_carpeta)

    if not carpeta.exists():
        raise FileNotFoundError(f"No se encontró la carpeta de eventos: {carpeta}")

    candidatos = []

    for archivo in sorted(carpeta.glob("*.xlsx")):
        try:
            libro = pd.ExcelFile(archivo)
        except Exception:
            continue

        texto_identificacion = normalizar_texto_evento(f"{archivo.stem} {' '.join(libro.sheet_names)}")
        tipo = None

        if any(palabra in texto_identificacion for palabra in ["racha", "viento", "gust"]):
            tipo = "rachas"
        elif any(palabra in texto_identificacion for palabra in ["lluvia", "precipitacion", "rain"]):
            tipo = "lluvia"

        if tipo is None:
            continue

        for hoja in libro.sheet_names:
            encabezados = pd.read_excel(libro, sheet_name=hoja, nrows=0).columns
            encabezados_norm = {normalizar_texto_evento(columna) for columna in encabezados}

            if {"evento", "fecha"}.issubset(encabezados_norm):
                muestra_eventos = pd.read_excel(
                    libro, sheet_name=hoja,
                    usecols=lambda columna: normalizar_texto_evento(columna) == "evento"
                )

                mayor_evento = pd.to_numeric(muestra_eventos.iloc[:, 0], errors="coerce").max()
                if pd.isna(mayor_evento):
                    mayor_evento = -1

                candidatos.append((tipo, float(mayor_evento), archivo, hoja))
                break

    encontrados = {}
    candidatos.sort(key=lambda elemento: elemento[1], reverse=True)

    for tipo, mayor_evento, archivo, hoja in candidatos:
        if tipo not in encontrados:
            encontrados[tipo] = (archivo, hoja)

    faltantes = [tipo for tipo in ["lluvia", "rachas"] if tipo not in encontrados]

    if faltantes:
        raise FileNotFoundError(
            "No se identificó el Excel de " + " y ".join(faltantes) + 
            ". El nombre del archivo o de su hoja debe incluir 'lluvia/precipitación' o 'rachas/viento'."
        )

    return encontrados


@st.cache_data(show_spinner=False, ttl=5)
def leer_tabla_eventos(ruta_archivo, hoja):

    tabla = pd.read_excel(ruta_archivo, sheet_name=hoja)
    tabla = tabla.loc[:, ~tabla.columns.astype(str).str.startswith("Unnamed:")].copy()
    equivalencias = {normalizar_texto_evento(columna): columna for columna in tabla.columns}

    columna_evento = equivalencias.get("evento")
    columna_fecha = equivalencias.get("fecha")

    if columna_evento is None or columna_fecha is None:
        raise ValueError("El Excel de eventos debe contener las columnas Evento y Fecha.")

    tabla = tabla.rename(columns={columna_evento: "Evento", columna_fecha: "Fecha"})
    tabla["Evento"] = pd.to_numeric(tabla["Evento"], errors="coerce")
    tabla["Fecha"] = pd.to_datetime(tabla["Fecha"], dayfirst=True, errors="coerce")

    return tabla.dropna(subset=["Evento"]).copy()


def convertir_evento_a_puntos(tabla, numero_evento, datos_base, variable):

    filas = tabla[tabla["Evento"] == numero_evento]

    if filas.empty:
        return pd.DataFrame(), pd.NaT

    fila = filas.iloc[0]

    estaciones_base = datos_base[[
        "estacion", "nombre_estacion", "lon", "lat"
    ]].drop_duplicates(subset=["estacion"]).copy()

    alias_estaciones = {
        "expo": "archivo", "expo gdl": "archivo", "agua azul": "base1", "transito": "base2",
        "zona industrial": "base3", "san andres": "base4", "periferico norte": "base5",
        "providencia acueducto": "colegio", "oblatos": "colmena", "crit": "crit",
        "dean": "dean", "estadio": "estadio", "centro": "mercado", "mercado": "mercado"
    }

    valores = {}

    for columna in tabla.columns:
        nombre_normalizado = normalizar_texto_evento(columna)
        estacion = alias_estaciones.get(nombre_normalizado)

        if estacion is not None:
            valores[estacion] = pd.to_numeric(fila[columna], errors="coerce")

    estaciones_base[variable] = estaciones_base["estacion"].map(valores)
    if variable != "Rain_Total":
        estaciones_base.loc[estaciones_base['estacion'].astype(str).str.strip().str.lower().eq('dean'), variable] = np.nan
    puntos = estaciones_base.dropna(subset=[variable]).copy()

    return puntos, fila["Fecha"]


@st.cache_data(show_spinner=False)
def leer_climatologia(origen_csv):

    if isinstance(origen_csv, bytes):
        origen_csv = io.BytesIO(origen_csv)

    clima = pd.read_csv(origen_csv, encoding="utf-8-sig")
    columnas_requeridas = {"FECHA", "PRECIP", "TMAX", "TMIN"}

    if not columnas_requeridas.issubset(clima.columns):
        raise ValueError("El CSV climatológico debe contener las columnas: FECHA, PRECIP, TMAX y TMIN.")

    clima["FECHA"] = pd.to_datetime(clima["FECHA"], dayfirst=True, errors="coerce")

    for columna in ["PRECIP", "TMAX", "TMIN"]:
        clima[columna] = pd.to_numeric(clima[columna], errors="coerce")

    return clima.dropna(subset=["FECHA"])


@st.cache_data(show_spinner=False)
def leer_rachas_historicas(origen_excel):

    if isinstance(origen_excel, bytes):
        origen_excel = io.BytesIO(origen_excel)

    rachas = pd.read_excel(origen_excel, sheet_name="seriecompleta")

    if not {"fecha", "km/h"}.issubset(rachas.columns):
        raise ValueError("El Excel histórico debe contener la hoja seriecompleta y las columnas fecha y km/h.")

    rachas["fecha"] = pd.to_datetime(rachas["fecha"], errors="coerce")
    rachas["km/h"] = pd.to_numeric(rachas["km/h"], errors="coerce")

    return rachas.dropna(subset=["fecha", "km/h"])


def calcular_referencia_anomalia(variable, fecha_inicial, fecha_final, clima, rachas):

    fechas_elegidas = pd.date_range(fecha_inicial, fecha_final, freq="D")
    mes_dia = set(fechas_elegidas.strftime("%m-%d"))

    if variable == "Gust_Max":
        historico = rachas.copy()
        columna_fecha = "fecha"
        columna_valor = "km/h"
        periodo_base = f"{historico[columna_fecha].min():%d/%m/%Y} al {historico[columna_fecha].max():%d/%m/%Y} (datos disponibles)"
    else:
        equivalencias = {"T_Max": "TMAX", "T_Min": "TMIN", "Rain_Total": "PRECIP"}
        historico = clima[clima["FECHA"].dt.year.between(1991, 2020)].copy()
        columna_fecha = "FECHA"
        columna_valor = equivalencias[variable]
        periodo_base = "1991–2020"

    historico = historico[historico[columna_fecha].dt.strftime("%m-%d").isin(mes_dia)].dropna(subset=[columna_valor])
    historico["anio"] = historico[columna_fecha].dt.year

    if variable == "Rain_Total":
        valores_anuales = historico.groupby("anio")[columna_valor].sum(min_count=1)
    elif variable in ["T_Max", "Gust_Max"]:
        valores_anuales = historico.groupby("anio")[columna_valor].max()
    else:
        valores_anuales = historico.groupby("anio")[columna_valor].min()

    if valores_anuales.empty:
        raise ValueError("No hay información histórica para las fechas elegidas.")

    return float(valores_anuales.mean()), periodo_base, int(valores_anuales.count())


def crear_csv_anomalias(datos, puntos, variable, clave_anomalia, referencia, fecha_inicial, fecha_final, periodo_base):

    estaciones = datos[["estacion", "nombre_estacion", "lon", "lat"]].drop_duplicates(subset=["estacion"])
    columnas = ["estacion", variable, clave_anomalia]

    tabla = estaciones.merge(puntos[columnas], on="estacion", how="left").sort_values("nombre_estacion")

    tabla.insert(0, "Fecha inicial", str(fecha_inicial))
    tabla.insert(1, "Fecha final", str(fecha_final))
    tabla["Referencia histórica"] = referencia
    tabla["Periodo base"] = periodo_base

    tabla = tabla.rename(columns={
        "estacion": "Clave de estación", "nombre_estacion": "Nombre de estación",
        "lon": "Longitud", "lat": "Latitud",
        variable: f"Valor observado ({VARIABLES[variable]['unidad']})",
        clave_anomalia: f"Anomalía ({VARIABLES[variable]['unidad']})"
    }).fillna("S/D")

    return tabla.to_csv(index=False, sep=",").encode("utf-8-sig")


# ============================================================
# DIRECCIÓN CARDINAL
# ============================================================

def direccion_cardinal(grados):
    if pd.isna(grados): return "S/D"

    direcciones = [
        "N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
        "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"
    ]
    indice = int((float(grados) + 11.25) // 22.5) % 16
    return direcciones[indice]


# ============================================================
# PROMEDIO CIRCULAR
# ============================================================

def promedio_circular(valores):
    valores = pd.to_numeric(valores, errors="coerce").dropna()
    if len(valores) == 0: return np.nan

    angulos = np.deg2rad(valores.to_numpy())
    seno = np.sin(angulos).mean()
    coseno = np.cos(angulos).mean()

    direccion = np.rad2deg(np.arctan2(seno, coseno)) % 360
    return direccion


# ============================================================
# RESUMIR EL PERIODO
# ============================================================

def resumir_periodo(datos, variable, fecha_inicial, fecha_final):

    seleccion = datos[
        (datos["Day"].dt.date >= fecha_inicial) & (datos["Day"].dt.date <= fecha_final)
    ].copy()

    grupos = ["estacion", "nombre_estacion", "lon", "lat"]

    if variable == "Rain_Total":
        resultado = seleccion.groupby(grupos, as_index=False)[variable].sum(min_count=1)
    elif variable in ["T_Max", "Wind_Max", "Gust_Max"]:
        resultado = seleccion.groupby(grupos, as_index=False)[variable].max()
    elif variable == "T_Min":
        resultado = seleccion.groupby(grupos, as_index=False)[variable].min()
    else:
        resultado = seleccion.groupby(grupos, as_index=False)[variable].mean()

    resultado = resultado.dropna(subset=[variable])

    if variable in ["Wind_Avg", "Wind_Max"]:
        if "Wind_Dir" in seleccion.columns:
            direcciones = seleccion.groupby(grupos, as_index=False)["Wind_Dir"].agg(promedio_circular)
            resultado = resultado.merge(direcciones, on=grupos, how="left")
            resultado["direccion_cardinal"] = resultado["Wind_Dir"].apply(direccion_cardinal)
        else:
            resultado["Wind_Dir"] = np.nan
            resultado["direccion_cardinal"] = "S/D"

    return resultado


# ============================================================
# PREPARAR DATOS DEL PERIODO PARA DESCARGA
# ============================================================

def crear_csv_periodo(datos, variable, fecha_inicial, fecha_final):

    estaciones = datos[["estacion", "nombre_estacion", "lon", "lat"]].drop_duplicates(subset=["estacion"])
    fechas = pd.DataFrame({"Fecha": pd.date_range(fecha_inicial, fecha_final, freq="D").date})

    fechas["_union"] = 1
    estaciones["_union"] = 1
    estructura_completa = fechas.merge(estaciones, on="_union", how="outer").drop(columns="_union")

    seleccion = datos[
        (datos["Day"].dt.date >= fecha_inicial) & (datos["Day"].dt.date <= fecha_final)
    ].copy()
    seleccion["Fecha"] = seleccion["Day"].dt.date
    grupos = ["Fecha", "estacion"]

    if variable == "Rain_Total":
        valores = seleccion.groupby(grupos, as_index=False)[variable].sum(min_count=1)
    elif variable in ["T_Max", "Wind_Max", "Gust_Max"]:
        valores = seleccion.groupby(grupos, as_index=False)[variable].max()
    elif variable == "T_Min":
        valores = seleccion.groupby(grupos, as_index=False)[variable].min()
    else:
        valores = seleccion.groupby(grupos, as_index=False)[variable].mean()

    if variable in ["Wind_Avg", "Wind_Max"] and "Wind_Dir" in seleccion.columns:
        direcciones = seleccion.groupby(grupos, as_index=False)["Wind_Dir"].agg(promedio_circular)
        valores = valores.merge(direcciones, on=grupos, how="left")
        valores["direccion_cardinal"] = valores["Wind_Dir"].apply(direccion_cardinal)

    seleccion = estructura_completa.merge(valores, on=grupos, how="left").sort_values(["Fecha", "nombre_estacion"])
    nombre_valor = f'{VARIABLES[variable]["nombre"]} ({VARIABLES[variable]["unidad"]})'

    nombres_columnas = {
        "Fecha": "Fecha", "nombre_estacion": "Nombre de estación", "estacion": "Clave de estación",
        "lon": "Longitud", "lat": "Latitud", variable: nombre_valor,
        "Wind_Dir": "Dirección (grados)", "direccion_cardinal": "Dirección cardinal"
    }

    seleccion = seleccion.rename(columns=nombres_columnas).fillna("S/D")
    return seleccion.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")


# ============================================================
# LÍMITE PROVISIONAL
# ============================================================

def limite_provisional():
    return np.array([
        [-103.407, 20.710], [-103.401, 20.694], [-103.404, 20.676], [-103.396, 20.660],
        [-103.400, 20.647], [-103.392, 20.628], [-103.399, 20.613], [-103.383, 20.605],
        [-103.362, 20.609], [-103.347, 20.604], [-103.329, 20.626], [-103.322, 20.645],
        [-103.282, 20.640], [-103.268, 20.662], [-103.270, 20.688], [-103.265, 20.704],
        [-103.282, 20.712], [-103.288, 20.737], [-103.311, 20.746], [-103.331, 20.755],
        [-103.349, 20.719], [-103.374, 20.714], [-103.392, 20.711], [-103.407, 20.710]
    ])


# ============================================================
# LEER LÍMITE MUNICIPAL
# ============================================================

def leer_limite(ruta):
    if not ruta or not Path(ruta).exists():
        return limite_provisional(), False

    try:
        import geopandas as gpd
        limite_gdf = gpd.read_file(ruta).to_crs(epsg=4326)

        try:
            geometria = limite_gdf.geometry.union_all()
        except Exception:
            geometria = limite_gdf.geometry.unary_union

        if geometria.geom_type == "MultiPolygon":
            geometria = max(geometria.geoms, key=lambda poligono: poligono.area)

        coordenadas = np.asarray(geometria.exterior.coords)
        return coordenadas, True

    except Exception as error:
        st.warning(f"No se pudo abrir el límite municipal. Se utilizará uno provisional. {error}")
        return limite_provisional(), False


# ============================================================
# LEER LÍMITES DE COLONIAS
# ============================================================

def leer_colonias(ruta):
    if not ruta: return None, False

    ruta_colonias = Path(ruta)

    if ruta_colonias.is_dir():
        archivos_shp = sorted(ruta_colonias.glob("*.shp"))
        if not archivos_shp:
            st.warning("No se encontró un archivo .shp dentro de la carpeta de colonias.")
            return None, False
        ruta_colonias = archivos_shp[0]
    elif not ruta_colonias.exists():
        posible_shp = Path(f"{ruta}.shp")
        if posible_shp.exists(): ruta_colonias = posible_shp
        else:
            st.warning("No se encontró el shapefile de colonias.")
            return None, False

    try:
        import geopandas as gpd
        colonias_gdf = gpd.read_file(ruta_colonias).to_crs(epsg=4326)
        colonias_gdf = colonias_gdf[colonias_gdf.geometry.notna() & ~colonias_gdf.geometry.is_empty]
        return colonias_gdf, True
    except Exception as error:
        st.warning(f"No fue posible abrir los límites de colonias. {error}")
        return None, False


# ============================================================
# INTERPOLACIÓN IDW
# ============================================================

def interpolar_idw(puntos, variable, limite, resolucion=500, potencia=2):

    margen = 0.004
    xmin, xmax = limite[:, 0].min() - margen, limite[:, 0].max() + margen
    ymin, ymax = limite[:, 1].min() - margen, limite[:, 1].max() + margen

    x, y = np.linspace(xmin, xmax, resolucion), np.linspace(ymin, ymax, resolucion)
    malla_x, malla_y = np.meshgrid(x, y)

    numerador, denominador = np.zeros_like(malla_x, dtype=float), np.zeros_like(malla_x, dtype=float)

    for fila in puntos.itertuples():
        valor = getattr(fila, variable)
        distancia = ((malla_x - fila.lon) ** 2 + (malla_y - fila.lat) ** 2)
        distancia = np.maximum(distancia, 1e-12)
        peso = 1 / (distancia ** (potencia / 2))
        numerador += peso * valor
        denominador += peso

    superficie = numerador / denominador

    ruta = RutaMatplotlib(limite)
    puntos_malla = np.column_stack([malla_x.ravel(), malla_y.ravel()])
    mascara = ruta.contains_points(puntos_malla).reshape(malla_x.shape)
    superficie = np.ma.masked_where(~mascara, superficie)

    return malla_x, malla_y, superficie


# ============================================================
# AGREGAR LOGO Y ROSA AL MAPA
# ============================================================

def agregar_imagen_al_mapa(eje, ruta, posicion, escala):
    if not Path(ruta).exists(): return
    try:
        imagen = mpimg.imread(ruta)
        caja_imagen = OffsetImage(imagen, zoom=escala)
        elemento = AnnotationBbox(
            caja_imagen, posicion, xycoords=eje.transAxes,
            frameon=False, box_alignment=(0.5, 0.5), zorder=20
        )
        eje.add_artist(elemento)
    except Exception: pass


def catalogo_estaciones(datos_base):
    columnas = ["estacion", "lon", "lat"]
    ubicaciones = pd.DataFrame(datos_base.attrs.get("ubicaciones", []), columns=columnas)
    respaldo = pd.DataFrame([
        {"estacion": "dean", "lon": -103.348055, "lat": 20.63889},
        {"estacion": "base4", "lon": -103.292131, "lat": 20.662189},
    ])
    catalogo = pd.concat([ubicaciones, datos_base[columnas], respaldo], ignore_index=True)
    for columna in ["lon", "lat"]: catalogo[columna] = pd.to_numeric(catalogo[columna], errors="coerce")
    catalogo = catalogo.dropna(subset=["lon", "lat"]).drop_duplicates("estacion")
    catalogo["nombre_estacion"] = catalogo["estacion"].map(NOMBRES_ESTACIONES).fillna(catalogo["estacion"])
    return catalogo


def estaciones_sin_dato(puntos, catalogo):
    if catalogo is None: return pd.DataFrame(columns=["estacion", "nombre_estacion", "lon", "lat"])
    return catalogo.loc[~catalogo["estacion"].isin(puntos["estacion"])].copy()


def aplicar_disponibilidad_dean(datos):
    """DEAN solo aporta precipitación; conserva su ubicación para la etiqueta S/D."""
    resultado = datos.copy()
    mascara = resultado['estacion'].astype(str).str.strip().str.lower().eq('dean')
    columnas = [c for c in COLUMNAS_METEOROLOGICAS if c != 'Rain_Total' and c in resultado.columns]
    resultado.loc[mascara, columnas] = np.nan
    return resultado


def serie_diaria_periodo(datos, variable, fecha_inicial, fecha_final):
    if variable != 'Rain_Total':
        datos = datos.loc[~datos['estacion'].astype(str).str.strip().str.lower().eq('dean')]
    seleccion = datos.loc[datos["Day"].dt.date.between(fecha_inicial, fecha_final), ["Day", "nombre_estacion", variable]].copy()
    seleccion["Day"] = seleccion["Day"].dt.normalize()
    grupos = seleccion.groupby(["Day", "nombre_estacion"])[variable]

    if variable == "Rain_Total": serie = grupos.sum(min_count=1)
    elif variable in ["T_Max", "Wind_Max", "Gust_Max"]: serie = grupos.max()
    elif variable == "T_Min": serie = grupos.min()
    else: serie = grupos.mean()
    
    return serie.unstack("nombre_estacion").reindex(pd.date_range(fecha_inicial, fecha_final, freq="D"))


def crear_png_grafica_periodo(tabla, nombre, unidad, fecha_inicial, fecha_final, colores):
    import matplotlib.dates as mdates

    un_dia = fecha_inicial == fecha_final
    periodo = (f"{fecha_inicial:%d/%m/%Y}" if un_dia else f"{fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}")
    with plt.rc_context({"font.size": 14, "text.color": "#243746", "axes.labelcolor": "#243746", "xtick.color": "#243746", "ytick.color": "#243746"}):
        figura, eje = plt.subplots(figsize=(14, 6), dpi=400)
        try:
            eje.set_facecolor("white")
            if un_dia:
                valores = tabla.iloc[0].to_numpy(dtype=float)
                posiciones = np.arange(len(tabla.columns))
                validos = np.isfinite(valores)
                eje.plot(posiciones[validos], valores[validos], color="#3298DC", linewidth=2.3, zorder=2)
                for i in posiciones[validos]:
                    eje.scatter(i, valores[i], color=colores[i % len(colores)], s=65, edgecolor="white", linewidth=1, zorder=3)
                    eje.annotate(f"{valores[i]:.1f}", (i, valores[i]), xytext=(0, 10), textcoords="offset points", ha="center", fontsize=12)
                eje.set_xticks(posiciones)
                eje.set_xticklabels(tabla.columns, rotation=30, ha="right", fontsize=12)
                eje.set_xlabel("Estación", fontsize=14, fontweight="bold")
                eje.margins(y=0.20)
            else:
                for i, estacion in enumerate(tabla.columns):
                    eje.plot(tabla.index, tabla[estacion], label=estacion, color=colores[i % len(colores)], linewidth=2, marker="o", markersize=4)
                localizador = mdates.AutoDateLocator(minticks=3, maxticks=9)
                eje.xaxis.set_major_locator(localizador)
                eje.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m/%Y"))
                eje.tick_params(axis="x", labelrotation=25)
                eje.set_xlabel("Fecha", fontsize=14, fontweight="bold")
                eje.legend(loc="upper center", bbox_to_anchor=(0.5, -0.27), ncol=4, frameon=False, fontsize=12)
            
            eje.set_title(f"{nombre}\n{periodo}", loc="left", fontsize=22, fontweight="bold", pad=20)
            eje.set_ylabel(f"{nombre} ({unidad})", fontsize=14, fontweight="bold")
            eje.grid(color="#E8EEF4", linewidth=0.8)
            eje.set_axisbelow(True)
            for borde in ["top", "right"]: eje.spines[borde].set_visible(False)
            for borde in ["bottom", "left"]: eje.spines[borde].set_color("#CED9E3")
            figura.tight_layout()
            salida = io.BytesIO()
            figura.savefig(salida, format="png", dpi=400, bbox_inches="tight", facecolor="white")
            return salida.getvalue()
        finally:
            plt.close(figura)


def especificacion_grafica_interactiva(tabla, nombre, unidad, un_dia, colores):
    import json
    registros = []
    for fecha, fila in tabla.iterrows():
        for posicion, estacion in enumerate(tabla.columns):
            valor = fila[estacion]
            registros.append({'Fecha':pd.Timestamp(fecha).isoformat(), 'Estacion':str(estacion),
                              'Posicion':posicion, 'Valor':None if pd.isna(valor) else float(valor)})
    eje_y = {'field':'Valor','type':'quantitative','title':f'{nombre} ({unidad})',
             'scale':{'zero':False},'axis':{'format':'.1f','gridColor':'#efe5e8'}}
    tooltip = [{'field':'Estacion','title':'Estación','type':'nominal'},
               {'field':'Fecha','title':'Fecha','type':'temporal','format':'%d/%m/%Y'},
               {'field':'Valor','title':f'Valor ({unidad})','type':'quantitative','format':'.1f'}]
    color = {'field':'Estacion','type':'nominal','title':'Estación',
             'scale':{'domain':[str(c) for c in tabla.columns], 'range':[colores[i%len(colores)] for i in range(len(tabla.columns))]},
             'legend':{'orient':'bottom','columns':3,'labelLimit':240}}
    zoom = {'name':'vista','select':{'type':'interval','encodings':['x','y']},'bind':'scales'}
    spec = {'$schema':'https://vega.github.io/schema/vega-lite/v5.json',
            'data':{'values':registros}, 'height':420,
            'params':[zoom], 'config':{'view':{'stroke':None},'background':'#ffffff',
            'axis':{'labelColor':'#111111','titleColor':'#111111','labelFontSize':13,'titleFontSize':15},
            'legend':{'labelColor':'#111111','titleColor':'#111111','labelFontSize':13},
            'font':'Arial'}}
    if un_dia:
        nombres = json.dumps([str(c) for c in tabla.columns],ensure_ascii=False)
        eje_x = {'field':'Posicion','type':'quantitative','title':'Estación',
                 'scale':{'domain':[-0.5,max(len(tabla.columns)-0.5,0.5)]},
                 'axis':{'values':list(range(len(tabla.columns))), 'labelExpr':nombres+'[datum.value]',
                         'labelAngle':-30,'labelLimit':200,'grid':False}}
        spec['encoding']={'x':eje_x,'y':eje_y}
        spec['layer']=[{'mark':{'type':'line','color':'#3298DC','strokeWidth':2.5,'clip':True}},
                       {'mark':{'type':'point','filled':True,'size':85,'clip':True},
                        'encoding':{'color':dict(color,legend=None),'tooltip':tooltip}}]
        # Registrar el zoom una sola vez: las capas comparten ambas escalas.
        # En el nivel superior, Vega genera dos señales vista_Posicion y falla.
        spec['layer'][0]['params'] = spec.pop('params')
    else:
        spec['params'].append({'name':'estaciones_visibles','select':{'type':'point','fields':['Estacion']},'bind':'legend'})
        spec['mark']={'type':'line','point':{'filled':True,'size':45},'strokeWidth':2.3,'clip':True}
        spec['encoding']={'x':{'field':'Fecha','type':'temporal','title':'Fecha','axis':{'format':'%d/%m','grid':False}},
                          'y':eje_y,'color':color,'tooltip':tooltip,
                          'opacity':{'condition':{'param':'estaciones_visibles','value':1},'value':0.12}}
    return spec


def mostrar_grafica_periodo(datos, variable, fecha_inicial, fecha_final):
    st.subheader("Evolución diaria por estación")
    tabla = serie_diaria_periodo(datos, variable, fecha_inicial, fecha_final)
    nombre = "Precipitación diaria" if variable == "Rain_Total" else VARIABLES[variable]["nombre"]
    unidad = VARIABLES[variable]["unidad"]
    un_dia = fecha_inicial == fecha_final
    periodo = (f"{fecha_inicial:%d/%m/%Y}" if un_dia else f"{fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}")
    colores = ["#3298DC", "#EF8872", "#42B8A5", "#A385DB", "#E5AF36", "#E67DB1", "#58B9D1", "#91B94D", "#CB9370", "#798FE3", "#D99EC8", "#54AC84"]
    
    if tabla.empty or not tabla.notna().any().any():
        st.info('No hay datos para graficar en el periodo seleccionado.')
        return
    st.caption(f'{nombre} · {periodo}')
    png_grafica = crear_png_grafica_periodo(tabla, nombre, unidad, fecha_inicial, fecha_final, colores)
    vista_estatica = control_ruemag(st.checkbox, 'Ver gráfica estática (respaldo)', key='grafica_periodo_respaldo')
    if vista_estatica:
        st.image(png_grafica, use_column_width=True)
    else:
        try:
            spec = especificacion_grafica_interactiva(tabla, nombre, unidad, un_dia, colores)
            st.vega_lite_chart(spec=spec, use_container_width=True)
            st.caption('Rueda del ratón: zoom · Arrastrar: mover · Doble clic: restablecer. En varios días, pulsa una estación en la leyenda para resaltarla; Mayús + clic permite seleccionar varias.')
        except Exception as error:
            st.warning('No se pudo abrir la vista interactiva. Se muestra la gráfica estática.')
            st.image(png_grafica, use_column_width=True)
            with st.expander('Detalle del error de la gráfica'): st.code(str(error))
    with st.expander('Consultar valores diarios de la gráfica'):
        st.dataframe(tabla.rename_axis('Fecha'), use_container_width=True)
    st.caption('La descarga PNG contiene el periodo completo, sin el zoom aplicado en pantalla.')
    descarga_ruemag(label='⬇ Descargar evolución por estación (PNG)', data=png_grafica, file_name=f'RUEMAG_Grafica_{variable}_{fecha_inicial}_{fecha_final}.png', mime='image/png', key='descargar_grafica_periodo_png', titulo=f'Evolución de {nombre.lower()}', contexto=f'{periodo} · Periodo completo, sin el zoom de pantalla', formato='PNG')


def crear_niveles_mapa(puntos, variable):
    minimo = float(puntos[variable].min())
    maximo = float(puntos[variable].max())

    if variable in ["ANOM_Rain_Total", "ANOM_Gust_Max"]: return np.linspace(-100, 100, 21)
    if variable.startswith("ANOM_"):
        extremo = max(abs(minimo), abs(maximo))
        if math.isclose(extremo, 0.0): extremo = 1.0
        return np.linspace(-extremo, extremo, 13)
    if math.isclose(minimo, maximo): return np.linspace(minimo - 0.5, maximo + 0.5, 12)
    return np.linspace(minimo, maximo, 12)


# ============================================================
# CREAR MAPA (MÁXIMA CALIDAD - DPI 400 - VERSIÓN COMPACTA EVENTOS)
# ============================================================

def crear_mapa(
    puntos, variable, fecha_inicial, fecha_final, limite,
    colonias=None, titulo_personalizado=None, niveles_personalizados=None,
    colores_personalizados=None, ticks_personalizados=None, decimales=1,
    extension_personalizada=None, catalogo=None, compacto=False
):

    nombre_variable = VARIABLES[variable]["nombre"]
    unidad = VARIABLES[variable]["unidad"]
    colores = colores_personalizados if colores_personalizados is not None else VARIABLES[variable]["colores"]

    malla_x, malla_y, superficie = interpolar_idw(puntos, variable, limite)

    # Tamaño de figura más pequeño y estético para eventos si compacto=True
    tam_fig = (6.5, 5.8) if compacto else (10, 9)
    figura, eje = plt.subplots(figsize=tam_fig, dpi=400)
    eje.set_facecolor("#f3ecd8")

    niveles = np.asarray(niveles_personalizados, dtype=float) if niveles_personalizados is not None else crear_niveles_mapa(puntos, variable)
    mapa_colores = LinearSegmentedColormap.from_list("ruemag", colores)

    relleno = eje.contourf(
        malla_x, malla_y, superficie, levels=niveles, cmap=mapa_colores, alpha=0.93,
        extend=(extension_personalizada if extension_personalizada is not None else ("both" if variable in ["ANOM_Rain_Total", "ANOM_Gust_Max"] else "neither"))
    )

    if colonias is not None and not colonias.empty:
        colonias.boundary.plot(ax=eje, color="#555555", linewidth=0.35 if compacto else 0.45, alpha=0.80, zorder=4)

    eje.contour(malla_x, malla_y, superficie, levels=niveles, colors="white", linewidths=0.30, alpha=0.30)
    eje.plot(limite[:, 0], limite[:, 1], color="black", linewidth=1.5 if compacto else 2)

    tam_pto = 35 if compacto else 45
    tam_txt = 9 if compacto else 11
    eje.scatter(puntos["lon"], puntos["lat"], color="black", edgecolor="white", linewidth=0.4, s=tam_pto, zorder=5)

    for fila in estaciones_sin_dato(puntos, catalogo).itertuples():
        eje.scatter(fila.lon, fila.lat, color="#777777", edgecolor="white", s=tam_pto, zorder=5)
        eje.annotate(f"{fila.nombre_estacion}\nS/D", xy=(fila.lon, fila.lat), xytext=(5, 5), textcoords="offset points", fontsize=tam_txt, fontweight="bold", zorder=7, bbox=dict(boxstyle="round,pad=0.15", facecolor="white", edgecolor="#777777", alpha=0.90))

    mostrar_direccion = variable in ["Wind_Avg", "Wind_Max"]

    for fila in puntos.itertuples():
        valor = getattr(fila, variable)
        if mostrar_direccion:
            direccion_texto = getattr(fila, "direccion_cardinal")
            etiqueta = f"{fila.nombre_estacion}\n{valor:.{decimales}f} {unidad}\n{direccion_texto}"
        else:
            etiqueta = f"{fila.nombre_estacion}\n{valor:.{decimales}f} {unidad}"

        eje.annotate(etiqueta, xy=(fila.lon, fila.lat), xytext=(5, 5), textcoords="offset points", fontsize=tam_txt, fontweight="bold", zorder=7, bbox={"boxstyle": "round,pad=0.15", "facecolor": "white", "edgecolor": "black", "alpha": 0.90})

        if mostrar_direccion:
            direccion_grados = getattr(fila, "Wind_Dir")
            if not pd.isna(direccion_grados):
                angulo = np.deg2rad(direccion_grados)
                longitud_flecha = 0.012
                desplazamiento_x = -np.sin(angulo) * longitud_flecha
                desplazamiento_y = -np.cos(angulo) * longitud_flecha
                eje.arrow(fila.lon, fila.lat, desplazamiento_x, desplazamiento_y, width=0.00035, head_width=0.0022, head_length=0.0025, color="black", length_includes_head=True, zorder=6)

    if fecha_inicial == fecha_final: periodo = fecha_inicial.strftime("%d/%m/%Y")
    else: periodo = f"{fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}"

    titulo_mapa = titulo_personalizado if titulo_personalizado is not None else f"{nombre_variable} — {periodo}"
    tam_titulo = 18 if compacto else 28
    eje.set_title("\n".join(textwrap.fill(linea, width=42) for linea in titulo_mapa.splitlines()), fontsize=tam_titulo, fontweight="bold", pad=12)
    eje.set_xlabel("Longitud", fontsize=12 if compacto else 16, fontweight="bold")
    eje.set_ylabel("Latitud", fontsize=12 if compacto else 16, fontweight="bold")
    eje.tick_params(labelsize=10 if compacto else 14)
    eje.grid(color="#aebcbc", linestyle=":", linewidth=0.8, alpha=0.70)
    eje.set_aspect("equal", adjustable="box")

    barra = figura.colorbar(relleno, ax=eje, shrink=0.78 if compacto else 0.82, pad=0.035)
    barra.set_label(unidad, fontsize=12 if compacto else 16, fontweight="bold")

    if ticks_personalizados is not None:
        barra.set_ticks(ticks_personalizados)
        barra.ax.set_yticklabels([f"{tick:.{decimales}f}" for tick in ticks_personalizados], fontsize=10 if compacto else 13)
    elif variable in ["ANOM_Rain_Total", "ANOM_Gust_Max"]:
        barra.set_ticks(np.arange(-100, 101, 20))
        barra.ax.yaxis.set_major_formatter(FormatStrFormatter(f'%.{decimales}f'))
    else:
        barra.ax.yaxis.set_major_formatter(FormatStrFormatter(f'%.{decimales}f'))
        barra.ax.tick_params(labelsize=10 if compacto else 13)

    agregar_imagen_al_mapa(eje=eje, ruta=RUTA_LOGO, posicion=(0.15, 0.88), escala=0.12 if compacto else 0.16)
    agregar_imagen_al_mapa(eje=eje, ruta=RUTA_ROSA, posicion=(0.88, 0.88), escala=0.045 if compacto else 0.055)

    figura.tight_layout()
    archivo_imagen = io.BytesIO()
    figura.savefig(archivo_imagen, format="png", dpi=400, bbox_inches=None if compacto else "tight", facecolor="white")
    plt.close(figura)
    archivo_imagen.seek(0)
    
    return archivo_imagen.getvalue()


# ============================================================
# CREAR KMZ CON EL MAPA IDW Y LAS ESTACIONES
# ============================================================

def crear_kmz(puntos, variable, fecha_inicial, fecha_final, limite, colonias=None, catalogo=None, niveles_personalizados=None, colores_personalizados=None, extension_personalizada=None):
    nombre_variable = VARIABLES[variable]["nombre"]
    unidad = VARIABLES[variable]["unidad"]
    mostrar_direccion = variable in ["Wind_Avg", "Wind_Max"]
    malla_x, malla_y, superficie = interpolar_idw(puntos, variable, limite)

    niveles = (np.asarray(niveles_personalizados, dtype=float) if niveles_personalizados is not None else crear_niveles_mapa(puntos, variable))
    mapa_colores = LinearSegmentedColormap.from_list("ruemag_kmz", colores_personalizados if colores_personalizados is not None else VARIABLES[variable]["colores"])

    proporcion = (malla_x.max() - malla_x.min()) / (malla_y.max() - malla_y.min())
    figura_overlay = plt.figure(figsize=(8 * proporcion, 8), dpi=250, frameon=False)
    eje_overlay = figura_overlay.add_axes([0, 0, 1, 1])

    eje_overlay.contourf(malla_x, malla_y, superficie, levels=niveles, cmap=mapa_colores, alpha=0.93, extend=(extension_personalizada if extension_personalizada is not None else ("both" if variable in ["ANOM_Rain_Total", "ANOM_Gust_Max"] else "neither")))
    if colonias is not None and not colonias.empty: colonias.boundary.plot(ax=eje_overlay, color="#555555", linewidth=0.55, alpha=0.85, zorder=4)

    eje_overlay.contour(malla_x, malla_y, superficie, levels=niveles, colors="white", linewidths=0.35, alpha=0.30)
    eje_overlay.plot(limite[:, 0], limite[:, 1], color="black", linewidth=1.2)
    eje_overlay.set_xlim(malla_x.min(), malla_x.max())
    eje_overlay.set_ylim(malla_y.min(), malla_y.max())
    eje_overlay.axis("off")

    imagen_overlay = io.BytesIO()
    figura_overlay.savefig(imagen_overlay, format="png", transparent=True, dpi=250)
    plt.close(figura_overlay)
    imagen_overlay.seek(0)

    marcadores = []
    for fila in puntos.itertuples():
        valor = getattr(fila, variable)
        nombre = escape(str(fila.nombre_estacion))
        if mostrar_direccion:
            direccion = getattr(fila, "direccion_cardinal")
            descripcion = escape(f"{nombre_variable}: {valor:.1f} {unidad}. Dirección: {direccion}")
        else:
            descripcion = escape(f"{nombre_variable}: {valor:.1f} {unidad}")
        
        marcadores.append(f"""
        <Placemark>
            <name>{nombre}</name>
            <description>{descripcion}</description>
            <Point><coordinates>{fila.lon},{fila.lat},0</coordinates></Point>
        </Placemark>
        """)

    for fila in estaciones_sin_dato(puntos, catalogo).itertuples():
        marcadores.append(f"""
        <Placemark>
            <name>{escape(str(fila.nombre_estacion))} — S/D</name>
            <description>Sin datos para la variable y el periodo seleccionados.</description>
            <Point><coordinates>{fila.lon},{fila.lat},0</coordinates></Point>
        </Placemark>
        """)

    contenido_kml = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
    <Document>
        <name>RUEMAG - {escape(nombre_variable)}</name>
        <description>{escape(f"Periodo: {fecha_inicial} a {fecha_final}. Variable: {nombre_variable} ({unidad}).")}</description>
        <GroundOverlay>
            <name>Mapa interpolado IDW</name>
            <Icon><href>mapa_idw.png</href></Icon>
            <LatLonBox>
                <north>{float(malla_y.max())}</north><south>{float(malla_y.min())}</south>
                <east>{float(malla_x.max())}</east><west>{float(malla_x.min())}</west>
            </LatLonBox>
        </GroundOverlay>
        {''.join(marcadores)}
    </Document>
</kml>
"""

    archivo_kmz = io.BytesIO()
    with zipfile.ZipFile(archivo_kmz, mode="w", compression=zipfile.ZIP_DEFLATED) as kmz:
        kmz.writestr("doc.kml", contenido_kml.encode("utf-8"))
        kmz.writestr("mapa_idw.png", imagen_overlay.getvalue())
    archivo_kmz.seek(0)
    return archivo_kmz.getvalue()


# ============================================================
# CREAR SHAPEFILE DE LA CAPA INTERPOLADA IDW
# ============================================================

def crear_estilo_qgis_idw(variable, niveles, colores_personalizados=None):
    mapa_colores = LinearSegmentedColormap.from_list("ruemag_qgis", colores_personalizados if colores_personalizados is not None else VARIABLES[variable]["colores"])
    categorias = []
    simbolos = []
    total_intervalos = len(niveles) - 1

    for indice in range(total_intervalos):
        minimo, maximo = float(niveles[indice]), float(niveles[indice + 1])
        posicion_color = (indice + 0.5) / total_intervalos if total_intervalos > 0 else 0.5
        color_hex = to_hex(mapa_colores(posicion_color), keep_alpha=False)
        rojo, verde, azul = [int(round(canal * 255)) for canal in to_rgb(color_hex)]
        valor_categoria, etiqueta = format(minimo, ".15g"), f"{minimo:.1f} a {maximo:.1f}"
        categorias.append(f'<category value="{valor_categoria}" label="{etiqueta}" symbol="{indice}" render="true"/>')
        simbolos.append(f"""
            <symbol name="{indice}" type="fill" alpha="1" clip_to_extent="1" force_rhr="0">
              <layer class="SimpleFill" enabled="1" locked="0" pass="0">
                <Option type="Map">
                  <Option name="color" value="{rojo},{verde},{azul},237" type="QString"/>
                  <Option name="outline_color" value="255,255,255,75" type="QString"/>
                  <Option name="outline_style" value="solid" type="QString"/>
                  <Option name="outline_width" value="0.15" type="QString"/>
                  <Option name="outline_width_unit" value="MM" type="QString"/>
                  <Option name="style" value="solid" type="QString"/>
                </Option>
              </layer>
            </symbol>
            """)

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<qgis version="3.28.0" styleCategories="Symbology">
  <renderer-v2 type="categorizedSymbol" attr="VAL_MIN" symbollevels="0" enableorderby="0" forceraster="0">
    <categories>{''.join(categorias)}</categories>
    <symbols>{''.join(simbolos)}</symbols>
  </renderer-v2>
  <blendMode>0</blendMode><featureBlendMode>0</featureBlendMode><layerOpacity>1</layerOpacity>
</qgis>
"""

def crear_shapefile_idw(puntos, variable, limite, niveles_personalizados=None, colores_personalizados=None):
    import geopandas as gpd
    from shapely.geometry import Polygon

    malla_x, malla_y, superficie = interpolar_idw(puntos, variable, limite)
    niveles = (np.asarray(niveles_personalizados, dtype=float) if niveles_personalizados is not None else crear_niveles_mapa(puntos, variable))
    
    if niveles_personalizados is not None:
        if np.nanmax(superficie) > niveles[-1]: niveles = np.append(niveles, np.nanmax(superficie) + 0.001)
        if np.nanmin(superficie) < niveles[0]: niveles = np.insert(niveles, 0, np.nanmin(superficie) - 0.001)

    if variable in ["ANOM_Rain_Total", "ANOM_Gust_Max"]: superficie = np.clip(superficie, -100, 100)

    figura, eje = plt.subplots(dpi=400)
    contornos = eje.contourf(malla_x, malla_y, superficie, levels=niveles)
    registros = []

    for indice_nivel, segmentos in enumerate(contornos.allsegs):
        if indice_nivel >= len(niveles) - 1: continue
        for segmento in segmentos:
            if len(segmento) < 4: continue
            poligono = Polygon(segmento)
            if not poligono.is_valid: poligono = poligono.buffer(0)
            if poligono.is_empty: continue
            registros.append({"VAL_MIN": float(niveles[indice_nivel]), "VAL_MAX": float(niveles[indice_nivel + 1]), "VARIABLE": variable, "geometry": poligono})
    plt.close(figura)

    if not registros: raise ValueError("No fue posible construir los polígonos de la capa IDW.")
    capa = gpd.GeoDataFrame(registros, geometry="geometry", crs="EPSG:4326")

    archivo_zip = io.BytesIO()
    with tempfile.TemporaryDirectory() as carpeta_temporal:
        ruta_shp = Path(carpeta_temporal) / "RUEMAG_IDW.shp"
        capa.to_file(ruta_shp, driver="ESRI Shapefile", encoding="utf-8")
        ruta_qml = Path(carpeta_temporal) / "RUEMAG_IDW.qml"
        ruta_qml.write_text(crear_estilo_qgis_idw(variable, niveles, colores_personalizados), encoding="utf-8")
        with zipfile.ZipFile(archivo_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as archivo:
            for componente in Path(carpeta_temporal).glob("RUEMAG_IDW.*"):
                archivo.write(componente, arcname=componente.name)
    archivo_zip.seek(0)
    return archivo_zip.getvalue()


# ============================================================
# TABLA TOP 12
# ============================================================

def crear_tabla_top(puntos, variable, unidad):
    mostrar_direccion = variable in ["Wind_Avg", "Wind_Max"]
    tabla = puntos.copy().sort_values(variable, ascending=False).head(12)
    filas_html = []
    posicion = 1

    for fila in tabla.itertuples():
        valor = getattr(fila, variable)
        columna_direccion = f"<td>{getattr(fila, 'direccion_cardinal')}</td>" if mostrar_direccion else ""
        fila_html = f"<tr><td>{posicion}</td><td>{escape(str(fila.nombre_estacion))}</td><td>{valor:.1f} {unidad}</td>{columna_direccion}</tr>"
        filas_html.append(fila_html)
        posicion += 1

    encabezado_direccion = "<th scope=col>Dirección</th>" if mostrar_direccion else ""
    tabla_html = f"""<div class="contenedor-top" role="region" aria-label="Registros por estación" tabindex="0"><table class="tabla-top">
        <thead><tr><th scope=col>Pos.</th><th scope=col>Estación</th><th scope=col>Valor</th>{encabezado_direccion}</tr></thead>
        <tbody>{''.join(filas_html)}</tbody></table></div>"""
    return tabla_html


# ============================================================
# ENCABEZADO PRINCIPAL CON LOGO
# ============================================================



encabezado_ruemag()
st.markdown('<span class="ru-nav-anchor"></span>', unsafe_allow_html=True)

seccion_pagina = control_ruemag(st.radio, 'Sección', ['🗺️ Mapas meteorológicos', '📊 Anomalías', '⛈️ Eventos de tormenta'], horizontal=True, label_visibility='collapsed', key='ruemag_modulo', format_func=lambda valor: valor.split(' ', 1)[1])


# ============================================================
# BARRA LATERAL (PROTEGIDA CON CONTRASEÑA)
# ============================================================

archivo_subido = None
archivo_climatologia = None
archivo_rachas_historicas = None

ruta_excel = RUTA_EXCEL
ruta_limite = RUTA_LIMITE
ruta_colonias = RUTA_COLONIAS
ruta_climatologia = RUTA_CLIMATOLOGIA
ruta_rachas_historicas = RUTA_RACHAS_HISTORICAS

with st.sidebar:
    st.markdown('<div class="ru-sidebar-titulo">' + icono_ruemag('candado') + '<h2>Configuración del sistema</h2></div>', unsafe_allow_html=True)
    contrasena_ingresada = st.text_input("Contraseña de acceso", type="password", key="pwd_sidebar", help="Acceso a la configuración y a las fuentes de datos. La contraseña se verifica en el servidor.")
    
    CONTRASEÑA_CORRECTA = os.environ.get("RUEMAG_PASSWORD", "")
    if not CONTRASEÑA_CORRECTA:
        try:
            CONTRASEÑA_CORRECTA = st.secrets.get("RUEMAG_PASSWORD", "")
        except FileNotFoundError:
            pass

    if CONTRASEÑA_CORRECTA and contrasena_ingresada == CONTRASEÑA_CORRECTA:
        st.success("Acceso concedido a configuración")
        st.caption("Excel RUEMAG: ruemag2base.xlsx")
        st.caption("Límite municipal: limite_gdl.gpkg")
        st.caption("Colonias: colonias_2022")

        if seccion_pagina == "⛈️ Eventos de tormenta":
            st.caption("Eventos: ACTUALIZACION_JULIO")
            st.caption("Histórico de Eventos: eventos.xlsx (todos los años disponibles)")

        if seccion_pagina == "📊 Anomalías":
            st.markdown("---")
            st.subheader("Climatología de anomalías")
            st.caption("Temperatura y precipitación: gdl_completo_precip.csv")
            st.caption("Rachas: seriecompletaiam+ruemagcentro.xlsx")
    else:
        if not CONTRASEÑA_CORRECTA:
            st.info("La configuración avanzada está bloqueada. Define RUEMAG_PASSWORD en los secretos de Streamlit o en las variables de entorno del servidor.")
        elif contrasena_ingresada != "":
            st.error("Contraseña incorrecta")
        else:
            st.info("Ingresa la contraseña para ver las rutas y configuraciones avanzadas.")


# ============================================================
# SELECCIONAR EXCEL
# ============================================================

fuente_excel = ruta_excel

if not Path(ruta_excel).exists():
    st.error("No se encontró el archivo Excel RUEMAG en la ruta fija:")
    st.code(ruta_excel)
    st.stop()


# ============================================================
# CARGAR DATOS
# ============================================================

try:
    with st.spinner("Leyendo el archivo Excel..."):
        datos = aplicar_disponibilidad_dean(leer_excel(fuente_excel))
except Exception as error:
    st.error("No fue posible leer el Excel.")
    st.exception(error)
    st.stop()


estado_datos_ruemag(datos)

# ============================================================
# SECCIÓN DE EVENTOS DE TORMENTA
# ============================================================

# Reportes complementarios: lectura sin modificar los Excel de mapas.
def validar_base_reportes(fuente):
    requeridas = {
        'Eventos': ['Anio','Evento','Emision','Inicio','Maxima_intensidad','Fin','Situacion','Nivel_DEAN','Avisos','Afectaciones_reportadas','Colonias_mas_afectadas'],
        'Mediciones': ['Anio','Evento','Estacion','Lluvia_mm','Racha_kmh','Intensidad_max_mmh','Rayos_max'],
        'Afectaciones': ['Anio','Evento','Servicio_ID','Categoria','Tipo','Direccion','Colonia','Latitud','Longitud']}
    libro = pd.read_excel(fuente, sheet_name=None)
    for hoja, columnas in requeridas.items():
        if hoja not in libro:
            raise ValueError(f'Falta la hoja {hoja}. Usa la plantilla reportes_eventos.xlsx.')
        df = libro[hoja].dropna(how='all').copy()
        faltan = set(columnas) - set(df.columns)
        if faltan:
            raise ValueError(f'{hoja}: faltan columnas {", ".join(sorted(faltan))}.')
        for col in ['Anio','Evento']:
            n = pd.to_numeric(df[col], errors='coerce')
            if (n.isna() | (n % 1 != 0) | (n < 1)).any():
                raise ValueError(f'{hoja}: {col} debe contener enteros positivos en todas las filas.')
            df[col] = n.astype(int)
        claves = ['Anio','Evento'] + ([] if hoja=='Eventos' else ['Estacion' if hoja=='Mediciones' else 'Servicio_ID'])
        if df[claves].isna().any().any() or df.duplicated(claves).any():
            raise ValueError(f'{hoja}: hay identificadores vacíos o duplicados ({", ".join(claves)}).')
        nums = {'Eventos':['Avisos','Afectaciones_reportadas'], 'Mediciones':columnas[3:], 'Afectaciones':['Latitud','Longitud']}[hoja]
        for col in nums:
            n = pd.to_numeric(df[col], errors='coerce')
            if (df[col].notna() & n.isna()).any():
                raise ValueError(f'{hoja}: {col} contiene texto donde se espera un número; deja vacío si no hay dato.')
            if col not in ['Latitud','Longitud'] and (n.dropna()<0).any():
                raise ValueError(f'{hoja}: {col} no admite valores negativos.')
            if col in ['Avisos','Afectaciones_reportadas','Rayos_max'] and (n.dropna()%1!=0).any():
                raise ValueError(f'{hoja}: {col} debe ser un conteo entero.')
            df[col] = n
        if hoja=='Eventos':
            for col in ['Emision','Inicio','Maxima_intensidad','Fin']:
                fechas = pd.to_datetime(df[col],errors='coerce',dayfirst=True)
                if (df[col].notna() & fechas.isna()).any():
                    raise ValueError(f'Eventos: revisa la fecha y hora en {col}.')
                df[col] = fechas
            if (df['Fin']<df['Inicio']).any() or ((df['Maxima_intensidad']<df['Inicio']) | (df['Maxima_intensidad']>df['Fin'])).any():
                raise ValueError('Eventos: revisa el orden Inicio → Máxima intensidad → Fin.')
            if (df['Inicio'].notna() & (df['Inicio'].dt.year!=df['Anio'])).any():
                raise ValueError('Eventos: Anio debe coincidir con el año de inicio.')
        libro[hoja] = df
    claves = set(map(tuple, libro['Eventos'][['Anio','Evento']].values))
    for hoja in ['Mediciones','Afectaciones']:
        if not set(map(tuple, libro[hoja][['Anio','Evento']].values)).issubset(claves):
            raise ValueError(f'{hoja}: hay registros sin una fila correspondiente en Eventos.')
    return libro


def texto_reporte(v):
    return 'S/D' if pd.isna(v) or str(v).strip()=='' else str(v)


def fecha_reporte(v):
    return 'S/D' if pd.isna(v) else pd.Timestamp(v).strftime('%d/%m/%Y | %H:%M h')


def reporte_markdown(r, mediciones, afectaciones):
    lineas=[f"# REPORTE FINAL DE EVENTO DE LLUVIA #{int(r['Evento'])} EN GUADALAJARA", fecha_reporte(r['Emision']), '',
            f"Inicio: {fecha_reporte(r['Inicio'])}", f"Hora de máxima intensidad: {fecha_reporte(r['Maxima_intensidad'])}",
            f"Fin: {fecha_reporte(r['Fin'])}", '', f"Situación: {texto_reporte(r['Situacion'])}", '']
    for col,titulo,unidad in [('Intensidad_max_mmh','⛈ Intensidad de lluvia máxima','mm/h'),('Lluvia_mm','🌧 RUEMAG · Lluvia acumulada','mm'),('Racha_kmh','💨 Rachas de viento','km/h'),('Rayos_max','⚡ Cantidad máxima de rayos','rayos')]:
        lineas += [f'## {titulo} ({unidad})']
        d=mediciones.sort_values(col,ascending=False,na_position='last')
        if col in ['Intensidad_max_mmh','Rayos_max']:
            d=d[d[col].notna()]
            d=d[d[col]==d[col].max()] if not d.empty else d
        if d.empty: lineas.append('S/D')
        for _,row in d.iterrows():
            valor='S/D' if pd.isna(row[col]) else (f'{row[col]:.0f}' if col=='Rayos_max' else f'{row[col]:.1f}')
            lineas.append(f"- {valor}, {row['Estacion']}")
        lineas.append('')
    lineas += ['## 🌊 Nivel máximo de agua (DEAN)',texto_reporte(r['Nivel_DEAN']), '## ⚠️ Avisos emitidos',texto_reporte(r['Avisos']),
               '## 🚒 Afectaciones',f"Total reportado: {texto_reporte(r['Afectaciones_reportadas'])} · Registros de detalle: {len(afectaciones)}"]
    for categoria,g in afectaciones.groupby('Categoria',sort=False,dropna=False):
        lineas.append(f'### {texto_reporte(categoria)} ({len(g)})')
        for tipo,sub in g.groupby('Tipo',sort=False,dropna=False):
            lineas.append(f"**{texto_reporte(tipo)} ({len(sub)})**")
            lineas += [f'- {texto_reporte(v)}' for v in sub['Direccion']]
    lineas += ['## Colonias más afectadas',texto_reporte(r['Colonias_mas_afectadas'])]
    return '\n\n'.join(lineas)


def fuentes_eventos_anio(carpeta, anio, archivos=None, tipo_forzado=None):
    """Todas las hojas se filtran por Fecha; nunca se asocia solo por número."""
    resultado={'lluvia':[], 'rachas':[]}; errores=[]
    fuentes=archivos if archivos is not None else (sorted(Path(carpeta).rglob('*.xlsx')) if Path(carpeta).exists() else [])
    for archivo in fuentes:
        nombre=getattr(archivo,'name',str(archivo))
        if Path(nombre).name.startswith('~$'):continue
        try:
            contenido=io.BytesIO(archivo.getvalue()) if hasattr(archivo,'getvalue') else archivo
            with pd.ExcelFile(contenido) as libro:
                for hoja in libro.sheet_names:
                    texto=normalizar_texto_evento(hoja)
                    tipo=tipo_forzado
                    for identificacion in [texto,normalizar_texto_evento(Path(nombre).stem)]:
                        if tipo:break
                        if any(t in identificacion for t in ['racha','viento','gust']):tipo='rachas'
                        elif any(t in identificacion for t in ['lluvia','precipitacion','rain']):tipo='lluvia'
                    if tipo is None:continue
                    d=pd.read_excel(libro,sheet_name=hoja)
                    equivalencias={normalizar_texto_evento(c):c for c in d.columns}
                    if not {'evento','fecha'}.issubset(equivalencias):continue
                    d=d.rename(columns={equivalencias['evento']:'Evento',equivalencias['fecha']:'Fecha'})
                    d['Evento']=pd.to_numeric(d['Evento'],errors='coerce')
                    d['Fecha']=pd.to_datetime(d['Fecha'],errors='coerce',dayfirst=True)
                    d=d[(d.Fecha.dt.year==anio)&d.Evento.notna()].copy()
                    if d.empty:continue
                    if (d.Evento%1!=0).any() or (d.Evento<1).any():
                        errores.append(f'{nombre} / {hoja}: hay números de evento inválidos.');continue
                    d['Evento']=d.Evento.astype(int)
                    if d.duplicated('Evento').any():
                        errores.append(f'{nombre} / {hoja}: hay eventos duplicados en {anio}; revisa esa hoja.');continue
                    resultado[tipo].append((f'{nombre} · {hoja}',d))
        except Exception as e:errores.append(f'{nombre}: {e}')
    return resultado,errores


def fechas_eventos_disponibles(tabla_lluvia, tabla_rachas, base, anio):
    """Relaciona eventos con días observados, incluidos cruces de medianoche."""
    fechas = {}
    for tabla in [tabla_lluvia, tabla_rachas]:
        for _, fila in tabla.iterrows():
            numero = int(fila['Evento'])
            fechas.setdefault(numero, set())
            fecha = pd.to_datetime(fila['Fecha'], errors='coerce')
            if pd.notna(fecha) and fecha.year == anio:
                fechas[numero].add(fecha.date())
    if base is not None:
        for _, fila in base['Eventos'].loc[base['Eventos'].Anio == anio].iterrows():
            numero = int(fila['Evento'])
            fechas.setdefault(numero, set())
            inicio, fin = fila['Inicio'], fila['Fin']
            if pd.notna(inicio):
                ultimo = fin if pd.notna(fin) else inicio
                for dia in pd.date_range(pd.Timestamp(inicio).normalize(), pd.Timestamp(ultimo).normalize()):
                    if dia.year == anio: fechas[numero].add(dia.date())
            elif pd.notna(fin) and pd.Timestamp(fin).year == anio:
                fechas[numero].add(pd.Timestamp(fin).date())
    return fechas


def mostrar_reporte_lateral(base, anio, numero_evento):
    st.markdown('### Reporte final')
    if base is None:
        st.info('El reporte aún no está cargado. Se puede agregar desde Configuración con contraseña.');return
    partes={h:d[(d.Anio==anio)&(d.Evento==numero_evento)] for h,d in base.items() if h in ['Eventos','Mediciones','Afectaciones']}
    if partes['Eventos'].empty:
        st.info(f'Sin reporte capturado para {anio} · Evento #{numero_evento}.');return
    r=partes['Eventos'].iloc[0];m=partes['Mediciones'];a=partes['Afectaciones']
    st.caption(f"GUADALAJARA · {anio} · EVENTO #{numero_evento}")
    st.caption('Emisión: '+fecha_reporte(r.Emision)+' · Hora local (UTC−6)')
    for titulo,col in [('Inicio','Inicio'),('Máxima intensidad','Maxima_intensidad'),('Fin','Fin')]:
        st.write(f'**{titulo}:** {fecha_reporte(r[col])}')
    if pd.notna(r.Inicio) and pd.notna(r.Fin):
        st.caption(f'Duración: {(r.Fin-r.Inicio).total_seconds()/60:g} minutos')
    st.markdown('#### Situación')
    st.write(texto_reporte(r.Situacion))
    for col,label,unit in [('Lluvia_mm','🌧 Mayor acumulado','mm'),('Racha_kmh','💨 Racha máxima','km/h'),('Intensidad_max_mmh','⛈ Intensidad máxima','mm/h'),('Rayos_max','⚡ Máximo de rayos','rayos')]:
        d=m.dropna(subset=[col]);v=d[col].max() if not d.empty else float('nan')
        valor='S/D' if pd.isna(v) else f'{v:g} {unit}'
        estaciones=' · '.join(d.loc[d[col]==v,'Estacion'].astype(str))
        st.write(f'**{label}: {valor}**')
        if estaciones:st.caption(estaciones)
    st.write('**⚠️ Avisos:** '+('S/D' if pd.isna(r.Avisos) else f'{r.Avisos:.0f}'))
    st.write('**🚒 Afectaciones reportadas:** '+('S/D' if pd.isna(r.Afectaciones_reportadas) else f'{r.Afectaciones_reportadas:.0f}'))
    if pd.notna(r.Afectaciones_reportadas) and r.Afectaciones_reportadas!=len(a):
        st.warning(f'El detalle tiene {len(a)} registros; no coincide con el total reportado.')
    for col,title in [('Lluvia_mm','🌧 Lluvia por estación (mm)'),('Racha_kmh','💨 Rachas por estación (km/h)')]:
        with st.expander(title):
            d=m[['Estacion',col]].sort_values(col,ascending=False,na_position='last').copy()
            d[col]=d[col].map(lambda v:'S/D' if pd.isna(v) else f'{v:.1f}')
            st.dataframe(d.rename(columns={'Estacion':'Estación',col:'Valor'}).set_index('Estación'),use_container_width=True)
    with st.expander('🌊 Nivel máximo de agua · DEAN'):st.write(texto_reporte(r.Nivel_DEAN))
    for cat,g in a.groupby('Categoria',sort=False,dropna=False):
        with st.expander(f'{texto_reporte(cat)} · {len(g)} servicios'):
            for tipo,sub in g.groupby('Tipo',sort=False,dropna=False):
                st.write(f'**{texto_reporte(tipo)} ({len(sub)})**')
                for direccion in sub.Direccion:st.write('• '+texto_reporte(direccion))
    st.write('**Colonias más afectadas:** '+texto_reporte(r.Colonias_mas_afectadas))
    with st.expander('📍 Ubicación de afectaciones'):
        valid=a.Latitud.between(-90,90)&a.Longitud.between(-180,180)
        geo=a.loc[valid,['Latitud','Longitud']].rename(columns={'Latitud':'lat','Longitud':'lon'})
        if geo.empty:st.info('Sin coordenadas capturadas para este evento.')
        else:st.map(geo)
        st.caption(f'{len(geo)} de {len(a)} servicios con coordenadas válidas.')
    texto=reporte_markdown(r,m,a)
    descarga_ruemag(f'⬇ Descargar reporte del evento #{numero_evento} (TXT)', texto.encode('utf-8-sig'), f'Reporte_{anio}_evento_{numero_evento}.txt', 'text/plain', key=f'reporte_{anio}_{numero_evento}', titulo='Reporte final del evento', contexto=f'Guadalajara · {anio} · Evento #{numero_evento}', formato='TXT')
    with st.expander('Copiar reporte completo'):st.code(texto,language=None)


if seccion_pagina == "⛈️ Eventos de tormenta":

    st.markdown('<div class="ev-header"><h2>Eventos de lluvia</h2><p>RUEMAG · Guadalajara | Reportes, mapas y evolución del temporal</p></div>', unsafe_allow_html=True)

    tab_mapas, tab_historico = st.tabs(["Eventos individuales", "Histórico de temporales"])
    base_reportes = None
    with tab_mapas:
        acceso_configuracion = contrasena_ingresada == CONTRASEÑA_CORRECTA
        if acceso_configuracion:
            with st.sidebar.expander('📂 Base de reportes', expanded=False):
                reporte_subido = st.file_uploader('Base de reportes (todos los años)',type=['xlsx'],key='base_reportes_eventos')
                if reporte_subido is not None:
                    st.session_state['_reporte_aplicado'] = reporte_subido.getvalue()
                if st.button('Usar base local',key='usar_base_local'):
                    st.session_state.pop('_reporte_aplicado',None)
        reporte_bytes = st.session_state.get('_reporte_aplicado')
        fuente_reporte = io.BytesIO(reporte_bytes) if reporte_bytes is not None else Path(__file__).resolve().parent / 'reportes_eventos.xlsx'
        if reporte_bytes is not None or fuente_reporte.exists():
            try: base_reportes = validar_base_reportes(fuente_reporte)
            except Exception as error: st.error(f'Revisa la base de reportes: {error}')

        anios = {2024, 2025, 2026}
        if base_reportes is not None: anios.update(base_reportes['Eventos'].Anio.astype(int))
        if Path(RUTA_HISTORICO_EVENTOS).exists():
            try:
                with pd.ExcelFile(RUTA_HISTORICO_EVENTOS) as libro_anios:
                    anios.update(int(h) for h in libro_anios.sheet_names if str(h).isdigit() and len(str(h))==4)
            except Exception as error: st.caption(f"No se pudieron consultar los años del histórico: {error}")
        for carpeta_temporal in Path(RUTA_EVENTOS).parent.parent.glob('TEMPORAL *'):
            sufijo=carpeta_temporal.name.split()[-1]
            if sufijo.isdigit() and len(sufijo)==4: anios.add(int(sufijo))
        anios.add(2023)
        st.markdown('### Consulta por año, evento o fecha')
        columna_anio, columna_evento = st.columns([1, 2])
        with columna_anio:
            anio_evento = control_ruemag(st.selectbox, '1. Año del evento', sorted(anios, reverse=True), key='anio_evento_individual')
        if anio_evento == 2023:
            numero_evento = None
            with columna_evento:
                fecha_consulta = control_ruemag(st.date_input, 'Fecha de consulta', value=datetime.date(2023, 1, 1), min_value=datetime.date(2023, 1, 1), max_value=datetime.date(2023, 12, 31), format='DD/MM/YYYY', key='fecha_evento_2023')
            st.info('No disponible')
        else:
            ruta_predeterminada = RUTA_EVENTOS if anio_evento==2026 else str(Path(RUTA_EVENTOS).parent.parent / f'TEMPORAL {anio_evento}')
            configuraciones = st.session_state.setdefault('_archivos_eventos_aplicados', {})
            configuracion = configuraciones.setdefault(anio_evento, {})
            carpeta_anio = configuracion.get('carpeta', ruta_predeterminada)
            carga_lluvia, carga_racha = configuracion.get('lluvia'), configuracion.get('rachas')
            if acceso_configuracion:
                ajustes_eventos = st.sidebar.expander(f'⚙️ Archivos de eventos · {anio_evento}', expanded=False)
                with ajustes_eventos:
                    carpeta_anio = st.text_input('Carpeta con los Excel de lluvia y rachas',value=carpeta_anio,key=f'carpeta_eventos_{anio_evento}')
                    nueva_lluvia = st.file_uploader('Excel de lluvia (opcional)',type=['xlsx'],key=f'lluvia_{anio_evento}')
                    nueva_racha = st.file_uploader('Excel de rachas (opcional)',type=['xlsx'],key=f'rachas_{anio_evento}')
                    if nueva_lluvia is not None: carga_lluvia = nueva_lluvia
                    if nueva_racha is not None: carga_racha = nueva_racha
                    if st.button('Usar archivos de la carpeta',key=f'restablecer_archivos_{anio_evento}'):
                        carga_lluvia, carga_racha = None, None
                    configuracion.update(carpeta=carpeta_anio,lluvia=carga_lluvia,rachas=carga_racha)
                    st.button('Actualizar archivos',key='actualizar_eventos_multianio')
            fuentes, errores = fuentes_eventos_anio(carpeta_anio,anio_evento)
            for tipo, carga in [('lluvia',carga_lluvia),('rachas',carga_racha)]:
                if carga is not None:
                    nuevas, fallos=fuentes_eventos_anio('',anio_evento,[carga],tipo)
                    fuentes[tipo]=nuevas[tipo];errores.extend(fallos)
            if errores and acceso_configuracion:
                with ajustes_eventos:
                    st.caption('Archivos que requieren revisión')
                    for mensaje in errores:st.warning(mensaje)
            tablas={}
            for tipo in ['lluvia','rachas']:
                opciones=fuentes[tipo]
                if opciones:
                    etiquetas = [op[0] for op in opciones]
                    predeterminado = max(range(len(opciones)), key=lambda i:(len(opciones[i][1]), opciones[i][1].Fecha.max()))
                    etiqueta_guardada = configuracion.get('fuente_'+tipo)
                    indice = etiquetas.index(etiqueta_guardada) if etiqueta_guardada in etiquetas else predeterminado
                    if acceso_configuracion:
                        with ajustes_eventos:
                            indice=control_ruemag(st.selectbox, f'Archivo de {tipo} para {anio_evento}', range(len(opciones)), index=indice, format_func=lambda i, op=opciones: op[i][0], key=f'fuente_{tipo}_{anio_evento}') if len(opciones)>1 else 0
                            st.caption(f'{tipo.capitalize()}: {opciones[indice][0]}')
                            configuracion['fuente_'+tipo] = opciones[indice][0]
                    tablas[tipo]=opciones[indice][1]
                else: tablas[tipo]=pd.DataFrame(columns=['Evento','Fecha'])
            tabla_lluvia,tabla_rachas=tablas['lluvia'],tablas['rachas']
            eventos_disponibles=set(tabla_lluvia.Evento)|set(tabla_rachas.Evento)
            if base_reportes is not None:
                eventos_disponibles.update(base_reportes['Eventos'].loc[base_reportes['Eventos'].Anio==anio_evento,'Evento'])
            eventos_disponibles=sorted(int(n) for n in eventos_disponibles)
            calendario_eventos=fechas_eventos_disponibles(tabla_lluvia,tabla_rachas,base_reportes,anio_evento)
            numero_evento=None
            with columna_evento:
                modo_evento=control_ruemag(st.selectbox, '2. Buscar por', ['Número de evento', 'Fecha'], key='modo_consulta_eventos')
                if modo_evento=='Número de evento':
                    numero_evento=control_ruemag(st.selectbox, f'Evento de {anio_evento}', eventos_disponibles, index=len(eventos_disponibles) - 1, format_func=lambda n: f'{anio_evento} · Evento #{n}', key=f'evento_individual_{anio_evento}') if eventos_disponibles else None
                    if numero_evento is not None:
                        dias=sorted(calendario_eventos.get(numero_evento,set()))
                        st.caption('Fecha(s): '+', '.join(d.strftime('%d/%m/%Y') for d in dias) if dias else 'Este evento no tiene fecha registrada.')
                else:
                    dias_cargados=sorted({d for dias in calendario_eventos.values() for d in dias})
                    fecha_defecto=dias_cargados[-1] if dias_cargados else datetime.date(anio_evento,1,1)
                    fecha_consulta=control_ruemag(st.date_input, 'Fecha de consulta', value=fecha_defecto, min_value=datetime.date(anio_evento, 1, 1), max_value=datetime.date(anio_evento, 12, 31), format='DD/MM/YYYY', key=f'fecha_evento_{anio_evento}')
                    coincidentes=sorted(n for n,dias in calendario_eventos.items() if fecha_consulta in dias)
                    sin_fecha=[n for n in eventos_disponibles if not calendario_eventos.get(n)]
                    if coincidentes:
                        numero_evento=control_ruemag(st.selectbox, 'Evento en la fecha seleccionada', coincidentes, format_func=lambda n: f'{anio_evento} · Evento #{n}', key=f'evento_fecha_{anio_evento}_{fecha_consulta}')
                        if len(coincidentes)>1: st.caption(f'Hay {len(coincidentes)} eventos registrados ese día.')
                    elif eventos_disponibles:
                        if sin_fecha:
                            st.warning('No se encontró un evento para esa fecha. Hay eventos sin fecha registrada, por lo que no se puede confirmar su ausencia.')
                        else:
                            st.info('No hubo evento de tormenta')
                            st.caption(f'{fecha_consulta:%d/%m/%Y} · Sin evento registrado en los archivos cargados.')
                    if sin_fecha:
                        st.caption('Eventos sin fecha: '+', '.join(f'#{n}' for n in sin_fecha)+'. Se pueden consultar por número.')
                    st.caption('Los eventos con inicio y fin en días distintos también aparecen en los días que abarcan.')
            if not eventos_disponibles:
                st.info(f'No hay mediciones por estación ni reportes cargados para {anio_evento}. Configura los archivos del año desde la barra lateral, ingresando la contraseña.')

    puntos_lluvia,puntos_rachas=pd.DataFrame(),pd.DataFrame()
    fecha_lluvia,fecha_rachas=pd.NaT,pd.NaT
    if numero_evento is not None:
        puntos_lluvia,fecha_lluvia=convertir_evento_a_puntos(tabla_lluvia,numero_evento,datos,'Rain_Total')
        puntos_rachas,fecha_rachas=convertir_evento_a_puntos(tabla_rachas,numero_evento,datos,'Gust_Max')

    # --------------------------------------------------------
    # PESTAÑA 1: MAPAS Y ESTADÍSTICA POR EVENTO
    # --------------------------------------------------------
    with tab_mapas:
        if numero_evento is not None:
            tiene_reporte = base_reportes is not None and ((base_reportes['Eventos'].Anio == anio_evento) & (base_reportes['Eventos'].Evento == numero_evento)).any()
            if tiene_reporte:
                panel_mapas, panel_reporte = st.columns([3, 1.15], gap="large")
            else:
                panel_mapas = st.container()
                panel_reporte = st.container()
            with panel_reporte:
                mostrar_reporte_lateral(base_reportes, anio_evento, numero_evento)
            with panel_mapas:
                if puntos_lluvia.empty or puntos_rachas.empty:
                    st.warning("El evento seleccionado no tiene suficientes datos numéricos para generar ambos mapas.")
                else:

                    limite_eventos, _ = leer_limite(ruta_limite)
                    colonias_eventos, _ = leer_colonias(ruta_colonias)

                    fecha_titulo_lluvia = (fecha_lluvia.strftime("%d/%m/%Y") if not pd.isna(fecha_lluvia) else "fecha S/D")
                    fecha_titulo_rachas = (fecha_rachas.strftime("%d/%m/%Y") if not pd.isna(fecha_rachas) else "fecha S/D")
                    fecha_mapa_lluvia = (fecha_lluvia.date() if not pd.isna(fecha_lluvia) else pd.Timestamp.today().date())
                    fecha_mapa_rachas = (fecha_rachas.date() if not pd.isna(fecha_rachas) else pd.Timestamp.today().date())

                    with st.spinner(f"Generando mapas compactos en HD para el evento {numero_evento}..."):
                        imagen_evento_lluvia = mapa_ruemag_cache(contexto_cartografico_ruemag(), puntos_lluvia, 'Rain_Total', fecha_mapa_lluvia, fecha_mapa_lluvia, limite_eventos, colonias_eventos, titulo_personalizado=f'Evento {numero_evento} — Precipitación — {fecha_titulo_lluvia}', niveles_personalizados=np.arange(0, 71, 5), ticks_personalizados=np.arange(0, 71, 10), decimales=0, extension_personalizada='max', catalogo=catalogo_estaciones(datos), compacto=True)

                        imagen_evento_rachas = mapa_ruemag_cache(contexto_cartografico_ruemag(), puntos_rachas, 'Gust_Max', fecha_mapa_rachas, fecha_mapa_rachas, limite_eventos, colonias_eventos, titulo_personalizado=f'Evento {numero_evento} — Racha máx — {fecha_titulo_rachas}', niveles_personalizados=np.arange(0, 101, 5), colores_personalizados=['#f3eef8', '#d8c5e8', '#b79ad2', '#9467b5', '#713b96', '#4b176d'], ticks_personalizados=np.arange(0, 101, 10), decimales=0, extension_personalizada='max', catalogo=catalogo_estaciones(datos), compacto=True)

                    max_lluvia = puntos_lluvia.loc[puntos_lluvia["Rain_Total"].idxmax()]
                    max_racha = puntos_rachas.loc[puntos_rachas["Gust_Max"].idxmax()]

                    met_e1, met_e2, met_e3 = st.columns(3)
                    with met_e1: mostrar_tarjeta("Evento Activo", str(numero_evento), fecha_titulo_lluvia, "neutral")
                    with met_e2: mostrar_tarjeta("Máxima Precipitación", f"{max_lluvia['Rain_Total']:.0f} mm", max_lluvia["nombre_estacion"], "maximo")
                    with met_e3: mostrar_tarjeta("Máxima Racha", f"{max_racha['Gust_Max']:.0f} km/h", max_racha["nombre_estacion"], "maximo")

                    st.markdown("---")
                    # Estructura centrada de 3 columnas para que los mapas queden más estéticos y compactos
                    col_map1, col_map2 = st.columns(2, gap="small")
                    with col_map1:
                        st.markdown("### Precipitación")
                        st.image(imagen_evento_lluvia, use_column_width=True)
                    with col_map2:
                        st.markdown("### Rachas de Viento")
                        st.image(imagen_evento_rachas, use_column_width=True)

        if numero_evento is not None and not puntos_lluvia.empty and not puntos_rachas.empty:
            # --- INTEGRACIÓN DE LA ESTADÍSTICA DEL EVENTO ---
            st.markdown("---")
            st.markdown(f"## Estadística por Estación — Evento {numero_evento}")
            st.caption(f"Detalle de precipitación y rachas máximas registradas en cada estación para el evento {numero_evento}.")

            df_evento_detalle = puntos_lluvia[["nombre_estacion", "Rain_Total"]].merge(
                puntos_rachas[["nombre_estacion", "Gust_Max"]], on="nombre_estacion", how="outer"
            ).sort_values("Rain_Total", ascending=False).reset_index(drop=True)

            def clasificar_lluvia(p):
                if pd.isna(p): return "Sin Datos"
                elif p <= 18: return "Ligera (1–18 mm)"
                elif p <= 26: return "Moderada (18.1–26 mm)"
                elif p <= 42: return "Fuerte (26.1–42 mm)"
                else: return "Intensa (>42 mm)"

            df_evento_detalle["Categoría Lluvia"] = df_evento_detalle["Rain_Total"].apply(clasificar_lluvia)

            col_est1, col_est2 = st.columns([1.2, 1], gap="large")

            with col_est1:
                st.markdown("### Tabla de Registros por Estación")
                filas_ev_html = []
                for idx, row in df_evento_detalle.iterrows():
                    val_lluvia = f"{row['Rain_Total']:.1f} mm" if pd.notna(row['Rain_Total']) else "S/D"
                    val_racha = f"{row['Gust_Max']:.1f} km/h" if pd.notna(row['Gust_Max']) else "S/D"
                    cat = row['Categoría Lluvia']
                    filas_ev_html.append(f"<tr><td>{idx+1}</td><td>{escape(str(row['nombre_estacion']))}</td><td>{val_lluvia}</td><td>{val_racha}</td><td>{cat}</td></tr>")

                tabla_est_html = f"""<div class="contenedor-top" role="region" aria-label="Registros por estación" tabindex="0"><table class="tabla-top">
                    <thead><tr><th scope=col>#</th><th scope=col>Estación</th><th scope=col>Precipitación</th><th scope=col>Racha Máx.</th><th scope=col>Categoría</th></tr></thead>
                    <tbody>{''.join(filas_ev_html)}</tbody></table></div>"""
                st.markdown(tabla_est_html, unsafe_allow_html=True)

            with col_est2:
                st.markdown("### Gráfica de Barras por Estación")
                fig_ev_bar, ax_ev_bar = plt.subplots(figsize=(8, 6.5), dpi=400)
                ax_ev_bar.set_facecolor("white")
            
                nombres_est = df_evento_detalle["nombre_estacion"].tolist()
                valores_lluvia = df_evento_detalle["Rain_Total"].fillna(0).to_numpy(dtype=float)
            
                def color_barra_ev(p):
                    if p <= 18: return "#2ecc71"
                    elif p <= 26: return "#f1c40f"
                    elif p <= 42: return "#e67e22"
                    else: return "#e74c3c"

                colores_b = [color_barra_ev(p) for p in valores_lluvia]
                pos_x = np.arange(len(nombres_est))
            
                barras_ev = ax_ev_bar.bar(pos_x, valores_lluvia, color=colores_b, edgecolor="black", linewidth=0.4, width=0.7)
            
                ax_ev_bar.axhline(18, color="#f1c40f", linestyle="--", linewidth=1.2, label="Umbral 18 mm")
                ax_ev_bar.axhline(26, color="#e67e22", linestyle="--", linewidth=1.2, label="Umbral 26 mm")
                ax_ev_bar.axhline(42, color="#e74c3c", linestyle="--", linewidth=1.2, label="Umbral 42 mm")

                for bar in barras_ev:
                    h = bar.get_height()
                    if h > 0:
                        ax_ev_bar.text(
                            bar.get_x() + bar.get_width()/2., h + 0.8, f"{h:.1f}",
                            ha='center', va='bottom', fontsize=9, fontweight='bold', color='#17232a'
                        )

                ax_ev_bar.set_title(f"Precipitación por Estación — Evento {numero_evento}", fontsize=16, fontweight="bold", pad=15)
                ax_ev_bar.set_ylabel("Precipitación acumulada (mm)", fontsize=14, fontweight="bold")
                ax_ev_bar.set_xticks(pos_x)
                ax_ev_bar.set_xticklabels(nombres_est, rotation=35, ha='right', fontsize=11, fontweight='bold')
                ax_ev_bar.grid(axis="y", linestyle=":", alpha=0.6)
                ax_ev_bar.set_axisbelow(True)
            
                from matplotlib.patches import Patch
                leg_ev = [
                    Patch(facecolor='#2ecc71', edgecolor='black', label='Ligera (1–18 mm)'),
                    Patch(facecolor='#f1c40f', edgecolor='black', label='Moderada (18–26 mm)'),
                    Patch(facecolor='#e67e22', edgecolor='black', label='Fuerte (26–42 mm)'),
                    Patch(facecolor='#e74c3c', edgecolor='black', label='Intensa (>42 mm)')
                ]
                ax_ev_bar.legend(handles=leg_ev, loc='upper right', fontsize=10, frameon=True)

                fig_ev_bar.tight_layout()
            
                img_ev_io = io.BytesIO()
                fig_ev_bar.savefig(img_ev_io, format="png", dpi=400, bbox_inches="tight", facecolor="white")
                img_ev_io.seek(0)
                st.image(img_ev_io, use_column_width=True)
                plt.close(fig_ev_bar)

            st.markdown("---")
            with st.expander("Descargar mapas y datos del evento"):
                st.markdown(f'<p class="ru-export-contexto">Evento #{numero_evento} · {anio_evento} · Precipitación y rachas de viento</p>', unsafe_allow_html=True)
                descarga_ruemag('⬇ Descargar gráfica de lluvia por estación (PNG)', img_ev_io, f'RUEMAG_Estadistico_Evento_{numero_evento}.png', 'image/png', key=f'dl_stat_ev_{numero_evento}', titulo='Precipitación por estación', contexto=f'{anio_evento} · Evento #{numero_evento} · Gráfica de barras', formato='PNG')
                for etiqueta, puntos_evento, variable_evento, fecha_evento, imagen_evento in [
                    ("Precipitación", puntos_lluvia, "Rain_Total", fecha_mapa_lluvia, imagen_evento_lluvia),
                    ("Rachas de viento", puntos_rachas, "Gust_Max", fecha_mapa_rachas, imagen_evento_rachas),
                ]:
                    st.markdown(f"**{etiqueta} · Evento {numero_evento}**")
                    prefijo = f"RUEMAG_{anio_evento}_Evento_{numero_evento}_{variable_evento}_{fecha_evento}"
                    niveles_evento = np.arange(0, 71, 5) if variable_evento == "Rain_Total" else np.arange(0, 101, 5)
                    colores_evento = (VARIABLES["Rain_Total"]["colores"] if variable_evento == "Rain_Total" else ["#f3eef8", "#d8c5e8", "#b79ad2", "#9467b5", "#713b96", "#4b176d"])
                    catalogo_evento = catalogo_estaciones(datos)
                    tabla_descarga = catalogo_evento.merge(puntos_evento[["estacion", variable_evento]], on="estacion", how="left")
                    tabla_descarga.insert(0, "Evento", numero_evento)
                    tabla_descarga.insert(1, "Fecha", fecha_evento)
                    tabla_descarga["Unidad"] = VARIABLES[variable_evento]["unidad"]
                    csv_evento = tabla_descarga.to_csv(index=False, na_rep="S/D").encode("utf-8-sig")
                    kmz_evento, shp_evento = None, None
                    try: kmz_evento = kmz_ruemag_cache(contexto_cartografico_ruemag(), puntos_evento, variable_evento, fecha_evento, fecha_evento, limite_eventos, colonias_eventos, catalogo=catalogo_evento, niveles_personalizados=niveles_evento, colores_personalizados=colores_evento, extension_personalizada='max')
                    except Exception: pass
                    try: shp_evento = shp_ruemag_cache(contexto_cartografico_ruemag(), puntos_evento, variable_evento, limite_eventos, niveles_personalizados=niveles_evento, colores_personalizados=colores_evento)
                    except Exception: pass

                    cd1, cd2, cd3, cd4 = st.columns(4, gap="small")
                    with cd1: descarga_ruemag('⬇ Descargar mapa (PNG)', imagen_evento, f'{prefijo}.png', 'image/png', key=f'{prefijo}_png', titulo=f'Mapa de {etiqueta.lower()}', contexto=f'{anio_evento} · Evento #{numero_evento} · {fecha_evento:%d/%m/%Y}', formato='PNG')
                    with cd2: descarga_ruemag('⬇ Descargar mediciones (CSV)', csv_evento, f'{prefijo}.csv', 'text/csv', key=f'{prefijo}_csv', titulo=f'Mediciones de {etiqueta.lower()}', contexto=f'{anio_evento} · Evento #{numero_evento} · {fecha_evento:%d/%m/%Y}', formato='CSV')
                    with cd3: descarga_ruemag('⬇ Descargar capa IDW (SHP)', shp_evento if shp_evento is not None else b'', f'{prefijo}.zip', 'application/zip', disabled=shp_evento is None, key=f'{prefijo}_shp', titulo=f'Capa IDW de {etiqueta.lower()}', contexto=f'{anio_evento} · Evento #{numero_evento} · {fecha_evento:%d/%m/%Y} · Archivo ZIP', formato='SHP')
                    with cd4: descarga_ruemag('⬇ Descargar mapa para Google Earth (KMZ)', kmz_evento if kmz_evento is not None else b'', f'{prefijo}.kmz', 'application/vnd.google-earth.kmz', disabled=kmz_evento is None, key=f'{prefijo}_kmz', titulo=f'Mapa para Google Earth de {etiqueta.lower()}', contexto=f'{anio_evento} · Evento #{numero_evento} · {fecha_evento:%d/%m/%Y}', formato='KMZ')

    # --------------------------------------------------------
    # PESTAÑA 2: HISTÓRICO GENERAL DE TEMPORALES (DIVIDIDO EN PARTES)
    # --------------------------------------------------------
    with tab_historico:
        st.markdown("## Histórico de Temporales")

        if not Path(RUTA_HISTORICO_EVENTOS).exists():
            st.info(f"No se encontró el archivo general de eventos en la ruta: {RUTA_HISTORICO_EVENTOS}")
        else:
            try:
                dfs_temporales = leer_historico_temporales(RUTA_HISTORICO_EVENTOS)
                anios_historicos = sorted(dfs_temporales, key=int)
                periodo_historico = f"{anios_historicos[0]} - {anios_historicos[-1]}"
                opcion_general = f"General (Todos los años: {periodo_historico})"
                opciones_anio = [opcion_general] + list(reversed(anios_historicos))
                sel_anio = control_ruemag(st.selectbox, 'Selecciona el temporal o año a analizar:', opciones_anio, key='filtro_anio_historico')

                if sel_anio == opcion_general:
                    df_fil = pd.concat([dfs_temporales[a] for a in anios_historicos], ignore_index=True)
                    df_fil = df_fil.sort_values(['Fecha', 'Evento']).reset_index(drop=True)
                    texto_temporal = f"Histórico general ({periodo_historico})"
                else:
                    df_fil = dfs_temporales[sel_anio].copy().reset_index(drop=True)
                    texto_temporal = f"Temporal {sel_anio}"
                st.caption('2023 se incluye únicamente en las estadísticas históricas. Las estaciones difieren de la red actual; la comparación entre años debe considerar ese cambio de cobertura.')

                def calcular_duracion(row):
                    inicio, final = row.get('inicio'), row.get('final')
                    if pd.isna(inicio) or pd.isna(final): return pd.NA
                    if isinstance(inicio, str): inicio = pd.to_datetime(inicio, errors='coerce')
                    if isinstance(final, str): final = pd.to_datetime(final, errors='coerce')
                    
                    inicio_time = inicio.time() if isinstance(inicio, datetime.datetime) else inicio
                    final_time = final.time() if isinstance(final, datetime.datetime) else final
                    if not isinstance(inicio_time, datetime.time) or not isinstance(final_time, datetime.time): return pd.NA
                    
                    start_sec = inicio_time.hour * 3600 + inicio_time.minute * 60 + inicio_time.second
                    end_sec = final_time.hour * 3600 + final_time.minute * 60 + final_time.second
                    if end_sec < start_sec: end_sec += 24 * 3600
                    return (end_sec - start_sec) / 3600

                df_fil['duracion_horas'] = pd.to_numeric(df_fil.apply(calcular_duracion, axis=1), errors='coerce')
                
                def categorizar_precip(p):
                    if pd.isna(p) or p <= 18: return "Ligera"
                    elif p <= 26: return "Moderada"
                    elif p <= 42: return "Fuerte"
                    else: return "Intensa"

                precips_serie = df_fil.get('precip', pd.Series([10]*len(df_fil)))
                df_fil['Cat_Precip'] = precips_serie.apply(categorizar_precip)
                
                dur_promedio = df_fil['duracion_horas'].mean()
                total_ev = len(df_fil)
                max_viento = df_fil['viento'].max() if 'viento' in df_fil.columns else 0.0
                
                mh1, mh2, mh3 = st.columns(3)
                with mh1: mostrar_tarjeta("Total de Eventos", str(total_ev), texto_temporal, "neutral")
                with mh2: mostrar_tarjeta("Duración Promedio", f"{dur_promedio:.1f} hrs", "Por evento", "neutral")
                with mh3: mostrar_tarjeta("Racha Máxima Registrada", f"{max_viento:.1f} km/h", "Del periodo", "maximo")
                
                st.markdown("---")
                
                conteo_hist = df_fil['Cat_Precip'].value_counts().reindex(["Ligera", "Moderada", "Fuerte", "Intensa"], fill_value=0)
                total_h = conteo_hist.sum()
                
                col_hp1, col_hp2 = st.columns(2, gap="large")
                
                with col_hp1:
                    st.markdown("### Categoría de Precipitación")
                    fig_hi, ax_hi = plt.subplots(figsize=(7, 7), dpi=400)
                    ax_hi.set_facecolor("white")
                    colores_past = ["#2ecc71", "#f1c40f", "#e67e22", "#e74c3c"]
                    
                    def autopct_format_hist(pct):
                        val = int(round(pct * total_h / 100.0))
                        return f"{val} eventos\n({pct:.1f}%)" if val > 0 else ""

                    wedges, texts, autotexts = ax_hi.pie(
                        conteo_hist, labels=conteo_hist.index, colors=colores_past,
                        autopct=autopct_format_hist, startangle=90, textprops=dict(color="black", weight="bold", fontsize=11),
                        wedgeprops=dict(width=0.48, edgecolor='white', linewidth=2)
                    )
                    for at in autotexts:
                        at.set_fontsize(10)
                        at.set_color('black')
                    ax_hi.set_title(f"Categoría de los {total_ev} eventos\n{texto_temporal}", fontsize=16, fontweight="bold", pad=15)
                    fig_hi.tight_layout()
                    
                    img_hi_io = io.BytesIO()
                    fig_hi.savefig(img_hi_io, format="png", dpi=400, bbox_inches="tight", facecolor="white")
                    img_hi_io.seek(0)
                    st.image(img_hi_io, use_column_width=True)
                    descarga_ruemag('⬇ Descargar gráfica de categorías (PNG)', img_hi_io, f'RUEMAG_Hist_CatPrecip_{sel_anio}.png', 'image/png', key='dl_hist_cat', titulo='Categorías de precipitación', contexto=f'{texto_temporal} · Distribución de los eventos por intensidad', formato='PNG')
                    plt.close(fig_hi)
                    
                with col_hp2:
                    st.markdown("### Duración de los Eventos")
                    bins = [-1, 1, 2, 3, 24]
                    labels = ["Menos de 1 h", "1 a 2 h", "2 a 3 h", "Más de 3 h"]
                    df_fil['Rango_Duracion'] = pd.cut(df_fil['duracion_horas'], bins=bins, labels=labels)
                    conteo_dur = df_fil['Rango_Duracion'].value_counts().reindex(labels, fill_value=0)
                    total_dur = conteo_dur.sum()
                    
                    fig_hd, ax_hd = plt.subplots(figsize=(7, 7), dpi=400)
                    ax_hd.set_facecolor("white")
                    
                    def autopct_format_dur(pct):
                        val = int(round(pct * total_dur / 100.0))
                        return f"{val} eventos\n({pct:.1f}%)" if val > 0 else ""

                    wedges, texts, autotexts = ax_hd.pie(
                        conteo_dur, labels=conteo_dur.index, colors=["#bdc3c7", "#2ecc71", "#f1c40f", "#e74c3c"],
                        autopct=autopct_format_dur, startangle=90, textprops=dict(color="black", weight="bold", fontsize=11),
                        wedgeprops=dict(width=0.48, edgecolor='white', linewidth=2)
                    )
                    for at in autotexts:
                        at.set_fontsize(10)
                        at.set_color('black')
                    ax_hd.set_title(f"Duración de los {total_ev} eventos\n{texto_temporal}", fontsize=16, fontweight="bold", pad=15)
                    fig_hd.tight_layout()
                    
                    img_hd_io = io.BytesIO()
                    fig_hd.savefig(img_hd_io, format="png", dpi=400, bbox_inches="tight", facecolor="white")
                    img_hd_io.seek(0)
                    st.image(img_hd_io, use_column_width=True)
                    descarga_ruemag('⬇ Descargar gráfica de duración (PNG)', img_hd_io, f'RUEMAG_Hist_Duracion_{sel_anio}.png', 'image/png', key='dl_hist_dur', titulo='Duración de los eventos', contexto=f'{texto_temporal} · Distribución de los eventos por duración', formato='PNG')
                    plt.close(fig_hd)

                # --- GRÁFICA DE BARRAS DE PRECIPITACIÓN (DIVIDIDA EN PARTES / BLOQUES) ---
                st.markdown("---")
                st.markdown("### Estadística de Precipitación Máxima por Evento")
                
                numeracion_barras = control_ruemag(st.radio, 'Numeración debajo de las barras', ['Por año', 'Consecutivo completo', 'Ambas'], index=2, horizontal=True, key='numeracion_barras_historico')

                tamanio_bloque = 100
                total_bloques = math.ceil(total_ev / tamanio_bloque)
                
                opciones_bloques = [f"Bloque {i+1}: Eventos {i*tamanio_bloque + 1} al {min((i+1)*tamanio_bloque, total_ev)}" for i in range(total_bloques)]
                opciones_bloques.insert(0, "Ver todos los eventos juntos (Completa)")
                
                seleccion_bloque = control_ruemag(st.selectbox, 'Selecciona la sección o bloque a visualizar a detalle:', opciones_bloques, key='selector_bloque_barras')
                
                if "Ver todos" in seleccion_bloque:
                    df_plot = df_fil.copy()
                    inicio_idx = 1
                else:
                    idx_bloque = opciones_bloques.index(seleccion_bloque) - 1
                    inicio_idx = idx_bloque * tamanio_bloque
                    fin_idx = min(inicio_idx + tamanio_bloque, total_ev)
                    df_plot = df_fil.iloc[inicio_idx:fin_idx].copy()
                    inicio_idx += 1

                num_eventos_plot = len(df_plot)
                ancho_figura = max(14, num_eventos_plot * 0.18)
                
                fig_bar, ax_bar = plt.subplots(figsize=(ancho_figura, 7.5), dpi=400)
                ax_bar.set_facecolor("white")
                
                precips_plot = df_plot.get('precip', pd.Series([10]*num_eventos_plot)).to_numpy(dtype=float)
                
                col_zona_precip = None
                for c in df_fil.columns:
                    if "zona" in c.lower() and "precip" in c.lower():
                        col_zona_precip = c
                        break
                if col_zona_precip is None:
                    col_zona_precip = "zona precip" if "zona precip" in df_fil.columns else df_fil.columns[1]

                estaciones_bar = df_plot.get(col_zona_precip, pd.Series(["Estación"]*num_eventos_plot)).astype(str).tolist()
                x_vals = np.arange(inicio_idx, inicio_idx + num_eventos_plot)
                # Mantener la numeración consecutiva y agrupar visualmente por año.
                etiquetas_eventos = (
                    df_plot['Consecutivo'] if numeracion_barras == "Consecutivo completo"
                    else df_plot['Evento']
                ).astype(int).to_numpy()
                
                def obtener_color_barra(p):
                    if pd.isna(p) or p <= 18: return "#2ecc71"
                    elif p <= 26: return "#f1c40f"
                    elif p <= 42: return "#e67e22"
                    else: return "#e74c3c"

                colores_barras = [obtener_color_barra(p) for p in precips_plot]
                bars = ax_bar.bar(x_vals, precips_plot, color=colores_barras, edgecolor="black", linewidth=0.3, width=0.8)
                
                ax_bar.axhline(18, color="#f1c40f", linestyle="--", linewidth=1.5, label="Umbral 18 mm")
                ax_bar.axhline(26, color="#e67e22", linestyle="--", linewidth=1.5, label="Umbral 26 mm")
                ax_bar.axhline(42, color="#e74c3c", linestyle="--", linewidth=1.5, label="Umbral 42 mm")
                
                for i, bar in enumerate(bars):
                    h = bar.get_height()
                    if pd.notna(h):
                        label_texto = f"{estaciones_bar[i]} - {precips_plot[i]:.1f} mm"
                        if h >= 16:
                            ax_bar.text(
                                bar.get_x() + bar.get_width() / 2., h - 1.5, label_texto,
                                ha='center', va='top', rotation=90, fontsize=8, fontweight='bold', color='white'
                            )
                        else:
                            ax_bar.text(
                                bar.get_x() + bar.get_width() / 2., h + 0.6, label_texto,
                                ha='center', va='bottom', rotation=90, fontsize=8, fontweight='bold', color='black'
                            )

                ax_bar.set_title(f"Precipitación máxima registrada por evento\n{texto_temporal} ({seleccion_bloque})", fontsize=18, fontweight="bold", pad=15)
                etiqueta_eje = {
                    "Por año": "Número de evento dentro del año",
                    "Consecutivo completo": "Consecutivo de todo el histórico",
                    "Ambas": "Fila superior: evento del año · Fila inferior: consecutivo histórico"
                }[numeracion_barras]
                ax_bar.set_xlabel(etiqueta_eje, fontsize=13, fontweight="bold",
                                  labelpad=90 if numeracion_barras == "Ambas" else 57)
                ax_bar.set_ylabel("Precipitación acumulada (mm)", fontsize=16, fontweight="bold")
                
                ax_bar.set_xlim(inicio_idx - 0.5, inicio_idx + num_eventos_plot - 0.5)
                ax_bar.set_ylim(0, max(np.nanmax(precips_plot) + 15, 75))
                
                ax_bar.set_xticks(x_vals)
                ax_bar.set_xticklabels(etiquetas_eventos, fontsize=9, rotation=90, fontweight='bold')

                if numeracion_barras == "Ambas":
                    for posicion, consecutivo in zip(x_vals, df_plot['Consecutivo']):
                        ax_bar.annotate(
                            str(int(consecutivo)), xy=(posicion, 0),
                            xycoords=ax_bar.get_xaxis_transform(),
                            xytext=(0, -36), textcoords='offset points',
                            ha='center', va='top', rotation=90, fontsize=9,
                            color='#455a64', annotation_clip=False
                        )
                anios_barras = df_plot['Anio'].astype(int).to_numpy()
                cortes_anio = np.r_[0, np.flatnonzero(anios_barras[1:] != anios_barras[:-1]) + 1, num_eventos_plot]
                for j in range(len(cortes_anio) - 1):
                    desde, hasta = cortes_anio[j], cortes_anio[j + 1]
                    izquierda = x_vals[desde] - 0.5
                    derecha = x_vals[hasta - 1] + 0.5
                    centro = (izquierda + derecha) / 2
                    ax_bar.annotate(
                        str(anios_barras[desde]), xy=(centro, 0),
                        xycoords=ax_bar.get_xaxis_transform(),
                        xytext=(0, -68 if numeracion_barras == "Ambas" else -43), textcoords='offset points',
                        ha='center', va='top', fontsize=12, fontweight='bold',
                        annotation_clip=False
                    )
                    if j > 0:
                        ax_bar.axvline(izquierda, color='#53666b', linewidth=1.0, linestyle=':', alpha=0.7)

                
                ax_bar.grid(axis="y", linestyle="--", alpha=0.5)
                ax_bar.set_axisbelow(True)
                
                from matplotlib.patches import Patch
                elementos_leyenda = [
                    Patch(facecolor='#2ecc71', edgecolor='black', label='Ligera: 1–18 mm'),
                    Patch(facecolor='#f1c40f', edgecolor='black', label='Moderada: 18.1–26 mm'),
                    Patch(facecolor='#e67e22', edgecolor='black', label='Fuerte: 26.1–42 mm'),
                    Patch(facecolor='#e74c3c', edgecolor='black', label='Intensa: >42 mm')
                ]
                ax_bar.legend(handles=elementos_leyenda, loc='upper left', fontsize=11, frameon=True)
                
                fig_bar.tight_layout()
                
                img_bar_io = io.BytesIO()
                fig_bar.savefig(img_bar_io, format="png", dpi=400, bbox_inches="tight", facecolor="white")
                img_bar_io.seek(0)
                st.image(img_bar_io, use_column_width=True)
                descarga_ruemag('⬇ Descargar barras de precipitación (PNG)', img_bar_io, f'RUEMAG_Barras_Precip_{sel_anio}.png', 'image/png', key='dl_hist_barras', titulo='Precipitación máxima por evento', contexto=f'{texto_temporal} · {seleccion_bloque}', formato='PNG')
                plt.close(fig_bar)
                
            except Exception as error:
                st.error("Error al procesar el archivo general de eventos.")
                st.exception(error)

    st.stop()


# ============================================================
# SECCIÓN DE ANOMALÍAS
# ============================================================

if seccion_pagina == "📊 Anomalías":

    titulo_vista_ruemag("Anomalías", "Compara las observaciones con la referencia histórica para las fechas elegidas.")

    variables_anomalia = [clave for clave in ["T_Max", "T_Min", "Rain_Total", "Gust_Max"] if clave in datos.columns]
    fecha_minima_a = datos["Day"].min().date()
    fecha_maxima_a = datos["Day"].max().date()

    sel_a1, sel_a2, sel_a3, sel_a4 = st.columns([1.4, 1, 1, 1.1])
    with sel_a1: variable_a = control_ruemag(st.selectbox, 'Variable de anomalía', variables_anomalia, format_func=lambda clave: VARIABLES[clave]['nombre'], key='variable_anomalia')
    with sel_a2: fecha_inicial_a = control_ruemag(st.date_input, 'Fecha inicial', fecha_maxima_a, min_value=fecha_minima_a, max_value=fecha_maxima_a, key='fecha_inicial_anomalia', format='DD/MM/YYYY', on_change=marcar_periodo_personalizado, args=('periodo_rapido_anomalia',))
    with sel_a3: fecha_final_a = control_ruemag(st.date_input, 'Fecha final', fecha_maxima_a, min_value=fecha_minima_a, max_value=fecha_maxima_a, key='fecha_final_anomalia', format='DD/MM/YYYY', on_change=marcar_periodo_personalizado, args=('periodo_rapido_anomalia',))

    with sel_a4:
        mostrar_periodos_ruemag('periodo_rapido_anomalia', 'fecha_inicial_anomalia', 'fecha_final_anomalia', fecha_minima_a, fecha_maxima_a)

    if fecha_inicial_a > fecha_final_a:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
        st.stop()

    try:
        clima_a = None
        rachas_a = None

        if variable_a == "Gust_Max":
            if archivo_rachas_historicas is not None: fuente_rachas_a = archivo_rachas_historicas.getvalue()
            else:
                fuente_rachas_a = ruta_rachas_historicas
                if not Path(ruta_rachas_historicas).exists(): raise FileNotFoundError(f"No se encontró el Excel histórico de rachas: {ruta_rachas_historicas}")
            rachas_a = leer_rachas_historicas(fuente_rachas_a)
        else:
            if archivo_climatologia is not None: fuente_clima_a = archivo_climatologia.getvalue()
            else:
                fuente_clima_a = ruta_climatologia
                if not Path(ruta_climatologia).exists(): raise FileNotFoundError(f"No se encontró el CSV climatológico: {ruta_climatologia}")
            clima_a = leer_climatologia(fuente_clima_a)

        referencia_a, periodo_base_a, anios_validos_a = calcular_referencia_anomalia(variable_a, fecha_inicial_a, fecha_final_a, clima_a, rachas_a)
    except Exception as error:
        st.error("No fue posible calcular la climatología.")
        st.exception(error)
        st.stop()

    puntos_observados_a = resumir_periodo(datos, variable_a, fecha_inicial_a, fecha_final_a)
    if puntos_observados_a.empty:
        st.warning("No hay datos RUEMAG para la variable y fechas elegidas.")
        st.stop()

    clave_a = f"ANOM_{variable_a}"
    unidad_a = VARIABLES[variable_a]["unidad"]

    VARIABLES[clave_a] = {
        "nombre": f"Anomalía de {VARIABLES[variable_a]['nombre'].lower()}", "unidad": unidad_a,
        "colores": ["#214f9c", "#78a9dc", "#f7f7f7", "#ef9a8d", "#a71930"]
    }

    if variable_a == "Rain_Total":
        VARIABLES[clave_a]["colores"] = ["#4a2600", "#7a4808", "#a86c1c", "#d1a35a", "#efe0bd", "#f7f7f2", "#d7efeb", "#9cd6d0", "#55b3ac", "#1e827f", "#004e4d"]
    elif variable_a == "Gust_Max":
        VARIABLES[clave_a]["colores"] = ["#352060", "#4b3480", "#66509c", "#8875b5", "#b5a8d2", "#f7f5fa", "#dec8eb", "#c394d7", "#a45abd", "#812d9c", "#5c176f"]

    puntos_a = puntos_observados_a.copy()
    puntos_a[clave_a] = puntos_a[variable_a] - referencia_a

    st.info(f"Periodo base: {periodo_base_a} · Referencia de las fechas elegidas: {referencia_a:.1f} {unidad_a} · Años con información: {anios_validos_a}")

    valores_a = puntos_a[clave_a]
    met_a1, met_a2, met_a3, met_a4 = st.columns(4)

    with met_a1: mostrar_tarjeta("Anomalía promedio", f"{valores_a.mean():+.1f} {unidad_a}", "Red RUEMAG", "neutral")
    with met_a2: mostrar_tarjeta("Anomalía máxima", f"{valores_a.max():+.1f} {unidad_a}", puntos_a.loc[valores_a.idxmax(), "nombre_estacion"], "maximo")
    with met_a3: mostrar_tarjeta("Anomalía mínima", f"{valores_a.min():+.1f} {unidad_a}", puntos_a.loc[valores_a.idxmin(), "nombre_estacion"], "minimo")
    with met_a4: mostrar_tarjeta("Estaciones con datos", str(len(puntos_a)), "Estaciones disponibles", "neutral")

    limite_a, _ = leer_limite(ruta_limite)
    colonias_a, _ = leer_colonias(ruta_colonias)

    try:
        with st.spinner("Generando mapa de anomalías…"):
            imagen_a = crear_mapa(puntos_a, clave_a, fecha_inicial_a, fecha_final_a, limite_a, colonias_a, catalogo=catalogo_estaciones(datos))
    except Exception as error:
        st.error("No se pudo generar el mapa de anomalías para esta consulta.")
        with st.expander("Ver el detalle del error del mapa"):
            st.exception(error)
        st.stop()

    mapa_a, tabla_a = st.columns([2.2, 1], gap="large")
    with mapa_a:
        st.image(imagen_a, use_column_width=True, output_format='PNG')
    with tabla_a:
        st.markdown("### Anomalías por estación")
        st.markdown(crear_tabla_top(puntos_a, clave_a, unidad_a), unsafe_allow_html=True)
        mostrar_estacion_ruemag(puntos_a, catalogo_estaciones(datos), clave_a, fecha_inicial_a, fecha_final_a, 'detalle_estacion_anomalia')

    with st.expander("Descargar mapa y datos"):
        st.markdown('<p class="ru-export-contexto">' + escape(VARIABLES[variable_a]['nombre']) + f' · {fecha_inicial_a:%d/%m/%Y} al {fecha_final_a:%d/%m/%Y}</p>', unsafe_allow_html=True)
        nombre_a = f"RUEMAG_ANOMALIA_{variable_a}_{fecha_inicial_a}_{fecha_final_a}"
        kmz_a = kmz_ruemag_cache(contexto_cartografico_ruemag(), puntos_a, clave_a, fecha_inicial_a, fecha_final_a, limite_a, colonias_a, catalogo=catalogo_estaciones(datos))
        csv_a = crear_csv_anomalias(datos, puntos_a, variable_a, clave_a, referencia_a, fecha_inicial_a, fecha_final_a, periodo_base_a)

        shp_a, error_shp_a = None, None
        try: shp_a = shp_ruemag_cache(contexto_cartografico_ruemag(), puntos_a, clave_a, limite_a)
        except Exception as error: error_shp_a = str(error)

        if error_shp_a:
            st.warning('La exportación SHP no está disponible para esta consulta. Los demás formatos se conservan.')
            with st.expander('Detalle de la exportación SHP'):
                st.code(error_shp_a)
        des_a1, des_a2, des_a3, des_a4 = st.columns(4, gap="small")
        with des_a1: descarga_ruemag('⬇ Descargar mapa de anomalías (PNG)', imagen_a, f'{nombre_a}.png', 'image/png', titulo='Mapa de anomalías', contexto=f'{VARIABLES[variable_a]['nombre']} · {fecha_inicial_a:%d/%m/%Y} al {fecha_final_a:%d/%m/%Y}', formato='PNG')
        with des_a2: descarga_ruemag('⬇ Descargar anomalías por estación (CSV)', csv_a, f'{nombre_a}.csv', 'text/csv', titulo='Anomalías y referencia por estación', contexto=f'{VARIABLES[variable_a]['nombre']} · {fecha_inicial_a:%d/%m/%Y} al {fecha_final_a:%d/%m/%Y}', formato='CSV')
        with des_a3: descarga_ruemag('⬇ Descargar capa de anomalías (SHP)', shp_a if shp_a is not None else b'', f'{nombre_a}_IDW.zip', 'application/zip', disabled=shp_a is None, help='Descarga ZIP con SHP, SHX, DBF y PRJ.', titulo='Capa IDW de anomalías', contexto=f'{VARIABLES[variable_a]['nombre']} · {fecha_inicial_a:%d/%m/%Y} al {fecha_final_a:%d/%m/%Y}' + ' · Archivo ZIP', formato='SHP')
        with des_a4: descarga_ruemag('⬇ Descargar anomalías para Google Earth (KMZ)', kmz_a, f'{nombre_a}.kmz', 'application/vnd.google-earth.kmz', titulo='Anomalías para Google Earth', contexto=f'{VARIABLES[variable_a]['nombre']} · {fecha_inicial_a:%d/%m/%Y} al {fecha_final_a:%d/%m/%Y}', formato='KMZ')

    st.stop()


# ============================================================
# SELECTORES (SECCIÓN PRINCIPAL)
# ============================================================

variables_disponibles = [variable for variable in VARIABLES if variable in datos.columns]
if not variables_disponibles:
    st.error("No se encontraron variables compatibles.")
    st.stop()

fecha_minima = datos["Day"].min().date()
fecha_maxima = datos["Day"].max().date()

titulo_vista_ruemag("Mapas meteorológicos", "Consulta la distribución espacial y la evolución de las mediciones de la red.")
selector1, selector2, selector3, selector4 = st.columns([1.4, 1, 1, 1.1])

with selector1: variable = control_ruemag(st.selectbox, 'Variable', variables_disponibles, format_func=lambda clave: VARIABLES[clave]['nombre'], key='variable_meteorologica')
with selector2: fecha_inicial = control_ruemag(st.date_input, 'Fecha inicial', fecha_maxima, min_value=fecha_minima, max_value=fecha_maxima, key='fecha_inicial_meteorologica', format='DD/MM/YYYY', on_change=marcar_periodo_personalizado, args=('periodo_rapido_meteorologico',))
with selector3: fecha_final = control_ruemag(st.date_input, 'Fecha final', fecha_maxima, min_value=fecha_minima, max_value=fecha_maxima, key='fecha_final_meteorologica', format='DD/MM/YYYY', on_change=marcar_periodo_personalizado, args=('periodo_rapido_meteorologico',))

with selector4:
    mostrar_periodos_ruemag('periodo_rapido_meteorologico', 'fecha_inicial_meteorologica', 'fecha_final_meteorologica', fecha_minima, fecha_maxima)

if fecha_inicial > fecha_final:
    st.error("La fecha inicial no puede ser posterior a la fecha final.")
    st.stop()


# ============================================================
# PROCESAR DATOS
# ============================================================

puntos = resumir_periodo(datos, variable, fecha_inicial, fecha_final)

if puntos.empty:
    st.warning("No hay datos para la variable y las fechas seleccionadas.")
    st.stop()


limite, limite_oficial = leer_limite(ruta_limite)
colonias, colonias_disponibles = leer_colonias(ruta_colonias)

valores = puntos[variable]
unidad = VARIABLES[variable]["unidad"]
estacion_maxima = puntos.loc[valores.idxmax(), "nombre_estacion"]
estacion_minima = puntos.loc[valores.idxmin(), "nombre_estacion"]


# ============================================================
# TARJETAS SUPERIORES
# ============================================================

metrica1, metrica2, metrica3, metrica4 = st.columns(4)

with metrica1: mostrar_tarjeta(titulo="Promedio", valor=f"{valores.mean():.1f} {unidad}", texto_inferior="Red RUEMAG", tipo="neutral")
with metrica2: mostrar_tarjeta(titulo="Máximo", valor=f"{valores.max():.1f} {unidad}", texto_inferior=estacion_maxima, tipo="maximo")
with metrica3: mostrar_tarjeta(titulo="Mínimo", valor=f"{valores.min():.1f} {unidad}", texto_inferior=estacion_minima, tipo="minimo")
with metrica4: mostrar_tarjeta(titulo="Estaciones con datos", valor=str(len(puntos)), texto_inferior="Estaciones disponibles", tipo="neutral")


# ============================================================
# GENERAR MAPA PRINCIPAL
# ============================================================

with st.spinner("Generando mapa IDW en máxima calidad..."):
    imagen_png = mapa_ruemag_cache(contexto_cartografico_ruemag(), puntos, variable, fecha_inicial, fecha_final, limite, colonias, catalogo=catalogo_estaciones(datos))


# ============================================================
# MAPA Y TOP 12
# ============================================================

columna_mapa, columna_top = st.columns([2.2, 1], gap="large")

with columna_mapa:
    st.image(imagen_png, use_column_width=True)
    if not limite_oficial: st.warning("Se utilizó un límite provisional. Verifica la ruta del archivo GPKG.")
    if not colonias_disponibles: st.warning("El mapa se generó sin los límites de colonias. Verifica la ruta indicada en la barra lateral.")

with columna_top:
    st.markdown("### Top 12 de estaciones")
    tabla_top = crear_tabla_top(puntos, variable, unidad)
    st.markdown(tabla_top, unsafe_allow_html=True)
    mostrar_estacion_ruemag(puntos, catalogo_estaciones(datos), variable, fecha_inicial, fecha_final, 'detalle_estacion_meteorologica')




# ============================================================
# DESCARGAS
# ============================================================

with st.expander("Descargar mapa y datos"):
    st.markdown('<p class="ru-export-contexto">' + escape(VARIABLES[variable]['nombre']) + f' · {fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}</p>', unsafe_allow_html=True)
    nombre_salida = f"RUEMAG_{variable}_{fecha_inicial}_{fecha_final}"
    archivo_kmz = kmz_ruemag_cache(contexto_cartografico_ruemag(), puntos, variable, fecha_inicial, fecha_final, limite, colonias, catalogo=catalogo_estaciones(datos))
    archivo_csv = crear_csv_periodo(datos, variable, fecha_inicial, fecha_final)

    archivo_shp, error_shp = None, None
    try: archivo_shp = shp_ruemag_cache(contexto_cartografico_ruemag(), puntos, variable, limite)
    except Exception as error: error_shp = str(error)

    if error_shp:
        st.warning('La exportación SHP no está disponible para esta consulta. Los demás formatos se conservan.')
        with st.expander('Detalle de la exportación SHP'):
            st.code(error_shp)
    boton1, boton2, boton3, boton4 = st.columns(4, gap="small")
    with boton1: descarga_ruemag(label='⬇ Descargar mapa (PNG)', data=imagen_png, file_name=f'{nombre_salida}.png', mime='image/png', titulo='Mapa meteorológico', contexto=f'{VARIABLES[variable]['nombre']} · {fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}', formato='PNG')
    with boton2: descarga_ruemag(label='⬇ Descargar mediciones por estación (CSV)', data=archivo_csv, file_name=f'RUEMAG_DATOS_{variable}_{fecha_inicial}_{fecha_final}.csv', mime='text/csv', titulo='Mediciones por fecha y estación', contexto=f'{VARIABLES[variable]['nombre']} · {fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}', formato='CSV')
    with boton3: descarga_ruemag(label='⬇ Descargar capa IDW (SHP)', data=archivo_shp if archivo_shp is not None else b'', file_name=f'{nombre_salida}_IDW.zip', mime='application/zip', disabled=archivo_shp is None, help='Descarga ZIP con SHP, SHX, DBF y PRJ.', titulo='Capa interpolada IDW', contexto=f'{VARIABLES[variable]['nombre']} · {fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}' + ' · Archivo ZIP', formato='SHP')
    with boton4: descarga_ruemag(label='⬇ Descargar mapa para Google Earth (KMZ)', data=archivo_kmz, file_name=f'{nombre_salida}.kmz', mime='application/vnd.google-earth.kmz', titulo='Mapa para Google Earth', contexto=f'{VARIABLES[variable]['nombre']} · {fecha_inicial:%d/%m/%Y} al {fecha_final:%d/%m/%Y}', formato='KMZ')


mostrar_grafica_periodo(datos, variable, fecha_inicial, fecha_final)
