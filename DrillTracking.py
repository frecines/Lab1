# -*- coding: utf-8 -*-
"""
Created on Sun Feb 22 17:56:52 2026

@author: frecines
"""

import sys
from PyQt5 import uic, QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtGui import QPalette
from PyQt5.QtGui import QColor
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt
import traceback
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
from matplotlib.lines import Line2D 
from scipy.interpolate import make_interp_spline
import matplotlib.patches as patches
import ezdxf
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.widgets import Button
import mplcursors
from matplotlib.widgets import RangeSlider, CheckButtons
from scipy.interpolate import make_interp_spline
from matplotlib.widgets import TextBox  # Importamos el buscador de texto
from mpl_toolkits.mplot3d import Axes3D
import io
import re
import unicodedata
import colorsys

try:
    from docx import Document
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

OBJETIVO_DIARIO_M = 35  # Metros objetivo por máquina por día
META_TOTAL_PROYECTO = 28000  # Meta contractual total del proyecto, en metros
STALE_DAYS = 10  # Días sin reportar avance, sin señales en comentarios, para considerar retirada

# Fechas de COMPROMISO CONTRACTUAL de ingreso de cada máquina. El ingreso físico a veces se
# demora, pero la meta contractual (Meta_Diaria_Contrato) se cuenta desde esta fecha porque es
# el compromiso de la empresa de perforación. Una máquina que no esté en este diccionario usa
# como respaldo la fecha de su primer reporte real.
FECHAS_COMPROMISO_INGRESO = {
    'HYDX-5A-05': '2026-01-04',
    'HYDX-5A-06': '2026-02-27',
    'HYDX-5A-07': '2025-12-26',
    'JINSHI YN-1500': '2026-01-20',
    'XZCR-N18A': '2026-03-25',
}

# Frases (sin tildes, minúsculas) que en el comentario de turno (celda O4) indican que la
# máquina salió realmente de operación — permiten detectar la salida el mismo día que se
# reporta, sin esperar a que pasen STALE_DAYS de silencio.
EXIT_KEYWORDS = [
    "finalizacion del contrato", "fin del contrato", "fin de contrato",
    "termino de contrato", "desmovilizacion", "retiro definitivo",
    "culminan las actividades operativas", "culmina las actividades operativas",
    "salida definitiva", "cierre de operaciones",
]

# Frases que indican que la máquina sigue activa aunque no esté perforando (traslado,
# estandarización de plataforma, etc.) — evitan "falsos retiros" cuando el silencio en los
# reportes se debe a esto y no a que la máquina realmente se fue.
STANDBY_KEYWORDS = [
    "estandarizacion", "desestandarizacion", "traslado de maquina",
    "traslado de la maquina", "movilizacion", "reubicacion", "cambio de plataforma",
]


def _strip_accents(texto):
    return ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))


# Log de errores no controlados: cualquier excepción que llegue hasta aquí se guarda en
# este archivo (además de mostrarse en pantalla), para poder revisarla después sin
# depender de tener abierta la consola de Spyder en el momento en que ocurrió.
LOG_ERRORES = "drilltracking_errores.log"


def _registrar_en_log(texto):
    try:
        with open(LOG_ERRORES, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n{texto}\n")
    except Exception:
        pass


def manejar_excepcion_global(exc_type, exc_value, exc_tb):
    """Captura cualquier excepción NO controlada por un try/except en toda la app
    (sys.excepthook). Sin esto, un error inesperado en un slot de PyQt5 puede cerrar la
    aplicación sin ningún mensaje o con un error genérico que no dice en qué archivo ni
    en qué línea ocurrió (como pasaba con el ValueError de formato de fecha al cargar un
    Excel de reporte). Aquí se muestra el tipo de error, dónde ocurrió exactamente
    (archivo y línea de nuestro propio código) y se deja constancia en LOG_ERRORES."""
    texto_traceback = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    print("\n" + "#" * 60)
    print("EXCEPCIÓN NO CONTROLADA:")
    print(texto_traceback)
    print("#" * 60 + "\n")
    _registrar_en_log(texto_traceback)

    ubicacion = ""
    frames = [f for f in traceback.extract_tb(exc_tb) if f.filename.endswith(('DrillTracking.py', os.path.basename(__file__)))]
    if frames:
        ultimo = frames[-1]
        ubicacion = f"\n\nUbicación: {os.path.basename(ultimo.filename)}, línea {ultimo.lineno}, función '{ultimo.name}'"

    mensaje = (
        f"Ocurrió un error inesperado que no fue capturado por la aplicación:\n\n"
        f"Tipo: {exc_type.__name__}\n"
        f"Detalle: {exc_value}"
        f"{ubicacion}\n\n"
        f"El detalle completo (traceback) se guardó en '{LOG_ERRORES}'.\n"
        f"La aplicación intentará seguir funcionando; si vuelve a fallar, cierre y "
        f"vuelva a abrir el programa."
    )
    try:
        QMessageBox.critical(None, "Error no controlado", mensaje)
    except Exception:
        pass


# Etapas de trabajo en el almacén de núcleos (recepción, fotografía, logueos, corte y
# muestreo), y en qué hoja/columna del Excel de GeoCore se registra el avance de cada una.
# (nombre_mostrado, hoja_excel, columna_valor, unidad)
ETAPAS_ALMACEN = [
    ("Recepción", "Recepción", "Metros", "m"),
    ("Fotografía", "Fotografía", "Avance", "m"),
    ("Logueo Geotécnico", "L_Geotécnico", "Avance", "m"),
    ("Logueo Geológico", "L_Geológico", "Avance", "m"),
    ("Corte", "Corte", "AVANCE", "m"),
    ("Muestreo", "Muestreo", "MUESTRAS", "muestras"),
]


# Categorías "canónicas" de paradas, derivadas analizando TODAS las etiquetas reales usadas
# en N14:N28 (cargables) y R16:R37 (no cargables) de los reportes de ejemplo — no una lista
# inventada. Cada entrada es (nombre_mostrado, prefijo_o_prefijos): una etiqueta de fila se
# agrupa bajo `nombre_mostrado` si, normalizada (sin tildes/mayúsculas/puntuación), EMPIEZA
# con alguno de sus prefijos (si `prefijo_o_prefijos` es None, se usa el propio nombre). Esto
# es necesario porque varias filas "catch-all" (sobre todo "Otros") se usan en la práctica
# para texto libre por turno (ej. "Otros: Traslado de personal/espera por grua"), y algunas
# categorías se escriben con pequeñas variaciones entre máquinas/turnos (ej. "Reperforando"
# vs "Reperforado", "Sac. y baj. tuberia otro." vs "...otros"). Lo que no calza con NINGÚN
# prefijo cae en "Otros (sin clasificar)" y se registra en self.categorias_paradas_no_reconocidas
# para poder revisarlo (puede ser una categoría nueva, o un dato mal tipeado en esa celda).
CATEGORIAS_CARGABLES_PARADAS = [
    ("Acondicionado/Obturado", None),
    ("Pruebas de Medición", None),
    ("Colocación y Arrastre", None),
    ("Movilizacion entre Fases", None),
    ("Movilización y estandarizacion por minera", None),
    ("Espera por plataforma", None),
    ("Espera por condiciones climáticas", None),
    ("Espera por decisión de cliente", None),
    ("Espera por voladura", None),
    ("Espera por tomernta eléctrica", None),  # así está escrito en la plantilla real (typo de origen)
    ("Espera por condiciones sociales", None),
    ("Espera Otros", None),
    ("Espera por combustible", None),
    ("Espera por repuesto", None),
]

CATEGORIAS_NO_CARGABLES_PARADAS = [
    ("Check list y Charla de seguridad", None),
    ("Orden y limpieza", None),
    ("Sac. y baj. tuberia cambio de linea", None),
    ("Sacando tuberia por fin de pozo", None),
    ("Sac. y baj. tuberia para revisar broca", None),
    ("Sac. y baj. tuberia cambio de broca", None),
    ("Sac. y baj. tuberia por tubo interior", None),
    ("Sac. y baj. tuberia otros", "Sac. y baj. tuberia otro"),
    ("Acondicionamiento de sondaje", None),
    ("Rimado e Instalación de Caising", None),
    ("Desistalación de caising", None),
    ("Recuperación tuberia/herramienta atrapada", None),
    ("Recuperación y maniobra por atrapamientos", None),
    ("Reperforando", "Reperfora"),  # cubre "Reperforando", "Reperforado" y "Reperforado/ Triconeo"
    ("Traslado de materiales, accesorios y máquina", "Traslado de materiales, accesorios y"),
    ("Instalación/Desinstalación de Plataforma o Máquina", "Instalacion"),
    ("Estandarización y Desestandarizacion de Plataforma", None),
    ("Estandarización de Plataforma", None),
    ("Stand By", None),
    ("Esperando por normas de seguridad", None),
    ("Horas no trabajadas por normas del cliente", None),
    ("Mantenimiento y Reparación", ("Mantenimiento y Reparación", "Mantenimiento y/o Reparación")),
    ("Mantenimiento", None),
    ("Horas muertas", None),
    ("Refrigerio", None),
    ("Inyeccion de Obturante", None),
    ("Ensanchado/Tricono", "Ensanchado"),
    ("Otros: Traslado de personal", "Otros traslado de personal"),
    ("Otros (Problemas Operativos)", "Otros problemas operativo"),
    ("Otros", None),
]


qtCreatorFile = "U45x.ui"
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)
class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.PlanPerf.setEnabled(False)
        self.CtrlProceso.setEnabled(False)
        self.CurvaS.setEnabled(False)
        self.Plano.setEnabled(False)
        self.Dashboard.setEnabled(False)
        self.GrafDesv.setEnabled(False)
        self.FileLab.setEnabled(False)
        self.PlotLab.setEnabled(False)
        self.Paradas.setEnabled(False)
        self.WordReport.setEnabled(False)
        self.PlotAlmacenBtn.setEnabled(False)
        self.Reportes.clicked.connect(self.LoadReport)
        self.PlanPerf.clicked.connect(self.LoadPlan)
        self.CtrlProceso.clicked.connect(self.LoadProc)
        self.Topografia.clicked.connect(self.LoadTopo)
        self.CurvaS.clicked.connect(self.PlotS)
        self.Plano.clicked.connect(self.PlotPlano)
        self.Dashboard.clicked.connect(self.PlotDash)
        self.Desviacion.clicked.connect(self.LoadDesv)
        self.GrafDesv.clicked.connect(self.PlotDesv)
        self.about.clicked.connect(self.toggle_text)
        self.FileLab.clicked.connect(self.ImportLab)
        self.PlotLab.clicked.connect(self.PlotearLab)
        self.Paradas.clicked.connect(self.PlotParadas)
        self.WordReport.clicked.connect(self.GenerateWordReport)
        self.FileAlmacen.clicked.connect(self.LoadAlmacen)
        self.PlotAlmacenBtn.clicked.connect(self.PlotAlmacen)
        self.Mensajes.setStyleSheet("color: blue; font-weight: bold;")
        self.errores_carga = []  # Errores detectados al leer los Excel de reportes (LoadReport)
        self.categorias_paradas_no_reconocidas = []  # Etiquetas de paradas sin categoría conocida

        fecha_limite = datetime(2026, 11, 30)
        fecha_actual = datetime.now()
        if fecha_actual > fecha_limite:
            mensaje = "Falta algunos componentes"
            QMessageBox.warning(self, "Error", mensaje)
            sys.exit()
            
    def _compute_machine_lifecycle(self):
        """Determina el inicio y fin de operación de cada máquina.

        - 'start': fecha de COMPROMISO CONTRACTUAL (FECHAS_COMPROMISO_INGRESO) si está
          configurada, ya que la meta contractual se cuenta desde ahí aunque el ingreso físico
          se haya demorado. Si la máquina no está en el diccionario (equipo nuevo aún no
          configurado), se usa la fecha de su primer reporte real.
        - 'end': se detecta revisando el comentario de turno (columna 'Comentario', celda O4
          de cada hoja) de los últimos reportes de la máquina:
            * si aparece una frase de EXIT_KEYWORDS -> se marca como retirada de inmediato,
              sin esperar a que pase el umbral de silencio.
            * si aparece una frase de STANDBY_KEYWORDS (traslado/estandarización) -> se
              mantiene activa aunque lleve varios días sin reportar avance, evitando "falsos
              retiros" durante paradas cortas de reubicación.
            * si no hay ninguna señal explícita, se usa STALE_DAYS de silencio como respaldo.
        """
        fecha_max_global = self.df["Fecha"].max()
        lifecycle = {}
        for maquina, grupo in self.df.groupby("Maquina"):
            fechas = grupo["Fecha"]
            inicio_compromiso = pd.to_datetime(FECHAS_COMPROMISO_INGRESO.get(maquina), errors='coerce')
            inicio = inicio_compromiso if pd.notna(inicio_compromiso) else fechas.min()

            ultimo_reporte = fechas.max()
            activa = (fecha_max_global - ultimo_reporte).days <= STALE_DAYS

            # fillna ANTES de astype(str): pandas deja los NaN como float incluso después de
            # astype(str) (no los convierte al texto "nan"), así que si el comentario de
            # turno viene vacío en alguna de las últimas hojas, " ".join(...) fallaba con
            # "TypeError: sequence item N: expected str instance, float found".
            comentarios_recientes = grupo.sort_values("Fecha").tail(8)["Comentario"].fillna("").astype(str)
            texto_reciente = _strip_accents(" ".join(comentarios_recientes).lower())

            if any(kw in texto_reciente for kw in EXIT_KEYWORDS):
                activa = False
            elif any(kw in texto_reciente for kw in STANDBY_KEYWORDS):
                activa = True

            lifecycle[maquina] = {
                "start": inicio,
                "end": None if activa else ultimo_reporte,
            }
        return lifecycle

    def _registrar_error_carga(self, archivo, hoja, tipo, detalle):
        """Guarda un error detectado al leer un reporte (archivo Excel, hoja, tipo de
        problema y detalle) para poder mostrarlo/exportarlo después, en vez de perderlo
        en la consola o interrumpir la carga de todos los demás archivos."""
        self.errores_carga.append({
            "archivo": str(archivo),
            "hoja": str(hoja),
            "tipo": str(tipo),
            "detalle": str(detalle),
        })
        print(f"[AVISO] {archivo} | hoja '{hoja}' | {tipo}: {detalle}")

    @staticmethod
    def _normalizar_para_match(texto):
        """Minúsculas, sin tildes y sin puntuación (todo lo que no sea letra/número se
        vuelve un espacio, y los espacios se colapsan). Así "Otros: Traslado de personal"
        y "Otros :Traslado de personal" (o "Otros/Traslado...") se comparan igual — la
        puntuación de estas celdas es inconsistente entre turnos, pero el texto no."""
        texto = _strip_accents(texto.lower())
        texto = re.sub(r"[^a-z0-9]+", " ", texto).strip()
        return texto

    def _canonizar_categoria_parada(self, etiqueta, categorias_canonicas):
        """Empareja `etiqueta` (texto tal cual viene de la celda) con la categoría canónica
        cuyo prefijo (normalizado, ver _normalizar_para_match) más largo calce con ella.
        `categorias_canonicas` es una lista de (nombre_mostrado, prefijo_o_prefijos); si el
        prefijo es None se usa el propio nombre. Esto agrupa correctamente variantes de
        texto libre como "Otros: Traslado de personal/espera por grua" bajo una única
        categoría, sin perder categorías legítimas cuyo texto tiene un sufijo distinto (ej.
        "Espera por decisión de cliente CIERRE DE VIAS" -> "Espera por decisión de
        cliente"). Devuelve None si no calza con ninguna."""
        etiqueta_norm = self._normalizar_para_match(etiqueta)
        mejor_nombre = None
        mejor_largo = -1
        for nombre_mostrado, prefijos in categorias_canonicas:
            if prefijos is None:
                prefijos = (nombre_mostrado,)
            elif isinstance(prefijos, str):
                prefijos = (prefijos,)
            for prefijo in prefijos:
                prefijo_norm = self._normalizar_para_match(prefijo)
                if etiqueta_norm.startswith(prefijo_norm) and len(prefijo_norm) > mejor_largo:
                    mejor_nombre = nombre_mostrado
                    mejor_largo = len(prefijo_norm)
        return mejor_nombre

    def _extraer_categoria_paradas(self, df_temp, fila_inicio, fila_fin, col_label, col_valor,
                                    categorias_canonicas, archivo, hoja, excluir_labels=()):
        """Lee pares (etiqueta, horas) fila a fila entre `fila_inicio` y `fila_fin` (números
        de fila de Excel, 1-based) de las columnas `col_label`/`col_valor` (0-based, iloc), y
        agrupa cada etiqueta bajo su categoría canónica más cercana (ver
        _canonizar_categoria_parada). Lo que no calza con ninguna categoría conocida se
        registra en "Otros (sin clasificar)" y se deja constancia en errores_carga, en vez de
        crear una columna nueva por cada variante de texto libre (`excluir_labels` filtra
        subtítulos de sección que no son una categoría real, ej. "Tiempo de Espera/Stand by")."""
        categorias = {}
        n_filas, n_cols = df_temp.shape
        for fila_excel in range(fila_inicio, fila_fin + 1):
            i = fila_excel - 1
            if i >= n_filas or col_label >= n_cols:
                continue
            etiqueta = df_temp.iat[i, col_label]
            if pd.isna(etiqueta):
                continue
            etiqueta = " ".join(str(etiqueta).split())
            if not etiqueta or etiqueta in excluir_labels:
                continue
            valor = df_temp.iat[i, col_valor] if col_valor < n_cols else np.nan
            valor = pd.to_numeric(valor, errors="coerce")
            valor = 0.0 if pd.isna(valor) else float(valor)

            categoria = self._canonizar_categoria_parada(etiqueta, categorias_canonicas)
            if categoria is None:
                categoria = "Otros (sin clasificar)"
                # Diagnóstico aparte de errores_carga: esto no es un error que impida cargar
                # el reporte, solo una etiqueta de parada que no calzó con ninguna categoría
                # conocida (posible texto libre o categoría nueva no contemplada todavía).
                self.categorias_paradas_no_reconocidas.append({
                    "archivo": archivo, "hoja": hoja, "fila": fila_excel, "etiqueta": etiqueta,
                })
            categorias[categoria] = categorias.get(categoria, 0.0) + valor
        return categorias

    def _extraer_tabla_brocas(self, df_temp, texto_ancla, col_diametro, col_desde, col_hasta, col_total,
                               ventana_inicio=26, ventana_fin=45, max_filas_datos=4):
        """Busca la fila cuya celda en `col_diametro` sea `texto_ancla` (ej. "BROCAS",
        "TRICONO") dentro de una ventana de filas de la hoja, en vez de asumir una posición
        fija: en algunos reportes esta tabla aparece corrida 1-2 filas por un encabezado
        duplicado insertado a mano (mismo tipo de inconsistencia que ya se vio con fechas y
        con la celda de Máquina). El encabezado "Diametro"/"Díametro" puede estar 1-3 filas
        debajo del ancla; se leen hasta `max_filas_datos` filas debajo de ESE encabezado.
        Una fila cuenta como dato real solo si su celda de diámetro es un código corto tipo
        HQ/NQ/PQ/BQ (2-3 letras) — así se descartan encabezados repetidos, textos sueltos
        ("USADA", etc.) u otras filas de la plantilla que caen dentro de la ventana."""
        n_filas, n_cols = df_temp.shape
        if col_diametro >= n_cols:
            return []

        fila_ancla = None
        ancla_norm = self._normalizar_para_match(texto_ancla)
        for i in range(max(0, ventana_inicio - 1), min(n_filas, ventana_fin)):
            valor = df_temp.iat[i, col_diametro]
            if pd.notna(valor) and self._normalizar_para_match(str(valor)) == ancla_norm:
                fila_ancla = i
                break
        if fila_ancla is None:
            return []

        fila_header = None
        for offset in (1, 2, 3):
            i = fila_ancla + offset
            if i >= n_filas:
                break
            valor = df_temp.iat[i, col_diametro]
            if pd.notna(valor) and self._normalizar_para_match(str(valor)).startswith("diametro"):
                fila_header = i
                break
        if fila_header is None:
            return []

        registros = []
        for offset in range(1, max_filas_datos + 1):
            i = fila_header + offset
            if i >= n_filas:
                break
            diametro = df_temp.iat[i, col_diametro]
            if pd.isna(diametro):
                continue
            diametro_txt = str(diametro).strip().upper()
            if not re.fullmatch(r"[A-Z]{2,3}", diametro_txt):
                continue
            desde = pd.to_numeric(df_temp.iat[i, col_desde], errors="coerce") if col_desde < n_cols else np.nan
            hasta = pd.to_numeric(df_temp.iat[i, col_hasta], errors="coerce") if col_hasta < n_cols else np.nan
            total = pd.to_numeric(df_temp.iat[i, col_total], errors="coerce") if col_total < n_cols else np.nan
            registros.append({
                "diametro": diametro_txt,
                "desde": float(desde) if pd.notna(desde) else np.nan,
                "hasta": float(hasta) if pd.notna(hasta) else np.nan,
                "total": float(total) if pd.notna(total) else np.nan,
            })
        return registros

    def _maquinas_validas(self, df=None, columna="Maquina"):
        """Lista ordenada de valores únicos de `columna`, ignorando vacíos/NaN. Evita el
        TypeError de sorted()/unique() al mezclar texto con NaN (float) cuando alguna
        celda de Máquina viene vacía en un reporte (ver PlotDash)."""
        origen = df if df is not None else self.df
        valores = origen[columna].dropna().astype(str).str.strip()
        valores = valores[valores != ""]
        return sorted(valores.unique())

    def _exportar_reporte_errores(self, carpeta):
        """Exporta a Word (o a texto plano si no hay python-docx instalado) el listado de
        errores detectados durante la carga de reportes: en qué archivo Excel, en qué
        hoja y por qué falló, para poder ubicarlos y corregirlos sin depender de tener
        abierta la consola de Spyder."""
        if not self.errores_carga:
            return None

        fecha_hoy = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        try:
            if DOCX_AVAILABLE:
                nombre_archivo = os.path.join(carpeta, f"Errores_Carga_Reportes_{fecha_hoy}.docx")
                doc = Document()
                titulo = doc.add_heading('Errores Detectados al Cargar Reportes de Perforación', level=0)
                titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}")
                doc.add_paragraph(
                    f"Se detectaron {len(self.errores_carga)} inconsistencia(s) durante la lectura de "
                    f"los archivos Excel de la carpeta seleccionada. Las filas afectadas fueron "
                    f"excluidas del análisis para que el resto de la información sí se pueda usar; "
                    f"corrija el Excel indicado (archivo y hoja) y vuelva a cargar la carpeta."
                )

                tabla = doc.add_table(rows=1, cols=4)
                tabla.style = 'Light Grid Accent 1'
                hdr = tabla.rows[0].cells
                hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = (
                    'Archivo Excel', 'Hoja', 'Tipo de Error', 'Detalle'
                )
                for err in self.errores_carga:
                    row = tabla.add_row().cells
                    row[0].text = err["archivo"]
                    row[1].text = err["hoja"]
                    row[2].text = err["tipo"]
                    row[3].text = err["detalle"]

                doc.save(nombre_archivo)
                return nombre_archivo
            else:
                nombre_archivo = os.path.join(carpeta, f"Errores_Carga_Reportes_{fecha_hoy}.txt")
                with open(nombre_archivo, "w", encoding="utf-8") as f:
                    f.write("ERRORES DETECTADOS AL CARGAR REPORTES DE PERFORACIÓN\n")
                    f.write(f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n")
                    for err in self.errores_carga:
                        f.write(f"Archivo: {err['archivo']}\n")
                        f.write(f"Hoja: {err['hoja']}\n")
                        f.write(f"Tipo: {err['tipo']}\n")
                        f.write(f"Detalle: {err['detalle']}\n")
                        f.write("-" * 60 + "\n")
                return nombre_archivo
        except Exception as e:
            print(f"No se pudo exportar el reporte de errores: {e}")
            return None

    def LoadReport(self):
            self.PlanPerf.setEnabled(False)
            self.CtrlProceso.setEnabled(False)
            self.CurvaS.setEnabled(False)
            self.Plano.setEnabled(False)
            self.Dashboard.setEnabled(False)      
            self.FileLab.setEnabled(False) 
            self.PlotLab.setEnabled(False) 
            self.Mensajes.clear()
            self.errores_carga = []
            self.categorias_paradas_no_reconocidas = []
            # 1. Obtener ruta y validar
            dir_path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta")
            if not dir_path:
                return

            ruta = Path(dir_path)
            archivos = list(ruta.glob("*.xlsx"))

            if not archivos:
                QMessageBox.information(self, "Aviso", "No se encontraron archivos .xlsx en la carpeta.")
                return

            def leer_reporte(ruta_archivo):
                datos = []
                try:
                    xls = pd.ExcelFile(ruta_archivo)
                except Exception as e:
                    self._registrar_error_carga(ruta_archivo.name, "-", "No se pudo abrir el archivo", e)
                    return datos

                for hoja in xls.sheet_names:
                    try:
                        # Usamos el objeto xls para mayor velocidad
                        df_temp = pd.read_excel(xls, sheet_name=hoja, header=None)

                        # Horas Perforación Normal (R14/U14): están DENTRO del rango que suma
                        # "Horas Operativas" (U38, Hrs Paradas no-cargables), mezcladas con el
                        # tiempo muerto real. Se extraen aparte para poder mostrarlas como su
                        # propia categoría (ver _build_resumen_paradas_detallado).
                        horas_perf_normal = pd.to_numeric(df_temp.iloc[13, 20], errors="coerce")
                        horas_perf_normal = 0.0 if pd.isna(horas_perf_normal) else float(horas_perf_normal)

                        # Detalle por categoría de paradas: se lee la etiqueta y el valor de
                        # cada fila del rango y se agrupa bajo su categoría canónica más
                        # cercana (ver _canonizar_categoria_parada), porque algunas filas se
                        # usan para texto libre por turno. "Tiempo de Espera/Stand by" es un
                        # subtítulo de sección dentro de N14:N28, no una categoría real.
                        cat_cargables = self._extraer_categoria_paradas(
                            df_temp, fila_inicio=14, fila_fin=28, col_label=13, col_valor=16,
                            categorias_canonicas=CATEGORIAS_CARGABLES_PARADAS,
                            archivo=ruta_archivo.name, hoja=hoja,
                            excluir_labels={"Tiempo de Espera/Stand by"}
                        )
                        cat_no_cargables = self._extraer_categoria_paradas(
                            df_temp, fila_inicio=16, fila_fin=37, col_label=17, col_valor=20,
                            categorias_canonicas=CATEGORIAS_NO_CARGABLES_PARADAS,
                            archivo=ruta_archivo.name, hoja=hoja,
                        )

                        # Tablas de brocas/tricono (A29:F32 y G29:L32 aprox.) para el resumen
                        # de metros por diámetro de broca. Se buscan por ancla ("BROCAS"/
                        # "TRICONO") en vez de posición fija: algunas hojas tienen esta tabla
                        # corrida por un encabezado duplicado insertado a mano.
                        detalle_brocas = self._extraer_tabla_brocas(
                            df_temp, "BROCAS", col_diametro=0, col_desde=3, col_hasta=4, col_total=5
                        )
                        detalle_tricono = self._extraer_tabla_brocas(
                            df_temp, "TRICONO", col_diametro=6, col_desde=7, col_hasta=9, col_total=11
                        )

                        # Extraemos datos (basado en tu lógica original). "Hoja" se conserva
                        # sin recortar para poder ubicar el error exacto más abajo, aunque
                        # "Turno" luego se recorte para quedarse solo con el nombre del turno.
                        datos.append([
                            ruta_archivo.name,    # Archivo de origen
                            hoja,                 # Hoja (nombre real de la pestaña en Excel)
                            hoja,                 # Turno (se recorta más abajo)
                            df_temp.iloc[5, 4],   # Maquina
                            df_temp.iloc[3, 6],   # Sondaje
                            df_temp.iloc[9, 6],   # Prof Programada
                            df_temp.iloc[8, 6],   # Avance
                            df_temp.iloc[28, 16], # Hrs Paradas (Horas Operativas Cargables, Q29)
                            df_temp.iloc[37, 20], # Hrs Operativas (Horas no Cargables, U38 - incluye Perf. Normal)
                            df_temp.iloc[3, 14],  # Comentario de turno (O4)
                            horas_perf_normal,     # Horas Perforacion Normal (U14)
                            cat_cargables,         # Detalle por categoría cargable (N14:N28 / Q14:Q28)
                            cat_no_cargables,      # Detalle por categoría no cargable (R16:R37 / U16:U37)
                            detalle_brocas,        # Lista de {diametro,desde,hasta,total} (tabla BROCAS)
                            detalle_tricono,       # Lista de {diametro,desde,hasta,total} (tabla TRICONO)
                        ])
                    except Exception as e:
                        self._registrar_error_carga(
                            ruta_archivo.name, hoja,
                            "No se pudo leer la hoja (estructura/celdas inesperadas)", e
                        )
                        continue
                return datos

            # 2. Procesar archivos
            data_total = []
            for archivo in archivos:
                data_total.extend(leer_reporte(archivo))

            if not data_total:
                mensaje = "No se pudo extraer información válida de los reportes, compruebe que la carpeta contenga los reportes correctos."
                if self.errores_carga:
                    archivo_errores = self._exportar_reporte_errores(str(ruta))
                    mensaje += f"\n\nSe detectaron {len(self.errores_carga)} error(es) al leer los Excel."
                    if archivo_errores:
                        mensaje += f"\nDetalle exportado a:\n{archivo_errores}"
                QMessageBox.warning(self, "Error", mensaje)
                return

            try:
                # 3. Crear DataFrame
                columnas = ["Archivo", "Hoja", "Turno", "Maquina", "Sondaje", "Profundidad Programada",
                            "Avance Turno", "Horas Paradas", "Horas Operativas", "Comentario",
                            "Horas Perforacion Normal", "Detalle_Cargables", "Detalle_No_Cargables",
                            "Detalle_Brocas", "Detalle_Tricono"]

                self.df = pd.DataFrame(data_total, columns=columnas)

                # --- Fecha: se parsea con el formato esperado (dd-mm-aa) y, para las filas
                # que no calcen (p.ej. una hoja nombrada con "_" en vez de "-", o con otro
                # separador), se intenta un segundo parseo más flexible ANTES de descartarlas.
                # Antes, un solo pd.to_datetime(..., format="%d-%m-%y") sin errors="coerce"
                # lanzaba un ValueError sin capturar y cerraba/colgaba la aplicación entera
                # (ej. "ValueError: time data '04_09-26' doesn't match format '%d-%m-%y'"),
                # sin decir en qué archivo/hoja estaba el problema.
                fecha_texto = self.df["Turno"].astype(str).str[:8]
                fecha_parseada = pd.to_datetime(fecha_texto, format="%d-%m-%y", errors="coerce")

                pendientes = fecha_parseada.isna()
                if pendientes.any():
                    fecha_alternativa = fecha_texto[pendientes].str.replace("_", "-", regex=False).str.strip()
                    fecha_parseada.loc[pendientes] = pd.to_datetime(
                        fecha_alternativa, dayfirst=True, errors="coerce"
                    )

                self.df["Fecha"] = fecha_parseada

                sin_fecha = self.df["Fecha"].isna()
                if sin_fecha.any():
                    for _, fila in self.df[sin_fecha].iterrows():
                        self._registrar_error_carga(
                            fila["Archivo"], fila["Hoja"],
                            "Fecha de la hoja no reconocida",
                            f'El nombre de la hoja empieza con "{fila["Hoja"][:8]}", que no coincide '
                            f'con el formato esperado dd-mm-aa (ej. "04-09-26"); la fila fue excluida.'
                        )
                    self.df = self.df[~sin_fecha].copy()

                # --- Fecha fuera de rango: una pestaña con el año mal tipeado (ej. "02-09-02"
                # en vez de "02-09-26") pasa el parseo SIN error, pero arruina las agrupaciones
                # por semana/mes y los rangos de los sliders (PlotDash, Curva S, etc.). Se marca
                # como sospechosa y se excluye si cae muy fuera de la ventana razonable.
                if not self.df.empty:
                    fecha_min_valida = pd.Timestamp("2020-01-01")
                    fecha_max_valida = pd.Timestamp.now() + pd.DateOffset(days=30)
                    fecha_fuera_rango = (self.df["Fecha"] < fecha_min_valida) | (self.df["Fecha"] > fecha_max_valida)
                    if fecha_fuera_rango.any():
                        for _, fila in self.df[fecha_fuera_rango].iterrows():
                            self._registrar_error_carga(
                                fila["Archivo"], fila["Hoja"],
                                "Fecha fuera de rango (posible año mal escrito)",
                                f'La hoja "{fila["Hoja"]}" se interpretó como '
                                f'{fila["Fecha"].strftime("%d/%m/%Y")}, fuera del rango esperado del '
                                f'proyecto; probablemente el año en el nombre de la pestaña está mal '
                                f'escrito. La fila fue excluida.'
                            )
                        self.df = self.df[~fecha_fuera_rango].copy()

                if self.df.empty:
                    archivo_errores = self._exportar_reporte_errores(str(ruta))
                    mensaje = f"No quedó ninguna fila válida: las {len(self.errores_carga)} hojas leídas tuvieron errores."
                    if archivo_errores:
                        mensaje += f"\n\nDetalle exportado a:\n{archivo_errores}"
                    QMessageBox.critical(self, "Error", mensaje)
                    return

                self.df["Turno"] = self.df["Turno"].str[9:]

                # --- Máquina: una celda vacía (NaN) mezclada con texto rompe cualquier
                # sorted()/unique() posterior con "TypeError: '<' not supported between
                # instances of 'float' and 'str'" (ver PlotDash). Se detecta y excluye aquí,
                # dejando constancia de en qué archivo/hoja está la celda vacía.
                maquina_vacia = self.df["Maquina"].isna() | (self.df["Maquina"].astype(str).str.strip() == "")
                if maquina_vacia.any():
                    for _, fila in self.df[maquina_vacia].iterrows():
                        self._registrar_error_carga(
                            fila["Archivo"], fila["Hoja"],
                            "Celda de Máquina vacía",
                            "La celda de Máquina (fila 6, columna E de la hoja) está vacía; la fila fue excluida del análisis."
                        )
                    self.df = self.df[~maquina_vacia].copy()
                self.df["Maquina"] = self.df["Maquina"].astype(str).str.strip()

                if self.df.empty:
                    archivo_errores = self._exportar_reporte_errores(str(ruta))
                    mensaje = f"No quedó ninguna fila válida después de limpiar Fecha/Máquina ({len(self.errores_carga)} error(es))."
                    if archivo_errores:
                        mensaje += f"\n\nDetalle exportado a:\n{archivo_errores}"
                    QMessageBox.critical(self, "Error", mensaje)
                    return

                self.df = self.df[~self.df['Sondaje'].astype(str).str.startswith('ED')]
                self.df = self.df.sort_values("Fecha")

                df_agrupado = self.df.groupby(['Fecha', 'Maquina'])['Avance Turno'].sum().reset_index()
                self.df_final = df_agrupado.pivot(index='Fecha', columns='Maquina', values='Avance Turno')
                self.df_final['Total Dia'] = self.df_final.sum(axis=1)

                self.df_final = self.df_final.reset_index()
                self.df_final.columns.name = None

                self.df_final = self.df_final.fillna(0)

                self._actualizar_revision_con_validacion_almacen()
                self.df["Fecha"] = pd.to_datetime(self.df["Fecha"])
                self.machine_lifecycle = self._compute_machine_lifecycle()

                print("Datos cargados exitosamente:")
                print(self.df.head())

                mensaje_exito = f"Se cargaron {len(self.df)} registros."
                if self.errores_carga:
                    archivo_errores = self._exportar_reporte_errores(str(ruta))
                    mensaje_exito += (
                        f"\n\nSe detectaron {len(self.errores_carga)} inconsistencia(s) en los Excel de "
                        f"origen (filas excluidas del análisis)."
                    )
                    if archivo_errores:
                        mensaje_exito += f"\nDetalle exportado a:\n{archivo_errores}"
                    else:
                        mensaje_exito += "\nRevise la consola para el detalle."

                QMessageBox.information(self, "Éxito", mensaje_exito)
                self.Mensajes.setText(
                    "Reportes Cargados" +
                    (f" — {len(self.errores_carga)} error(es), ver reporte exportado" if self.errores_carga else "")
                )
                self.CtrlProceso.setEnabled(True)

            except Exception as e:
                # --- MANEJO CON TRACEBACK: si algo más adelante (pivot, groupby, etc.)
                # falla, se muestra el error detallado (incluyendo dónde ocurrió exactamente)
                # en vez de cerrar la aplicación o dejar al usuario sin ninguna pista.
                print("\n" + "="*50)
                print("ERROR DETECTADO EN LOADREPORT:")
                traceback.print_exc()
                print("="*50 + "\n")

                ubicacion = ""
                frames = traceback.extract_tb(e.__traceback__)
                if frames:
                    ultimo = frames[-1]
                    ubicacion = f"\nUbicación: {os.path.basename(ultimo.filename)}, línea {ultimo.lineno}, función '{ultimo.name}'"

                error_detallado = f"Tipo de error: {type(e).__name__}\nDetalle: {str(e)}{ubicacion}"
                mensaje = f"Error al procesar los reportes: {error_detallado}\n\nRevise la consola para más detalles."

                # Si ya se habían detectado inconsistencias en filas anteriores antes de que
                # este error interrumpiera el resto del procesamiento, se exportan igual: no
                # tiene sentido perder ese diagnóstico solo porque un paso posterior falló.
                if self.errores_carga:
                    archivo_errores = self._exportar_reporte_errores(str(ruta))
                    if archivo_errores:
                        mensaje += f"\n\nAdemás se habían detectado {len(self.errores_carga)} inconsistencia(s) antes del error; detalle exportado a:\n{archivo_errores}"

                QMessageBox.critical(self, "Error de Ejecución", mensaje)
                self.Mensajes.setText(mensaje)
            
    def LoadPlan(self):
            self.Mensajes.clear()
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Seleccionar Avance de perforacion planeado", "", "Excel Files (*.xlsx; *.xlsm)"
            )
            
            if file_path:
                try:
                    # 1. Cargar el archivo Excel
                    self.plan = pd.read_excel(file_path)
                    
                    # Aseguramos que las fechas en ambos DF sean datetime para evitar errores de merge
                    if hasattr(self, 'df'):
                        self.df["Fecha"] = pd.to_datetime(self.df["Fecha"])
                    else:
                        raise Exception("No se han cargado los reportes previos (self.df no existe).")
    
                    self.BD_Union["INICIO.1"] = pd.to_datetime(self.BD_Union["INICIO.1"])
                    self.BD_Union["FIN.1"] = pd.to_datetime(self.BD_Union["FIN.1"])
                    
                    # Calculamos la cantidad de días del rango (incluyente)
                    self.BD_Union["Dias_Rango"] = (self.BD_Union["FIN.1"] - self.BD_Union["INICIO.1"]).dt.days + 1
                    
                    # Calculamos el descuento diario por cada Identificador
                    self.BD_Union["Descuento_Diario"] = self.BD_Union["Metraje Descontado"] / self.BD_Union["Dias_Rango"]
                    
                    # 2. Crear un dataframe de descuentos expandido por fecha
                    descuentos_lista = []
                    
                    for _, fila in self.BD_Union.iterrows():
                        if fila["Metraje Descontado"] > 0:
                            # Generar rango de fechas para este descuento
                            rango_fechas = pd.date_range(start=fila["INICIO.1"], end=fila["FIN.1"])
                            
                            df_temp = pd.DataFrame({
                                "Fecha": rango_fechas,
                                "Descuento_Valor": fila["Descuento_Diario"]
                            })
                            descuentos_lista.append(df_temp)
                    
                    # Consolidar todos los descuentos por fecha
                    if descuentos_lista:
                        df_descuentos_total = pd.concat(descuentos_lista)
                        df_descuentos_diarios = df_descuentos_total.groupby("Fecha")["Descuento_Valor"].sum().reset_index()
                    else:
                        df_descuentos_diarios = pd.DataFrame(columns=["Fecha", "Descuento_Valor"])
                    
                    # 3. Procesar el avance real (tu código base)
                    self.df["Fecha"] = pd.to_datetime(self.df["Fecha"])
                    real_diario = (
                        self.df.groupby("Fecha")["Avance Turno"]
                        .sum()
                        .reset_index()
                    )
                    real_diario.rename(columns={"Avance Turno": "Avance_Real_Bruto"}, inplace=True)
                    
                    # 4. Unir Avance Real con Descuentos y calcular el Neto
                    real_diario = real_diario.merge(df_descuentos_diarios, on="Fecha", how="left").fillna(0)
                    
                    # Restamos el descuento prorrateado al avance del día
                    real_diario["Avance_Real_Diario"] = real_diario["Avance_Real_Bruto"] - real_diario["Descuento_Valor"]
                    
                    # 5. Generar Acumulados y Curva S
                    real_diario = real_diario.sort_values("Fecha")
                    real_diario["Real_Acumulado"] = real_diario["Avance_Real_Diario"].cumsum()
                    
                    curva_s = self.plan.merge(
                        real_diario,
                        on="Fecha",
                        how="left"
                    )
                    
                    # --- El resto de tu lógica de etiquetas y marcadores se mantiene igual ---
                    curva_s["es_miercoles"] = curva_s["Fecha"].dt.weekday == 2
                    curva_s["label_plan"] = curva_s.apply(
                        lambda r: f'{r["Plan_Acumulado"]:,.0f}' if r["es_miercoles"] else "", axis=1
                    )
                    curva_s["label_real"] = curva_s.apply(
                        lambda r: f'{r["Real_Acumulado"]:,.0f}' 
                        if r["es_miercoles"] and not pd.isna(r["Real_Acumulado"]) else "", axis=1
                    )
                    curva_s["marker_plan"] = curva_s["es_miercoles"].map(lambda x: 7 if x else 3)
                    curva_s["marker_real"] = (
                        curva_s["es_miercoles"] & curva_s["Real_Acumulado"].notna()
                    ).map(lambda x: 7 if x else 3)
    
                    # Guardamos el resultado en el objeto de la clase
                    self.curva_s = curva_s
                    
                    QMessageBox.information(self, "Éxito", "Se importó y procesó correctamente el archivo.")
                    print(self.curva_s)
                    self.Mensajes.setText("Se cargo el Plan para Curva S")
                    self.CurvaS.setEnabled(True)
                    self.Plano.setEnabled(True)
                    self.Dashboard.setEnabled(True)
                    self.FileLab.setEnabled(True)
                    self.Paradas.setEnabled(True)
                    self.WordReport.setEnabled(True)

                    
    
                except Exception as e:
                    # --- AQUÍ ESTÁ EL MANEJO CON TRACEBACK ---
                    print("\n" + "="*50)
                    print("ERROR DETECTADO EN LOADPLAN:")
                    traceback.print_exc() # Esto imprime el error completo con línea exacta en la consola
                    print("="*50 + "\n")
    
                    # Mensaje para el usuario final en la interfaz
                    error_detallado = f"Tipo de error: {type(e).__name__}\nDetalle: {str(e)}"
                    mensaje = f"Error al procesar: {error_detallado}\n\nRevise la consola para más detalles."
                    QMessageBox.critical(self, "Error de Ejecución", mensaje)
                    self.Mensajes.setText(mensaje)
                    
    def _build_df_clean(self):
        """Aplica el ciclo de vida detectado de cada máquina (self.machine_lifecycle) al
        avance diario (self.df_final), y calcula la meta contractual diaria de forma
        dinámica: Maquinas_Activas(fecha) * OBJETIVO_DIARIO_M. Reemplaza los antiguos
        diccionarios de fechas y umbrales fijos por metraje planificado."""
        lifecycle_config = self.machine_lifecycle
        exclude_cols = ['Fecha', 'Total Dia', 'Semana_Label', 'Promedio_Diario_Maquinas']
        machine_cols = [col for col in self.df_final.columns if col not in exclude_cols]

        self.df_final['Fecha'] = pd.to_datetime(self.df_final['Fecha'])
        df_clean = self.df_final.copy()

        # Limpieza de datos basada en el ciclo de vida (Asignar NaN fuera de operación)
        for machine in machine_cols:
            if machine in lifecycle_config:
                start_dt = lifecycle_config[machine]['start']
                df_clean.loc[df_clean['Fecha'] < start_dt, machine] = np.nan

                end_dt = lifecycle_config[machine]['end']
                if end_dt is not None:
                    df_clean.loc[df_clean['Fecha'] > end_dt, machine] = np.nan

        def _maquinas_activas_meta(fecha):
            # La meta contractual cuenta las 5 máquinas contratadas (FECHAS_COMPROMISO_INGRESO)
            # desde el inicio del proyecto, sin importar cuándo ingresaron físicamente cada una
            # (el compromiso de metraje aplica igual desde el día 1). Solo la SALIDA detectada
            # automáticamente (cfg['end']) reduce la cantidad de máquinas contadas.
            return sum(
                1 for maquina, cfg in lifecycle_config.items()
                if maquina in FECHAS_COMPROMISO_INGRESO and (cfg['end'] is None or fecha <= cfg['end'])
            )

        df_clean['Maquinas_Activas'] = df_clean['Fecha'].apply(_maquinas_activas_meta)
        df_clean['Meta_Diaria_Contrato'] = df_clean['Maquinas_Activas'] * OBJETIVO_DIARIO_M
        # Agrupación Semanal (Miércoles a Martes)
        df_clean['Semana_Label'] = df_clean['Fecha'].dt.to_period('W-TUE')
        return df_clean, machine_cols

    def _build_resumen_semanal(self, df_clean, machine_cols):
        """Agregación semanal: suma total vs meta contractual, y promedio diario por máquina."""
        resumen_semanal_1 = df_clean.groupby('Semana_Label').agg(
            Suma_Semanal=('Total Dia', 'sum'),
            Meta_Semanal=('Meta_Diaria_Contrato', 'sum'),
            Fecha_Inicio=('Fecha', 'min'),
            Fecha_Fin=('Fecha', 'max')
        ).reset_index()
        resumen_semanal_1['Eje_X'] = (resumen_semanal_1['Fecha_Inicio'].dt.strftime('%d%b') + "-" +
                                      resumen_semanal_1['Fecha_Fin'].dt.strftime('%d%b'))

        # Al promediar las columnas de máquinas limpias, los NaN (máquinas fuera de operación) se ignoran automáticamente
        df_prom = df_clean.copy()
        df_prom['Promedio_Diario_Maquinas'] = df_prom[machine_cols].mean(axis=1)
        resumen_semanal_2 = df_prom.groupby('Semana_Label').agg(
            Promedio_Semanal=('Promedio_Diario_Maquinas', 'mean'),
            Fecha_Inicio=('Fecha', 'min'),
            Fecha_Fin=('Fecha', 'max')
        ).reset_index()
        resumen_semanal_2['Eje_X'] = (resumen_semanal_2['Fecha_Inicio'].dt.strftime('%d%b') + "-" +
                                      resumen_semanal_2['Fecha_Fin'].dt.strftime('%d%b'))
        return resumen_semanal_1, resumen_semanal_2

    def _build_resumen_paradas(self):
        """Agregación semanal POR MÁQUINA de horas operativas vs horas paradas y % de
        eficiencia. Se calcula por máquina (en vez de un total de la flota) porque un equipo
        con muchas horas paradas puede quedar oculto al promediar con el resto."""
        df_horas = self.df.copy()
        df_horas['Fecha'] = pd.to_datetime(df_horas['Fecha'])
        df_horas['Semana_Label'] = df_horas['Fecha'].dt.to_period('W-TUE')

        resumen_paradas = df_horas.groupby(['Semana_Label', 'Maquina']).agg(
            Horas_Operativas=('Horas Operativas', 'sum'),
            Horas_Paradas=('Horas Paradas', 'sum'),
        ).reset_index()
        # Eje_X se deriva de los límites propios de la semana (miércoles a martes), NO de las
        # fechas realmente reportadas por cada máquina: si se usara min/max de lo reportado,
        # dos máquinas en la MISMA semana calendario podían terminar con una etiqueta o fecha
        # de inicio ligeramente distinta (según qué día exacto reportaron turno), y eso rompía
        # la alineación de columnas entre subplots de distintas máquinas.
        resumen_paradas['Fecha_Inicio'] = resumen_paradas['Semana_Label'].dt.start_time.dt.normalize()
        resumen_paradas['Fecha_Fin'] = resumen_paradas['Semana_Label'].dt.end_time.dt.normalize()
        resumen_paradas['Eje_X'] = (resumen_paradas['Fecha_Inicio'].dt.strftime('%d%b') + "-" +
                                    resumen_paradas['Fecha_Fin'].dt.strftime('%d%b'))
        total_horas = resumen_paradas['Horas_Operativas'] + resumen_paradas['Horas_Paradas']
        resumen_paradas['Pct_Eficiencia'] = np.where(
            total_horas > 0, resumen_paradas['Horas_Operativas'] / total_horas * 100, np.nan
        )
        return resumen_paradas

    def _build_resumen_paradas_detallado(self, fecha_inicio=None, fecha_fin=None):
        """Resumen DIARIO por máquina de las 3 categorías pedidas: Horas Cargables (Q29 de
        cada turno: paradas facturables como Pruebas de Medición o Espera por tormenta
        eléctrica), Horas No Cargables sin Perforación Normal (tiempo muerto real) y Horas
        Perforación Normal (U14: tiempo productivo, que estaba mezclado dentro del total de
        "Horas Operativas"/U38 y por eso no se veía aparte). `fecha_inicio`/`fecha_fin`
        acotan el rango (ej. el Desde-Hasta del gráfico/exportable)."""
        df = self.df.copy()
        df['Fecha'] = pd.to_datetime(df['Fecha']).dt.normalize()
        if fecha_inicio is not None:
            df = df[df['Fecha'] >= pd.Timestamp(fecha_inicio).normalize()]
        if fecha_fin is not None:
            df = df[df['Fecha'] <= pd.Timestamp(fecha_fin).normalize()]

        resumen = df.groupby(['Fecha', 'Maquina']).agg(
            Horas_Cargables=('Horas Paradas', 'sum'),
            Horas_No_Cargables_Total=('Horas Operativas', 'sum'),
            Horas_Perforacion_Normal=('Horas Perforacion Normal', 'sum'),
        ).reset_index()
        resumen['Horas_No_Cargables'] = (
            resumen['Horas_No_Cargables_Total'] - resumen['Horas_Perforacion_Normal']
        ).clip(lower=0)
        resumen['Total_Horas'] = resumen['Horas_Cargables'] + resumen['Horas_No_Cargables_Total']

        for col_pct, col_horas in [
            ('Pct_Cargables', 'Horas_Cargables'),
            ('Pct_No_Cargables', 'Horas_No_Cargables'),
            ('Pct_Perforacion_Normal', 'Horas_Perforacion_Normal'),
        ]:
            resumen[col_pct] = np.where(
                resumen['Total_Horas'] > 0, resumen[col_horas] / resumen['Total_Horas'] * 100, np.nan
            )
        return resumen

    def _build_reporte_categorias_paradas(self, fecha_inicio=None, fecha_fin=None):
        """Reporte 'por medición y operativo': una fila por Fecha+Máquina con una columna
        por cada categoría real encontrada en los reportes (Pruebas de Medición, Espera por
        tormenta eléctrica, Mantenimiento y Reparación, etc.), más Perforación Normal y el
        Total del turno — el mismo desglose que el encabezado de referencia, pero calculado
        directamente desde lo que cada Excel trae en N14:N28/Q14:Q28 y R16:R37/U16:U37."""
        df = self.df.copy()
        df['Fecha'] = pd.to_datetime(df['Fecha']).dt.normalize()
        if fecha_inicio is not None:
            df = df[df['Fecha'] >= pd.Timestamp(fecha_inicio).normalize()]
        if fecha_fin is not None:
            df = df[df['Fecha'] <= pd.Timestamp(fecha_fin).normalize()]

        filas = []
        for _, fila in df.iterrows():
            registro = {'Fecha': fila['Fecha'], 'Maquina': fila['Maquina']}
            registro.update(fila['Detalle_Cargables'] or {})
            registro['Perforacion Normal'] = fila['Horas Perforacion Normal']
            registro.update(fila['Detalle_No_Cargables'] or {})
            registro['Total'] = fila['Horas Paradas'] + fila['Horas Operativas']
            filas.append(registro)

        if not filas:
            return pd.DataFrame(columns=['Fecha', 'Maquina', 'Total'])

        reporte = pd.DataFrame(filas)
        # "Total" se agrega al final de cada registro, pero si una fila posterior introduce
        # una categoría nueva que las anteriores no tenían (ej. "Otros (sin clasificar)"),
        # pd.DataFrame la agrega DESPUÉS de "Total" por orden de aparición. Se reordena
        # explícitamente para que "Total" quede siempre como última columna.
        columnas_categoria = [c for c in reporte.columns if c not in ('Fecha', 'Maquina', 'Total')]
        reporte[columnas_categoria + ['Total']] = reporte[columnas_categoria + ['Total']].fillna(0.0)
        reporte = reporte.groupby(['Fecha', 'Maquina'], as_index=False)[columnas_categoria + ['Total']].sum()
        reporte = reporte[['Fecha', 'Maquina'] + columnas_categoria + ['Total']]
        return reporte.sort_values(['Fecha', 'Maquina']).reset_index(drop=True)

    def _build_reporte_cargables_paradas(self, fecha_inicio=None, fecha_fin=None):
        """Versión reducida de _build_reporte_categorias_paradas: solo las categorías de
        Horas Cargables (N14:N28/Q14:Q28) más su subtotal, una fila por Fecha+Máquina —
        la hoja "solo cargables con el detalle" pedida para el Excel exportable."""
        df = self.df.copy()
        df['Fecha'] = pd.to_datetime(df['Fecha']).dt.normalize()
        if fecha_inicio is not None:
            df = df[df['Fecha'] >= pd.Timestamp(fecha_inicio).normalize()]
        if fecha_fin is not None:
            df = df[df['Fecha'] <= pd.Timestamp(fecha_fin).normalize()]

        filas = []
        for _, fila in df.iterrows():
            registro = {'Fecha': fila['Fecha'], 'Maquina': fila['Maquina']}
            registro.update(fila['Detalle_Cargables'] or {})
            registro['Total Cargables'] = fila['Horas Paradas']
            filas.append(registro)

        if not filas:
            return pd.DataFrame(columns=['Fecha', 'Maquina', 'Total Cargables'])

        reporte = pd.DataFrame(filas)
        columnas_categoria = [c for c in reporte.columns if c not in ('Fecha', 'Maquina', 'Total Cargables')]
        reporte[columnas_categoria + ['Total Cargables']] = reporte[columnas_categoria + ['Total Cargables']].fillna(0.0)
        reporte = reporte.groupby(['Fecha', 'Maquina'], as_index=False)[columnas_categoria + ['Total Cargables']].sum()
        reporte = reporte[['Fecha', 'Maquina'] + columnas_categoria + ['Total Cargables']]
        return reporte.sort_values(['Fecha', 'Maquina']).reset_index(drop=True)

    def _mes_personalizado(self, fecha):
        """Periodo mensual 'operativo' del proyecto: cada mes va del día 20 del mes anterior
        al día 19 del mes actual (ej. Marzo = 20-Feb a 19-Mar), EXCEPTO en el cruce de año:
        Diciembre se extiende hasta el 31 (20-Nov a 31-Dic, no rueda a Enero del año
        siguiente) y por eso Enero empieza limpio el 01 (01-Ene a 19-Ene). Devuelve un
        pd.Period mensual (año-mes) que se usa como llave de agrupación."""
        fecha = pd.Timestamp(fecha)
        if fecha.month == 12 and fecha.day >= 20:
            return pd.Period(year=fecha.year, month=12, freq="M")
        if fecha.day >= 20:
            return (fecha + pd.DateOffset(months=1)).to_period("M")
        return fecha.to_period("M")

    def _distribuir_en_buckets(self, desde, hasta, ancho_bucket=100, num_buckets=5):
        """Reparte los metros perforados en un turno (de `desde` a `hasta`, metraje
        ACUMULADO de esa broca específica) entre los tramos de `ancho_bucket` metros de su
        vida útil ("0-100", "101-200", ..., y ">500" para lo que exceda `num_buckets`),
        prorrateando cuando el rango de un turno cruza el límite entre dos tramos (ej. de
        95 a 110: 5 m para "0-100" y 5 m para "101-200")."""
        resultado = {}
        if pd.isna(desde) or pd.isna(hasta) or hasta <= desde:
            return resultado
        cursor = float(desde)
        fin = float(hasta)
        while cursor < fin:
            idx_bucket = int(cursor // ancho_bucket)
            limite_bucket = (idx_bucket + 1) * ancho_bucket
            siguiente = min(fin, limite_bucket)
            metros = siguiente - cursor
            if idx_bucket < num_buckets:
                nombre = "0-100" if idx_bucket == 0 else f"{idx_bucket * ancho_bucket + 1}-{limite_bucket}"
            else:
                nombre = f">{num_buckets * ancho_bucket}"
            resultado[nombre] = resultado.get(nombre, 0.0) + metros
            cursor = siguiente
        return resultado

    def _build_resumen_brocas(self, fecha_inicio=None, fecha_fin=None):
        """Resumen de metros perforados por diámetro de broca (HQ/NQ/PQ/...) y tramo de vida
        útil (0-100, 101-200, ..., >500), día a día: cada turno se atribuye al tramo en que
        cayó el metraje ACUMULADO de esa broca (self.df['Detalle_Brocas']), prorrateando si
        el turno cruzó un límite de tramo (ver _distribuir_en_buckets). El tricono
        (self.df['Detalle_Tricono']) no se desglosa por tramo, solo se suma como categoría
        "Tricono". Devuelve un DataFrame largo: Fecha, Maquina, Mes_Periodo, Categoria,
        Metros — listo para pivotear por máquina/global y por mes."""
        df = self.df.copy()
        df['Fecha'] = pd.to_datetime(df['Fecha']).dt.normalize()
        if fecha_inicio is not None:
            df = df[df['Fecha'] >= pd.Timestamp(fecha_inicio).normalize()]
        if fecha_fin is not None:
            df = df[df['Fecha'] <= pd.Timestamp(fecha_fin).normalize()]

        filas = []
        for _, fila in df.iterrows():
            mes_periodo = self._mes_personalizado(fila['Fecha'])
            for broca in (fila['Detalle_Brocas'] or []):
                buckets = self._distribuir_en_buckets(broca['desde'], broca['hasta'])
                for tramo, metros in buckets.items():
                    if metros <= 0:
                        continue
                    filas.append({
                        'Fecha': fila['Fecha'], 'Maquina': fila['Maquina'], 'Mes_Periodo': mes_periodo,
                        'Categoria': f"{broca['diametro']} {tramo}", 'Metros': metros,
                    })
            for tricono in (fila['Detalle_Tricono'] or []):
                metros = tricono['total']
                if pd.isna(metros) and pd.notna(tricono['desde']) and pd.notna(tricono['hasta']):
                    metros = tricono['hasta'] - tricono['desde']
                if pd.isna(metros) or metros <= 0:
                    continue
                filas.append({
                    'Fecha': fila['Fecha'], 'Maquina': fila['Maquina'], 'Mes_Periodo': mes_periodo,
                    'Categoria': 'Tricono', 'Metros': metros,
                })

        columnas = ['Fecha', 'Maquina', 'Mes_Periodo', 'Categoria', 'Metros']
        if not filas:
            return pd.DataFrame(columns=columnas)
        return pd.DataFrame(filas)[columnas]

    def _orden_categorias_brocas(self, categorias):
        """Ordena las categorías de brocas para las tablas/gráficos: primero por diámetro
        (alfabético) y dentro de cada diámetro por tramo de vida útil ascendente, con
        "Tricono" siempre al final (fila 2.11 en la plantilla de referencia)."""
        def clave(categoria):
            if categoria == 'Tricono':
                return (1, '', 0)
            diametro, _, tramo = categoria.rpartition(' ')
            inicio_tramo = int(tramo.split('-')[0].lstrip('>') or 0)
            return (0, diametro, inicio_tramo)
        return sorted(categorias, key=clave)

    def _tabla_brocas_pivote(self, resumen_brocas, maquina=None):
        """Tabla con una fila por categoría (diámetro+tramo, más Tricono) y una columna por
        mes operativo (ver _mes_personalizado), igual a la plantilla de referencia. Si
        `maquina` es None, agrega TODAS las máquinas (tabla "Global")."""
        datos = resumen_brocas if maquina is None else resumen_brocas[resumen_brocas['Maquina'] == maquina]
        if datos.empty:
            return pd.DataFrame()

        tabla = datos.groupby(['Categoria', 'Mes_Periodo'])['Metros'].sum().unstack('Mes_Periodo', fill_value=0.0)
        tabla = tabla.reindex(self._orden_categorias_brocas(tabla.index), axis=0)
        tabla = tabla.reindex(sorted(tabla.columns), axis=1)
        tabla.columns = [p.strftime('%b-%y') for p in tabla.columns]
        tabla['Total'] = tabla.sum(axis=1)
        return tabla

    def _draw_brocas_bar_figure(self, resumen_brocas, figsize=(12, 6)):
        """Barras de metros totales por categoría (diámetro+tramo, más Tricono), coloreadas
        por diámetro de broca — vista global (todas las máquinas, todo el periodo cargado)."""
        totales = resumen_brocas.groupby('Categoria')['Metros'].sum()
        categorias = self._orden_categorias_brocas(totales.index)
        valores = [totales[c] for c in categorias]
        colores_diametro = {'HQ': '#5B84C1', 'NQ': '#2E7D32', 'PQ': '#C62828', 'Tricono': '#7A7A7A'}
        colores = [colores_diametro.get(c.split(' ')[0] if c != 'Tricono' else 'Tricono', '#F09440')
                   for c in categorias]

        fig, ax = plt.subplots(figsize=figsize)
        barras = ax.bar(categorias, valores, color=colores, width=0.65)
        for barra, valor in zip(barras, valores):
            if valor > 0:
                ax.text(barra.get_x() + barra.get_width() / 2, valor, f'{valor:,.0f}',
                        ha='center', va='bottom', fontsize=8)

        ax.set_title('Metros Perforados por Diámetro de Broca y Tramo de Vida Útil',
                     fontweight='bold', fontsize=13)
        ax.set_ylabel('Metros')
        plt.setp(ax.get_xticklabels(), rotation=40, ha='right', fontsize=8)
        ax.grid(axis='y', linestyle='--', alpha=0.4)
        for s in ['top', 'right']:
            ax.spines[s].set_visible(False)

        legend_elements = [patches.Patch(facecolor=color, label=diametro)
                            for diametro, color in colores_diametro.items()]
        ax.legend(handles=legend_elements, loc='upper right', frameon=False, fontsize=9)
        fig.tight_layout()
        return fig

    def _render_brocas_bar_png(self, resumen_brocas):
        fig = self._draw_brocas_bar_figure(resumen_brocas)
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def _figsize_estatico(self, num_subplots, alto_por_subplot=1.6, alto_base=1.0, alto_maximo=9.0, ancho=9):
        """Tamaño de figura para las imágenes ESTÁTICAS del reporte Word: crece con la cantidad
        de subplots pero con un tope, para que un gráfico con muchos subplots (5 máquinas, 6
        etapas, etc.) no genere una imagen desproporcionadamente alta que en Word termina
        ocupando una hoja casi en blanco al escalarse al ancho de página."""
        alto = min(alto_por_subplot * num_subplots + alto_base, alto_maximo)
        return (ancho, alto)

    def _periodos_comunes(self, resumen, col_periodo='Semana_Label', meses_recientes=2, max_sin_filtrar=10):
        """Lista de periodos (columna `col_periodo` + Eje_X) únicos de `resumen`, para usar
        como eje X común entre subplots de distintas categorías (máquinas, etapas de almacén,
        etc.): así la misma columna siempre corresponde al mismo periodo calendario, en vez de
        que cada categoría muestre solo sus propios periodos (lo que antes daba la falsa
        impresión de que todas operaron todo el tiempo). Si hay muchos periodos en total y
        `meses_recientes` no es None, se limita a los últimos `meses_recientes` para que las
        barras sigan siendo legibles; con `meses_recientes=None` no se recorta (útil para vistas
        mensuales, que ya tienen pocos periodos)."""
        periodos = (resumen[[col_periodo, 'Eje_X']]
                    .drop_duplicates()
                    .sort_values(col_periodo))
        if meses_recientes is not None and len(periodos) > max_sin_filtrar:
            fecha_corte = periodos[col_periodo].dt.end_time.max() - pd.DateOffset(months=meses_recientes)
            periodos = periodos[periodos[col_periodo].dt.end_time > fecha_corte]
        return periodos

    def _draw_paradas_figure(self, resumen_paradas, figsize):
        """Dibuja un subplot de Horas Operativas vs Paradas + % Eficiencia por cada máquina
        (mismo patrón visual que PlotDash). Reutilizado por PlotParadas (interactivo) y
        _render_paradas_png (versión estática para el reporte Word)."""
        maquinas = self._maquinas_validas(resumen_paradas)
        num_maquinas = len(maquinas)
        fig, axes = plt.subplots(num_maquinas, 1, figsize=figsize, sharex=True)
        if num_maquinas == 1:
            axes = [axes]

        # Eje X común a TODAS las máquinas (ver _periodos_comunes). Con sharex=True +
        # label_outer(), las etiquetas de semana solo se dibujan una vez, en el subplot de
        # abajo, y todas las máquinas quedan alineadas bajo esas mismas columnas.
        semanas_comunes = self._periodos_comunes(resumen_paradas)

        for ax, maquina in zip(axes, maquinas):
            datos_m = semanas_comunes.merge(
                resumen_paradas[resumen_paradas['Maquina'] == maquina],
                on=['Semana_Label', 'Eje_X'], how='left'
            )
            ax.bar(datos_m['Eje_X'], datos_m['Horas_Operativas'], color='#2E7D32', width=0.6)
            ax.bar(datos_m['Eje_X'], datos_m['Horas_Paradas'],
                   bottom=datos_m['Horas_Operativas'], color='#C62828', width=0.6)

            ax2 = ax.twinx()
            ax2.plot(datos_m['Eje_X'], datos_m['Pct_Eficiencia'], color='#F09440', linewidth=2, marker='o')
            ax2.set_ylim(0, 110)
            ax2.yaxis.set_major_formatter(ticker.PercentFormatter())

            ax.set_title(f'Máquina: {maquina}', fontweight='bold', loc='left', fontsize=10)
            ax.set_ylabel('Horas', fontsize=9)
            ax.grid(axis='y', linestyle='--', alpha=0.4)
            ax.label_outer()

        plt.setp(axes[-1].get_xticklabels(), rotation=40, ha='right', fontsize=7)

        legend_elements = [
            patches.Patch(facecolor='#2E7D32', label='Horas Operativas'),
            patches.Patch(facecolor='#C62828', label='Horas Paradas'),
            Line2D([0], [0], color='#F09440', lw=2, marker='o', label='% Eficiencia'),
        ]
        fig.suptitle('Resumen Semanal - Horas Operativas vs Paradas por Máquina',
                     fontsize=14, fontweight='bold', y=1.08)
        fig.legend(handles=legend_elements, loc='upper center', bbox_to_anchor=(0.5, 1.02),
                   ncol=3, frameon=False, fontsize=10)
        fig.tight_layout()
        return fig

    def _build_resumen_disponibilidad(self):
        """Disponibilidad REAL de cada máquina: horas operativas reportadas frente al total de
        horas de calendario posibles en cada semana (días dentro de su ventana operativa ×
        24 h), no solo frente a las horas que llegaron a reportarse. El % de 'Eficiencia' de
        _build_resumen_paradas solo compara Operativas vs Paradas DENTRO de lo reportado, así
        que un día sin ningún turno reportado (parada no registrada) no lo penaliza — por eso
        da porcentajes altos (>90%) que no reflejan la disponibilidad real de la máquina."""
        HORAS_POR_DIA = 24
        fecha_max_global = self.df["Fecha"].max()

        df_horas = self.df.copy()
        df_horas['Fecha'] = pd.to_datetime(df_horas['Fecha'])
        df_horas['Semana_Label'] = df_horas['Fecha'].dt.to_period('W-TUE')
        horas_reportadas = df_horas.groupby(['Semana_Label', 'Maquina']).agg(
            Horas_Operativas=('Horas Operativas', 'sum'),
            Horas_Paradas_Reportadas=('Horas Paradas', 'sum'),
        ).reset_index()

        filas_calendario = []
        for maquina, cfg in self.machine_lifecycle.items():
            fin = cfg['end'] if cfg['end'] is not None else fecha_max_global
            if fin < cfg['start']:
                continue
            dias = pd.Series(pd.date_range(cfg['start'], fin))
            semanas = dias.dt.to_period('W-TUE')
            for semana, grupo_dias in dias.groupby(semanas):
                filas_calendario.append({
                    'Semana_Label': semana,
                    'Maquina': maquina,
                    'Fecha_Inicio': grupo_dias.min(),
                    'Fecha_Fin': grupo_dias.max(),
                    'Horas_Calendario': len(grupo_dias) * HORAS_POR_DIA,
                })
        calendario = pd.DataFrame(filas_calendario)

        resumen = calendario.merge(horas_reportadas, on=['Semana_Label', 'Maquina'], how='left')
        resumen[['Horas_Operativas', 'Horas_Paradas_Reportadas']] = \
            resumen[['Horas_Operativas', 'Horas_Paradas_Reportadas']].fillna(0)
        resumen['Pct_Disponibilidad'] = np.where(
            resumen['Horas_Calendario'] > 0,
            resumen['Horas_Operativas'] / resumen['Horas_Calendario'] * 100,
            np.nan
        )
        # Eje_X se deriva de los límites propios de la semana (miércoles a martes), no de
        # Fecha_Inicio/Fecha_Fin (que sí quedan recortadas a la ventana operativa de cada
        # máquina para calcular Horas_Calendario correctamente, pero no deben usarse para la
        # etiqueta visible: dos máquinas en la misma semana calendario deben mostrar el mismo
        # texto de columna para poder alinearse entre subplots).
        resumen['Eje_X'] = (resumen['Semana_Label'].dt.start_time.dt.strftime('%d%b') + "-" +
                            resumen['Semana_Label'].dt.end_time.dt.strftime('%d%b'))
        return resumen

    def _draw_disponibilidad_figure(self, resumen_disp, figsize):
        """Un subplot por máquina con el % de disponibilidad real semanal (horas operativas
        sobre horas de calendario posibles). Las barras se colorean por severidad para que los
        porcentajes bajos, típicos de este indicador, se lean de un vistazo."""
        maquinas = self._maquinas_validas(resumen_disp)
        num_maquinas = len(maquinas)
        fig, axes = plt.subplots(num_maquinas, 1, figsize=figsize, sharex=True)
        if num_maquinas == 1:
            axes = [axes]

        def color_por_valor(v):
            if pd.isna(v): return '#CCCCCC'
            if v < 40: return '#C62828'
            if v < 65: return '#F09440'
            return '#2E7D32'

        # Eje X común a todas las máquinas (ver _periodos_comunes). Con sharex=True +
        # label_outer(), las etiquetas de semana solo se dibujan una vez, en el subplot de abajo.
        semanas_comunes = self._periodos_comunes(resumen_disp)

        for ax, maquina in zip(axes, maquinas):
            datos_m = semanas_comunes.merge(
                resumen_disp[resumen_disp['Maquina'] == maquina],
                on=['Semana_Label', 'Eje_X'], how='left'
            )
            colores = datos_m['Pct_Disponibilidad'].apply(color_por_valor)
            ax.bar(datos_m['Eje_X'], datos_m['Pct_Disponibilidad'], color=colores, width=0.6)

            for x, v in zip(datos_m['Eje_X'], datos_m['Pct_Disponibilidad']):
                if pd.notna(v):
                    ax.text(x, v + 2, f'{v:.0f}%', ha='center', va='bottom', fontsize=7)

            ax.set_title(f'Máquina: {maquina}', fontweight='bold', loc='left', fontsize=10)
            ax.set_ylabel('% Disponibilidad', fontsize=9)
            ax.set_ylim(0, 110)
            ax.grid(axis='y', linestyle='--', alpha=0.4)
            ax.label_outer()

        plt.setp(axes[-1].get_xticklabels(), rotation=40, ha='right', fontsize=7)

        legend_elements = [
            patches.Patch(facecolor='#C62828', label='< 40%'),
            patches.Patch(facecolor='#F09440', label='40% - 65%'),
            patches.Patch(facecolor='#2E7D32', label='> 65%'),
        ]
        fig.suptitle('Disponibilidad Semanal Real (Horas Operativas / Horas de Calendario)',
                     fontsize=14, fontweight='bold', y=1.09)
        fig.legend(handles=legend_elements, loc='upper center', bbox_to_anchor=(0.5, 1.02),
                   ncol=3, frameon=False, fontsize=10, title='Disponibilidad')
        fig.tight_layout()
        return fig

    def PlotS(self):
        self.Mensajes.clear()
        try:
            # =============================
            # 1. PREPARACIÓN DE DATOS
            # =============================
            # Nota: Asegúrate de que 'curva_s' sea tu DataFrame original con las columnas 'Fecha', 'Plan_Acumulado' y 'Real_Acumulado'
            df_real_valid = self.curva_s.dropna(subset=["Real_Acumulado"])
            ultima_fecha = df_real_valid["Fecha"].max()
            plan_acum_val = self.curva_s.loc[self.curva_s["Fecha"] == ultima_fecha, "Plan_Acumulado"].values[0]
            real_acum_val = self.curva_s.loc[self.curva_s["Fecha"] == ultima_fecha, "Real_Acumulado"].values[0]
            cumplimiento = (real_acum_val / plan_acum_val) * 100 if plan_acum_val > 0 else 0
            
            meses_unicos = pd.to_datetime(self.curva_s['Fecha'].dt.strftime('%Y-%m-01')).unique()
            meses_nums = np.arange(len(meses_unicos))
            
            # =============================
            # 2. CONFIGURACIÓN DE FIGURA
            # =============================
            fig, ax = plt.subplots(figsize=(15, 8))
            plt.subplots_adjust(bottom=0.25, top=0.9, left=0.1)
            
            def dibujar_grafico(idx_inicio, idx_fin):
                ax.clear()
                
                # Rango de fechas basado en el slider
                fecha_inicio = meses_unicos[int(idx_inicio)]
                fecha_fin = meses_unicos[int(idx_fin)] + pd.offsets.MonthEnd(0)
                data_plot = self.curva_s[(self.curva_s['Fecha'] >= fecha_inicio) & (self.curva_s['Fecha'] <= fecha_fin)]
            
                def plot_linea(df_sub, col_y, color, is_real=False):
                    temp = df_sub.dropna(subset=[col_y])
                    if temp.empty: return
                    
                    x_num = mdates.date2num(temp['Fecha'])
                    y_val = temp[col_y].values
                    
                    # Dibujar la línea suavizada (Spline)
                    if len(x_num) > 3:
                        x_new = np.linspace(x_num.min(), x_num.max(), 300)
                        spl = make_interp_spline(x_num, y_val, k=3)
                        ax.plot(mdates.num2date(x_new), spl(x_new), color=color, lw=2.5, alpha=0.8, zorder=2)
                    else:
                        ax.plot(temp['Fecha'], y_val, color=color, lw=2.5, alpha=0.8, zorder=2)
            
                    # Definir tamaños de puntos y colores
                    sizes = []
                    for f in temp['Fecha']:
                        if f == ultima_fecha: sizes.append(80) # Punto de corte actual (Negro)
                        elif f == temp['Fecha'].max() and not is_real: sizes.append(60) # Punto final Plan
                        elif f.weekday() == 2: sizes.append(40) # Miércoles
                        else: sizes.append(15) # Días normales
                    
                    colors = ['black' if f == ultima_fecha else color for f in temp['Fecha']]
                    ax.scatter(temp['Fecha'], y_val, s=sizes, c=colors, edgecolors='white', lw=0.6, zorder=3)
            
                    # Etiquetas de valores numéricos
                    max_fecha_plan = temp['Fecha'].max()
                    for i, row in temp.iterrows():
                        f, v = row['Fecha'], row[col_y]
                        
                        # Etiqueta de avance real actual (punto negro)
                        if f == ultima_fecha:
                            offset = -1400 if not is_real else 1000
                            lbl_color = 'blue' if not is_real else 'red'
                            ax.text(f, v + offset, f'{v:,.0f}', color=lbl_color, fontweight='bold', ha='center', fontsize=11)
                        
                        # Etiqueta del valor final (Ej. 28,000)
                        elif f == max_fecha_plan and not is_real:
                            ax.text(f, v + 900, f'{v:,.0f}', color='blue', fontweight='bold', ha='center', fontsize=9)
                        
                        # Hitos semanales (Miércoles)
                        elif f.weekday() == 2:
                            ax.text(f, v + 800, f'{v:,.0f}', color='black', fontsize=8, ha='center', alpha=0.7)
            
                # Dibujar ambas curvas
                plot_linea(data_plot, "Plan_Acumulado", "blue")
                plot_linea(data_plot, "Real_Acumulado", "#d62728", is_real=True)
            
                # --- Cuadro Resumen (Avance) ---
                texto_box = (
                    f"$\mathbf{{Avance \ ({ultima_fecha.strftime('%d/%m/%y')})}}$\n"
                    f"Plan Acum.: {plan_acum_val:,.0f} mts\n"
                    f"Real Acum.: {real_acum_val:,.0f} mts\n"
                    f"Cumplimiento: {cumplimiento:.1f}%"
                )
                ax.annotate(texto_box, xy=(0.98, 0.05), xycoords='axes fraction',
                            fontsize=11, ha='right', va='bottom', multialignment='left',
                            bbox=dict(boxstyle="round,pad=0.8", facecolor="white", edgecolor="black", lw=1))
            
                # --- AJUSTES DE EJES Y TÍTULOS ---
                ax.set_title("Curva S — Avance de Perforación Diamantina", loc='center', fontweight='bold', fontsize=16, pad=20)
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m-%Y'))
                plt.xticks(rotation=45)
                
                ax.set_ylabel("Metros perforados", fontweight='bold', fontsize=12)
                # Límite en 31k para que el texto de 28k no se corte
                ax.set_ylim(-500, 31000) 
                ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
                
                ax.grid(True, which='major', color='#dddddd', linestyle='--', alpha=0.7)
                for s in ['top', 'right']: ax.spines[s].set_visible(False)
            
                # Leyenda personalizada
                legend_elements = [
                    Line2D([0], [0], color='blue', lw=2, label='Plan acumulado', marker='o', markersize=5),
                    Line2D([0], [0], color='#d62728', lw=2, label='Real acumulado', marker='o', markersize=5)
                ]
                ax.legend(handles=legend_elements, loc='upper left', frameon=False, fontsize=10)
                
                fig.canvas.draw_idle()
            
            # =============================
            # 3. RANGE SLIDER
            # =============================
            ax_slider = plt.axes([0.20, 0.05, 0.60, 0.03], facecolor='#f0f0f0')
            slider_meses = RangeSlider(ax_slider, 'Rango Meses ', meses_nums[0], meses_nums[-1], 
                                      valinit=(meses_nums[0], meses_nums[-1]), valstep=1)
            
            def actualizar_labels_slider(val):
                inicio_str = meses_unicos[int(val[0])].strftime('%b %y')
                fin_str = meses_unicos[int(val[1])].strftime('%b %y')
                slider_meses.valtext.set_text(f'{inicio_str} - {fin_str}')
                dibujar_grafico(val[0], val[1])
            
            slider_meses.on_changed(actualizar_labels_slider)
            
            # Inicializar gráfico
            actualizar_labels_slider((meses_nums[0], meses_nums[-1]))
            
            plt.show()
            
            
            # =============================
            # 1. PREPARACIÓN DE DATOS
            # =============================
            META_TOTAL = META_TOTAL_PROYECTO
            
            df_real_valid = self.curva_s.dropna(subset=["Real_Acumulado"])
            ultima_fecha = df_real_valid["Fecha"].max()
            
            # Valores actuales para el cuadro de resumen
            plan_acum_val = self.curva_s.loc[self.curva_s["Fecha"] == ultima_fecha, "Plan_Acumulado"].values[0]
            real_acum_val = self.curva_s.loc[self.curva_s["Fecha"] == ultima_fecha, "Real_Acumulado"].values[0]
            cumplimiento = (real_acum_val / plan_acum_val) * 100 if plan_acum_val > 0 else 0
            
            meses_unicos = pd.to_datetime(self.curva_s['Fecha'].dt.strftime('%Y-%m-01')).unique()
            meses_nums = np.arange(len(meses_unicos))
            
            # =============================
            # 2. CONFIGURACIÓN DE FIGURA
            # =============================
            fig, ax = plt.subplots(figsize=(15, 8))
            plt.subplots_adjust(bottom=0.25, top=0.9, left=0.1)
            
            def dibujar_grafico_porcentaje(idx_inicio, idx_fin):
                ax.clear()
                
                fecha_inicio = meses_unicos[int(idx_inicio)]
                fecha_fin = meses_unicos[int(idx_fin)] + pd.offsets.MonthEnd(0)
                data_plot = self.curva_s[(self.curva_s['Fecha'] >= fecha_inicio) & (self.curva_s['Fecha'] <= fecha_fin)]
            
                def plot_linea_pct(df_sub, col_y, color, is_real=False):
                    temp = df_sub.dropna(subset=[col_y])
                    if temp.empty: return
                    
                    x_num = mdates.date2num(temp['Fecha'])
                    # CONVERSIÓN A PORCENTAJE
                    y_val = (temp[col_y].values / META_TOTAL) * 100
                    
                    if len(x_num) > 3:
                        x_new = np.linspace(x_num.min(), x_num.max(), 300)
                        spl = make_interp_spline(x_num, y_val, k=3)
                        ax.plot(mdates.num2date(x_new), spl(x_new), color=color, lw=2.5, alpha=0.8, zorder=2)
                    else:
                        ax.plot(temp['Fecha'], y_val, color=color, lw=2.5, alpha=0.8, zorder=2)
            
                    # Puntos
                    sizes = [80 if f == ultima_fecha else (40 if f.weekday() == 2 else 15) for f in temp['Fecha']]
                    colors = ['black' if f == ultima_fecha else color for f in temp['Fecha']]
                    ax.scatter(temp['Fecha'], y_val, s=sizes, c=colors, edgecolors='white', lw=0.6, zorder=3)
            
                    # Etiquetas (en %)
                    for f, v in zip(temp['Fecha'], y_val):
                        if f == ultima_fecha:
                            offset = -5 if not is_real else 4
                            lbl_color = 'blue' if not is_real else 'red'
                            ax.text(f, v + offset, f'{v:.1f}%', color=lbl_color, fontweight='bold', ha='center', fontsize=10)
                        elif f.weekday() == 2 or f == temp['Fecha'].max():
                            ax.text(f, v + 3, f'{v:.0f}%', color='black', fontsize=8, ha='center', alpha=0.7)
            
                plot_linea_pct(data_plot, "Plan_Acumulado", "blue")
                plot_linea_pct(data_plot, "Real_Acumulado", "#d62728", is_real=True)
            
                # --- Cuadro Resumen ---
                texto_box = (
                    f"$\mathbf{{Estado \ al \ {ultima_fecha.strftime('%d/%m/%y')}}}$\n"
                    f"Plan Acum.: {plan_acum_val:,.0f} mts\n"
                    f"Real Acum.: {real_acum_val:,.0f} mts\n"
                    f"Cumplimiento: {cumplimiento:.1f}%"
                )
                ax.annotate(texto_box, xy=(0.98, 0.05), xycoords='axes fraction',
                            fontsize=11, ha='right', va='bottom',
                            bbox=dict(boxstyle="round,pad=0.8", facecolor="white", edgecolor="black", lw=1))
            
                # --- AJUSTES DE EJES ---
                ax.set_title(f"Curva S — Porcentaje de Avance (Meta: {META_TOTAL:,.0f} mts)", fontweight='bold', fontsize=16, pad=20)
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m-%Y'))
                plt.xticks(rotation=45)
                
                ax.set_ylabel("Porcentaje de Avance (%)", fontweight='bold', fontsize=12)
                ax.set_ylim(-5, 110) # Límite hasta 110% para visualizar etiquetas
                ax.yaxis.set_major_formatter(ticker.PercentFormatter()) # Formato 0%
                
                ax.grid(True, which='major', color='#dddddd', linestyle='--', alpha=0.7)
                for s in ['top', 'right']: ax.spines[s].set_visible(False)
            
                legend_elements = [
                    Line2D([0], [0], color='blue', lw=2, label='% Plan acumulado'),
                    Line2D([0], [0], color='#d62728', lw=2, label='% Real acumulado')
                ]
                ax.legend(handles=legend_elements, loc='upper left', frameon=False)
                
                fig.canvas.draw_idle()
            
            # =============================
            # 3. RANGE SLIDER
            # =============================
            ax_slider = plt.axes([0.20, 0.05, 0.60, 0.03], facecolor='#f0f0f0')
            slider_meses = RangeSlider(ax_slider, 'Rango Meses ', meses_nums[0], meses_nums[-1], 
                                      valinit=(meses_nums[0], meses_nums[-1]), valstep=1)
            
            def actualizar(val):
                inicio_str = meses_unicos[int(val[0])].strftime('%b %y')
                fin_str = meses_unicos[int(val[1])].strftime('%b %y')
                slider_meses.valtext.set_text(f'{inicio_str} - {fin_str}')
                dibujar_grafico_porcentaje(val[0], val[1])
            
            slider_meses.on_changed(actualizar)
            actualizar((meses_nums[0], meses_nums[-1]))
            
            plt.show()  
            
            # Reportes semanales: ciclo de vida de máquinas (automático) y meta contractual dinámica
            df_clean, machine_cols = self._build_df_clean()
            resumen_semanal_1, resumen_semanal_2 = self._build_resumen_semanal(df_clean, machine_cols)

            # Extraer valor de la meta actual (Última semana)
            meta_actual_valor = resumen_semanal_1['Meta_Semanal'].iloc[-1]
            
            # Generación del Gráfico 1
            plt.figure(figsize=(15, 8))
            
            # Barras de producción real acumulada
            ax1 = plt.bar(resumen_semanal_1['Eje_X'], resumen_semanal_1['Suma_Semanal'], color='#5B84C1', width=0.6, label='Suma Semanal')
            
            # Línea de Meta Contractual
            plt.plot(resumen_semanal_1['Eje_X'], resumen_semanal_1['Meta_Semanal'], color='#B04848', linewidth=3, 
                     marker='o', label=f'Objetivo Semanal actual ({int(meta_actual_valor)} m)')
            
            # Línea de Media Total Real obtenida
            media_total_1 = resumen_semanal_1['Suma_Semanal'].mean()
            plt.axhline(y=media_total_1, color='#F09440', linewidth=3, label=f'Media Semanal Real ({media_total_1:.1f} m)')
            
            # Etiquetas de datos sobre las barras
            for bar in ax1:
                yval = bar.get_height()
                if not np.isnan(yval):
                    plt.text(bar.get_x() + bar.get_width()/2, yval + 10, f'{yval:.1f}', ha='center', va='bottom', fontsize=10, 
                             bbox=dict(facecolor='white', edgecolor='#cccccc', boxstyle='round,pad=0.2'))
            
            # Configuración de Ejes y Títulos
            plt.title('Resumen Semanal - Metros Perforados (Suma Total vs Meta Contractual)', fontsize=16, fontweight='bold', pad=25, color='#444444')
            plt.xticks(rotation=40, ha='right')
            plt.ylabel('METRAJE SEMANAL (MTS)', fontsize=11, fontweight='bold')
            plt.xlabel('FECHA SEMANAL', fontsize=11, fontweight='bold')
            
            plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False, fontsize=11)
            plt.grid(axis='y', linestyle='--', alpha=0.4)
            plt.tight_layout()
            plt.show()
            
            
            # =========================================================================
            # ## REPORTE SEMANAL 2: PROMEDIO POR MÁQUINA (ya calculado en _build_resumen_semanal)
            # =========================================================================
            # Generación del Gráfico 2
            plt.figure(figsize=(15, 8))
            ax2 = resumen_semanal_2['Promedio_Semanal'].plot(kind='bar', color='#5B84C1', width=0.6, label='Promedio Semanal')
            
            objetivo_diario = 35
            media_total_2 = resumen_semanal_2['Promedio_Semanal'].mean()
            
            plt.axhline(y=objetivo_diario, color='#B04848', linewidth=3, label=f'Objetivo Diario ({objetivo_diario} m)')
            plt.axhline(y=media_total_2, color='#F09440', linewidth=3, label=f'Media Total ({media_total_2:.2f} m)')
            
            # Etiquetas de datos sobre las barras
            for i, v in enumerate(resumen_semanal_2['Promedio_Semanal']):
                if not np.isnan(v):
                    ax2.text(i, v + 0.5, f'{v:.1f}', ha='center', va='bottom', fontsize=10, 
                             bbox=dict(facecolor='white', edgecolor='#cccccc', boxstyle='round,pad=0.2'))
            
            plt.title('Resumen Semanal - Promedio Perforados Semanal-Maquina', fontsize=16, fontweight='bold', pad=25, color='#444444')
            plt.xticks(range(len(resumen_semanal_2)), resumen_semanal_2['Eje_X'], rotation=40, ha='right')
            plt.ylabel('METRAJE PROMEDIO SEMANAL (MTS)', fontsize=11, fontweight='bold')
            plt.xlabel('FECHA SEMANAL', fontsize=11, fontweight='bold')
            plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False, fontsize=11)
            plt.grid(axis='y', linestyle='--', alpha=0.4)
            plt.tight_layout()
            plt.show()
            self.Mensajes.setText("Curva S Ploteada")
            self.PlanPerf.setEnabled(False)
            self.CtrlProceso.setEnabled(False)

        except Exception as e:
            print(f"Error al plotear revisa los archivos {e}")
         
    def LoadProc(self):
            self.Mensajes.clear()
            # 1. Seleccionar el archivo de Control de Proceso
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Seleccionar Control de Proceso (DDH Avance Plan)", "", "Excel Files (*.xlsx; *.xlsm)"
            )
            
            if file_path:
                try:
                    # --- PROCESAMIENTO DE EQUIVALENCIAS ---
                    # sheet_name="Control de Proceso", fila 3 (header=2)
                    # usecols="D,E,AE,AF,AG" -> Equivalen a índices o nombres
                    Equivalencias = pd.read_excel(
                        file_path,
                        sheet_name="Control de Proceso",
                        header=2,
                        usecols="B,D,E,Y,Z,AE,AF,AG,BD" 
                    )
                    
                    # Limpiar espacios en nombres de columnas por seguridad
                    #Equivalencias.columns = Equivalencias.columns.astype(str).str.strip()
                    
                    # Renombrar columnas según tu lógica
                    Equivalencias = Equivalencias.rename(columns={
                        "IDENTIFICADOR": "Identificador", # Columna B
                        "ESTE.2": "Este_Rep",
                        "NORTE.2": "Norte_Rep",
                        "COTA.2": "Cota_Rep",
                        "PLATAFORMA": "Plataforma",
                        "DDH": "Sondaje",
                        "INICIO": "Inicio",             # Columna Y
                        "FIN": "Fin",                   # Columna Z
                        "METRAJE DESCONTADO": "Metraje_Descontado" # Columna BD
                    })

    
                    # --- PROCESAMIENTO DE BD_TOTAL ---
                    # sheet_name="DDH", fila 2 (header=1)
                    BD_Total = pd.read_excel(
                        file_path,
                        sheet_name="DDH",
                        header=1,
                        usecols="A,B,F,G,H,I,V"
                    )
                    
                    #BD_Total.columns = BD_Total.columns.astype(str).str.strip()
                    BD_Total = BD_Total.rename(columns={BD_Total.columns[0]: "Identificador"})
                    BD_Union = pd.merge(
                        BD_Total,
                        Equivalencias,
                        on="Identificador",
                        how="left",
                        suffixes=('', '_Equiv') # Maneja columnas duplicadas si las hay
                    )
                    
                    # Opcional: Limpiar columnas repetidas de plataforma si existen en ambas hojas
                    if 'Plataforma_Equiv' in BD_Union.columns:
                        BD_Union = BD_Union.drop(columns=['Plataforma_Equiv'])
                    self.BD_Union = BD_Union
                    # Resultado final
                    print(BD_Union.head())
                    
                    # Éxito
                    print("--- BD_Union Cargada Correctamente ---")
                    print(self.BD_Union.head())
                    
                    resumen = self.df.groupby("Sondaje").agg(Maquina=("Maquina", "last"), Avance_Actual=("Avance Turno", "sum")).reset_index()
                    resumen["Sondaje"] = resumen["Sondaje"].replace("MR25015-11", "MR26015-11")
                    self.BD_Union_total = pd.merge(self.BD_Union, resumen, on="Sondaje", how="left")
                    self.BD_Union_total["Metro_Faltante"] = (self.BD_Union_total["LENGTH"] - self.BD_Union_total["Avance_Actual"]).clip(lower=0)
                    self.BD_Union_total["Estado"] = self.BD_Union_total["Status"].fillna("Programado")
                    
                    # 2.1 Cálculos de Proyección
                    avance_diario = self.df.groupby(["Sondaje", "Fecha"])["Avance Turno"].sum().reset_index()
                    promedio_2dias = avance_diario.sort_values(["Sondaje", "Fecha"]).groupby("Sondaje").tail(2).groupby("Sondaje")["Avance Turno"].mean().rename("Avance_2dias_Prom").reset_index()
                    dias_perforando = avance_diario.groupby("Sondaje")["Fecha"].agg(lambda x: (x.max() - x.min()).days + 1).rename("Dias_Perforando_Count").reset_index()
                    
                    self.BD_Union_total = self.BD_Union_total.merge(promedio_2dias, on="Sondaje", how="left").merge(dias_perforando, on="Sondaje", how="left")
                    self.BD_Union_total["Proyeccion_Dias"] = self.BD_Union_total.apply(
                        lambda r: round(r["Metro_Faltante"] / r["Avance_2dias_Prom"], 1) if (r["Estado"] == "Perforando" and r.get("Avance_2dias_Prom", 0) > 0) else np.nan, axis=1
                    )
                    
                    df_plot = self.BD_Union_total.copy()
                    df_plot["Proyeccion_fmt"] = df_plot["Proyeccion_Dias"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "-")
                    df_plot["Sondaje_fmt"] = df_plot["Sondaje"].fillna("-").astype(str)
                    df_plot["Maquina_fmt"] = df_plot["Maquina"].fillna("S/M").astype(str)
                    QMessageBox.information(self, "Éxito", "Control de proceso (BD_Union) cargado y vinculado correctamente.")
                    self.Mensajes.setText("Se cargo el Archivo con el control de procesos \nrevisar que sea el actualizado")
                    self.PlanPerf.setEnabled(True)

    
                except Exception as e:
                    # --- MANEJO DE ERRORES CON TRACEBACK ---
                    print("\n" + "!"*50)
                    print("ERROR EN LOADPROC:")
                    traceback.print_exc()
                    print("!"*50 + "\n")
    
                    error_type = type(e).__name__
                    mensaje = (f"Ocurrió un error al procesar el archivo:\n\n"
                               f"Tipo: {error_type}\n"
                               f"Detalle: {str(e)}\n\n"
                               f"Verifique que las hojas 'Control de Proceso' y 'DDH' existan y tengan el formato correcto.")
                    QMessageBox.critical(self, "Error de Procesamiento", mensaje)
                    self.Mensajes.setText(mensaje)
        
    def LoadTopo(self):
            self.Mensajes.clear()
            """
            Selecciona una carpeta, valida la presencia de archivos DXF y 
            guarda la ruta para su posterior uso en PlotPlano.
            """
            dir_path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta con archivos DXF")
            
            if not dir_path:
                return
    
            try:
                # 1. Definir la ruta y buscar archivos DXF (ignorando mayúsculas/minúsculas)
                ruta = Path(dir_path)
                archivos_dxf = list(ruta.glob("*.[dD][xX][fF]"))
    
                if not archivos_dxf:
                    QMessageBox.warning(self, "Aviso", "No se encontraron archivos .dxf en la carpeta seleccionada.")
                    return
    
                # 2. OPTIMIZACIÓN: Validamos los archivos rápidamente
                # En lugar de abrirlos todos aquí (que puede ser lento), 
                # verificamos que sean legibles y guardamos la lista.
                self.datos_topografia = []
                for archivo in archivos_dxf:
                    # Una validación rápida: ¿Es un archivo real y no está vacío?
                    if archivo.is_file() and archivo.stat().st_size > 0:
                        self.datos_topografia.append(archivo)
    
                # 3. ASIGNACIÓN CRÍTICA PARA PlotPlano
                # Guardamos la ruta como string para compatibilidad con os.path y ezdxf
                self.ruta_topo = str(ruta)
    
                QMessageBox.information(
                    self, 
                    "Éxito", 
                    f"Carpeta vinculada correctamente.\nArchivos DXF detectados: {len(self.datos_topografia)}"
                )
                
                print(f"Ruta de topografía establecida en: {self.ruta_topo}")
                self.Mensajes.setText("Carpeta con topografia cargada")
    
            except Exception as e:
                print("\n" + "!"*50)
                print("ERROR EN LOADTOPO:")
                traceback.print_exc()
                print("!"*50 + "\n")
                QMessageBox.critical(self, "Error", f"Ocurrió un fallo al procesar la carpeta: {e}")
                self.Mensajes.setText("Ocurrio algun error")
                
    def ImportLab(self):
            self.Mensajes.clear()

            # 1. Abrir el cuadro de diálogo para seleccionar UN solo archivo DXF
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                "Seleccionar archivo DXF de Labores", 
                "", 
                "Archivos DXF (*.dxf *.DXF)"
            )
            
            if not file_path:
                return  # El usuario canceló la selección
            
            try:
                ruta_archivo = Path(file_path)
                
                # 2. Validación rápida: ¿Es un archivo real y no está vacío?
                if not ruta_archivo.is_file() or ruta_archivo.stat().st_size == 0:
                    QMessageBox.warning(self, "Aviso", "El archivo seleccionado no es válido o está vacío.")
                    return
    
                # 3. ASIGNACIÓN CRÍTICA PARA PlotPlano / ezdxf
                # Guardamos la ruta absoluta del archivo como string para evitar conflictos de compatibilidad
                self.labdxf = str(ruta_archivo.resolve())
                
                # Si mantienes lógica de listas en otra sección, puedes guardar el objeto Path aquí:
                self.datos_lab = [ruta_archivo]
    
                QMessageBox.information(
                    self, 
                    "Éxito", 
                    f"Archivo DXF seleccionado correctamente:\n{ruta_archivo.name}"
                )
                
                print(f"Ruta del archivo de labores establecida en: {self.labdxf}")
                self.Mensajes.setText("Archivo de labores cargado exitosamente")
                self.PlotLab.setEnabled(True) 

    
            except Exception as e:
                print("\n" + "!"*50)
                print("ERROR EN LOADTOPO:")
                traceback.print_exc()
                print("!"*50 + "\n")
                QMessageBox.critical(self, "Error", f"Ocurrió un fallo al procesar el archivo DXF: {e}")
                self.Mensajes.setText("Ocurrió algún error al cargar el archivo")
                
    def PlotearLab(self):
            # 1. Inicializar variables de estado como ATRIBUTOS DE INSTANCIA (self)
            # Esto garantiza que no se pierdan al salir del método.
            self.ruta_dxf = self.labdxf 
            self.sondaje_actual_idx = 0
            self.indice_scroll_lista = 0
            self.max_visibles_lista = 20
            self.texto_filtro = ""  # Almacena lo que el usuario escribe en el buscador
            
            def cargar_labores_dxf(ruta_dxf):
                try:
                    doc = ezdxf.readfile(ruta_dxf)
                    msp = doc.modelspace()
                except IOError:
                    print(f"No se pudo abrir el archivo {ruta_dxf}")
                    return []
                except ezdxf.DXFStructureError:
                    print(f"Archivo DXF inválido.")
                    return []
            
                lineas_labores = []
                for entity in msp.query('LINE POLYLINE LWPOLYLINE 3DFACE'):
                    if entity.dxftype() == 'LINE':
                        p1 = np.array(entity.dxf.start)
                        p2 = np.array(entity.dxf.end)
                        lineas_labores.append((p1, p2))
                    elif entity.dxftype() in ('POLYLINE', 'LWPOLYLINE'):
                        puntos = [np.array(p) for p in entity.points()]
                        for i in range(len(puntos) - 1):
                            lineas_labores.append((puntos[i], puntos[i+1]))
                return lineas_labores
            
            def calcular_intersecciones(sondaje_inicio, sondaje_fin, lineas_labores, radio_halo=30.0):
                A = sondaje_inicio
                B = sondaje_fin
                u = B - A
                u_len = np.linalg.norm(u)
                if u_len == 0: return []
                u_hat = u / u_len
            
                labores_cercanas = []
                for p1, p2 in lineas_labores:
                    for p in (p1, p2):
                        v = p - A
                        dist_proyectada = np.dot(v, u_hat)
                        
                        if 0 <= dist_proyectada <= u_len:
                            punto_recta = A + dist_proyectada * u_hat
                            dist_perpendicular = np.linalg.norm(p - punto_recta)
                            
                            if dist_perpendicular <= radio_halo:
                                labores_cercanas.append({
                                    'punto_labor': p,
                                    'metraje_sondaje': dist_proyectada,
                                    'distancia': dist_perpendicular,
                                    'tramo': (p1, p2)
                                })
                                break 
                return labores_cercanas
            
            # Cargar los datos iniciales
            self.lineas_labores = cargar_labores_dxf(self.ruta_dxf)
            
            # =========================================================================
            # 3. ACTUALIZACIÓN DEL GRÁFICO (Modificado para usar self)
            # =========================================================================
            def actualizar_grafico():
                ax2d.clear()
                ax3d.clear()
                
                # Obtener el registro actual usando self
                row = self.BD_Union.iloc[self.sondaje_actual_idx]
                nombre = row['Plataforma']
                
                collar = np.array([row['ESTE'], row['NORTE'], row['ELEV']])
                fin_sondaje = collar - np.array([0, 0, row['LENGTH']])
                
                cruces = calcular_intersecciones(collar, fin_sondaje, self.lineas_labores, radio_halo=30.0)
                
                # 2D View
                ax2d.plot([0, 0], [collar[2], fin_sondaje[2]], color='darkred', lw=3, label=f'Sondaje {nombre}')
                ax2d.scatter(0, collar[2], color='red', s=100, zorder=5)
                ax2d.axvspan(-30, 30, color='orange', alpha=0.08, label='Halo ±30m')
                
                cruces_por_nivel = {}
                for cruce in cruces:
                    p_lab = cruce['punto_labor']
                    nivel_clave = round(p_lab[2] / 5.0) * 5
                    if nivel_clave not in cruces_por_nivel or cruce['distancia'] < cruces_por_nivel[nivel_clave]['distancia']:
                        cruces_por_nivel[nivel_clave] = cruce
            
                for nivel, cruce in cruces_por_nivel.items():
                    p_lab = cruce['punto_labor']
                    metraje = cruce['metraje_sondaje']
                    dist_perpend = cruce['distancia']
                    elev_labor = p_lab[2]
                    
                    x_plot = dist_perpend if (p_lab[0] > collar[0]) else -dist_perpend
                    
                    ax2d.scatter(x_plot, elev_labor, color='blue', s=80, marker='s', zorder=4)
                    ax2d.plot([0, x_plot], [collar[2] - metraje, elev_labor], 'b:', alpha=0.6)
                    
                    ax2d.text(x_plot + (4 if x_plot >= 0 else -24), elev_labor, 
                              f"CRUCE A: {metraje:.1f} m\nDist: {dist_perpend:.1f}m", 
                              fontsize=8, bbox=dict(facecolor='yellow', alpha=0.85, boxstyle='round,pad=0.2'))
            
                ax2d.set_title(f"Sección Vertical Control: {nombre}", fontsize=10, fontweight='bold')
                ax2d.set_xlabel("Distancia a la Labor (m)")
                ax2d.set_ylabel("Elevación (msnm)")
                ax2d.set_xlim(-55, 55)
                ax2d.set_ylim(fin_sondaje[2] - 15, collar[2] + 15)
                ax2d.grid(True, linestyle='--', alpha=0.4)
                ax2d.legend(loc='upper right', fontsize=8)
                
                # 3D View
                ax3d.plot([collar[0], fin_sondaje[0]], [collar[1], fin_sondaje[1]], [collar[2], fin_sondaje[2]], 
                          color='darkred', lw=4, label=nombre)
                ax3d.scatter(collar[0], collar[1], collar[2], color='red', s=100)
                
                tramos_dibujados = set()
                for cruce in cruces:
                    p1, p2 = cruce['tramo']
                    id_tramo = (p1.tobytes(), p2.tobytes())
                    if id_tramo not in tramos_dibujados:
                        ax3d.plot([p1[0], p2[0]], [p1[1], p2[1]], [p1[2], p2[2]], color='blue', lw=2.5)
                        tramos_dibujados.add(id_tramo)
                        
                ax3d.set_title("Visualización Operacional 3D", fontsize=10, fontweight='bold')
                ax3d.set_xlabel("Este (X)")
                ax3d.set_ylabel("Norte (Y)")
                ax3d.set_zlabel("Elevación (Z)")
                
                ax3d.set_xlim(collar[0] - 60, collar[0] + 60)
                ax3d.set_ylim(collar[1] - 60, collar[1] + 60)
                ax3d.set_zlim(fin_sondaje[2] - 10, collar[2] + 10)
                
                dibujar_lista_lateral()
                fig.canvas.draw_idle()
            
            # =========================================================================
            # 4. MANEJO DE BUSQUEDA Y FILTRADO DINÁMICO
            # =========================================================================
            def obtener_sondajes_filtrados():
                nombres_completos = self.BD_Union['Plataforma'].tolist()
                filtrados = []
                for i, nom in enumerate(nombres_completos):
                    if self.texto_filtro.upper() in nom.upper():
                        filtrados.append((i, nom))
                return filtrados
            
            def dibujar_lista_lateral():
                ax_lista.clear()
                ax_lista.set_axis_off()
                
                sondajes_visibles = obtener_sondajes_filtrados()
                ax_lista.text(0.05, 0.96, f"Resultados ({len(sondajes_visibles)})", fontsize=9, fontweight='bold', transform=ax_lista.transAxes)
                
                for i in range(self.max_visibles_lista):
                    idx_lista_filtrada = self.indice_scroll_lista + i
                    if idx_lista_filtrada >= len(sondajes_visibles): break
                    
                    idx_original, nombre_sondaje = sondajes_visibles[idx_lista_filtrada]
                    y_pos = 0.88 - (i * 0.045)
                    
                    if idx_original == self.sondaje_actual_idx:
                        ax_lista.text(0.05, y_pos, f"▶ {nombre_sondaje}", color='blue', fontweight='bold', fontsize=9, transform=ax_lista.transAxes, bbox=dict(facecolor='lightblue', alpha=0.5, boxstyle='round,pad=0.15'))
                    else:
                        ax_lista.text(0.05, y_pos, f"  {nombre_sondaje}", color='black', fontsize=9, transform=ax_lista.transAxes)
            
            def al_escribir_buscador(texto):
                # Se eliminó 'global', ahora modifica los atributos de la instancia 'self'
                self.texto_filtro = texto
                self.indice_scroll_lista = 0  
                
                sondajes_visibles = obtener_sondajes_filtrados()
                if len(sondajes_visibles) > 0:
                    self.sondaje_actual_idx = sondajes_visibles[0][0]
                    actualizar_grafico()
                else:
                    dibujar_lista_lateral()
                    fig.canvas.draw_idle()
            
            def al_hacer_click_lista(event):
                if event.inaxes == ax_lista:
                    y_click = event.ydata
                    linea_presionada = int((0.90 - y_click) / 0.045)
                    
                    sondajes_visibles = obtener_sondajes_filtrados()
                    idx_lista_filtrada = self.indice_scroll_lista + linea_presionada
                    
                    if 0 <= idx_lista_filtrada < len(sondajes_visibles):
                        self.sondaje_actual_idx = sondajes_visibles[idx_lista_filtrada][0]
                        actualizar_grafico() # <-- Ahora llamará con el nuevo index guardado en self
            
            def al_usar_scroll(event):
                if event.inaxes == ax_lista:
                    sondajes_visibles = obtener_sondajes_filtrados()
                    if event.button == 'up' and self.indice_scroll_lista > 0:
                        self.indice_scroll_lista -= 1
                    elif event.button == 'down' and self.indice_scroll_lista < len(sondajes_visibles) - self.max_visibles_lista:
                        self.indice_scroll_lista += 1
                    dibujar_lista_lateral()
                    fig.canvas.draw_idle()
            
            # =========================================================================
            # 5. INICIALIZACIÓN DEL ENTORNO
            # =========================================================================
            fig = plt.subplots(figsize=(11, 8.5))[0]
            plt.subplots_adjust(left=0.22, bottom=0.08, right=0.95, top=0.90, wspace=0.25)
            
            ax_lista = plt.axes([0.01, 0.05, 0.18, 0.78])
            ax2d = plt.axes([0.24, 0.08, 0.33, 0.80])
            ax3d = plt.axes([0.62, 0.08, 0.35, 0.80], projection='3d')
            
            ax_box = plt.axes([0.02, 0.88, 0.16, 0.04])
            box_buscar = TextBox(ax_box, 'Buscar: ', initial="", label_pad=0.05)
            box_buscar.on_submit(al_escribir_buscador)
            box_buscar.on_text_change(al_escribir_buscador) 
            
            fig.canvas.mpl_connect('button_press_event', al_hacer_click_lista)
            fig.canvas.mpl_connect('scroll_event', al_usar_scroll)
            
            # Guardar referencias de los objetos interactivos de Matplotlib en self 
            # para que el Garbage Collector de Python no los destruya en el entorno PyQt.
            self.fig = fig
            self.box_buscar = box_buscar 
            
            actualizar_grafico()
            plt.show()
                
    def _limpiar_saltos_de_un_turno(self, sondajes):
        """Colapsa saltos de un solo turno que vuelven inmediatamente al sondaje anterior
        (secuencia A, A, B, A, A -> la "B" aislada se trata como A). Un cambio de sondaje
        real dura varios turnos; una sola hoja distinta rodeada del mismo sondaje antes y
        después es casi siempre una celda mal tipeada en esa hoja (a menudo con 0 m
        perforados), no un traslado de la máquina ida y vuelta el mismo día."""
        lista = list(sondajes)
        for i in range(1, len(lista) - 1):
            if lista[i] != lista[i - 1] and lista[i - 1] == lista[i + 1]:
                lista[i] = lista[i - 1]
        return lista

    def _build_movimientos_maquina(self):
        """Detecta, para cada máquina, los cambios de sondaje entre turnos consecutivos
        (terminó uno y pasó a otro) y calcula la distancia en línea recta entre las
        coordenadas de collar de ambos (self.BD_Union: ESTE/NORTE/ELEV, la misma fuente que
        usa la Vista en Planta — más completa que Este_Rep/Norte_Rep/Cota_Rep de "Control de
        Proceso"). Es un APROXIMADO: la distancia real de traslado depende del camino/accesos
        entre plataformas, no de la línea recta. Requiere haber cargado Control de Proceso
        (LoadProc); si no está disponible devuelve un DataFrame vacío."""
        if not hasattr(self, 'BD_Union'):
            return pd.DataFrame(columns=['Fecha', 'Maquina', 'Sondaje_Origen', 'Sondaje_Destino', 'Distancia_Aprox_m'])

        coords = (
            self.BD_Union[['Sondaje', 'ESTE', 'NORTE', 'ELEV']]
            .dropna(subset=['Sondaje'])
            .drop_duplicates(subset='Sondaje')
            .set_index('Sondaje')
        )

        df = self.df.copy()
        df['Fecha'] = pd.to_datetime(df['Fecha']).dt.normalize()
        # A nivel de turno (no de día) para no perderse un cambio de sondaje ocurrido entre
        # el turno DIA y NOCHE del mismo día. "DIA" ordena antes que "NOCHE" alfabéticamente,
        # que además coincide con el orden cronológico real de los turnos.
        secuencia = (df[['Maquina', 'Fecha', 'Turno', 'Sondaje']]
                     .dropna(subset=['Sondaje'])
                     .sort_values(['Maquina', 'Fecha', 'Turno']))

        movimientos = []
        for maquina, grupo in secuencia.groupby('Maquina'):
            grupo = grupo.reset_index(drop=True)
            # Se limpia la secuencia ANTES de comparar turno a turno: una sola hoja con el
            # sondaje mal tipeado (ej. HYDX-5A-05 con una hoja suelta de "MR26061-11" en
            # medio de varias semanas seguidas en "MR26065-11", con 0 m perforados esa hoja)
            # generaba un "movimiento" de ida y vuelta el mismo día que nunca ocurrió, y de
            # paso ensuciaba las coordenadas de otras máquinas que sí trabajaron ese sondaje
            # de verdad. Ver _limpiar_saltos_de_un_turno.
            sondajes_limpios = self._limpiar_saltos_de_un_turno(grupo['Sondaje'].tolist())
            sondaje_anterior = None
            for idx, sondaje_actual in enumerate(sondajes_limpios):
                fila = grupo.loc[idx]
                if sondaje_anterior is not None and sondaje_actual != sondaje_anterior:
                    c1 = coords.loc[sondaje_anterior] if sondaje_anterior in coords.index else None
                    c2 = coords.loc[sondaje_actual] if sondaje_actual in coords.index else None
                    distancia = np.nan
                    if c1 is not None and c2 is not None:
                        distancia = float(np.sqrt(
                            (c2['ESTE'] - c1['ESTE']) ** 2 +
                            (c2['NORTE'] - c1['NORTE']) ** 2 +
                            (c2['ELEV'] - c1['ELEV']) ** 2
                        ))
                    movimientos.append({
                        'Fecha': fila['Fecha'], 'Maquina': maquina,
                        'Sondaje_Origen': sondaje_anterior, 'Sondaje_Destino': sondaje_actual,
                        'Distancia_Aprox_m': distancia,
                    })
                sondaje_anterior = sondaje_actual

        columnas = ['Fecha', 'Maquina', 'Sondaje_Origen', 'Sondaje_Destino', 'Distancia_Aprox_m']
        if not movimientos:
            return pd.DataFrame(columns=columnas)
        return pd.DataFrame(movimientos)[columnas].sort_values(['Fecha', 'Maquina']).reset_index(drop=True)

    def _resumen_taladros(self):
        """Taladros perforados vs planeados del proyecto, contando por PLATAFORMA (no por
        Sondaje/DDH): cuando una perforación se detiene por una falla operativa se perfora
        un "gemelo" desde la MISMA plataforma con otro identificador (ej. "33" y "33B" en
        Control de Proceso, mismas coordenadas de collar), y ambos deben contar como un solo
        taladro planeado, no dos — de ahí las plataformas duplicadas que se ven en ese Excel.
        "Perforados" = plataformas cuyo sondaje (el vigente, no el gemelo abandonado) quedó
        en estado "Ejecutado". Requiere Control de Proceso cargado (self.BD_Union_total); si
        no está disponible devuelve None."""
        if not hasattr(self, 'BD_Union_total'):
            return None
        plataformas = self.BD_Union_total.dropna(subset=['Plataforma'])
        planeados = plataformas['Plataforma'].nunique()
        perforados = plataformas.loc[plataformas['Estado'] == 'Ejecutado', 'Plataforma'].nunique()
        return perforados, planeados

    def _build_df_plot_plano(self):
        """Prepara df_plot para la Vista en Planta: proyección de días restantes, jittering de
        coordenadas duplicadas y campos formateados para el hover / las etiquetas. Reutilizado
        por PlotPlano (interactivo) y _render_plano_png (versión estática para el reporte Word)."""
        self.df['Fecha'] = pd.to_datetime(self.df['Fecha'])
        avance_diario = self.df.groupby(["Sondaje", "Fecha"])["Avance Turno"].sum().reset_index()

        promedio_2dias = avance_diario.sort_values(["Sondaje", "Fecha"]).groupby("Sondaje").tail(2).groupby("Sondaje")["Avance Turno"].mean().rename("Avance_2dias_Prom").reset_index()
        dias_perforando = avance_diario.groupby("Sondaje")["Fecha"].agg(lambda x: (x.max() - x.min()).days + 1).rename("Dias_Perforando_Count").reset_index()

        self.BD_Union_total = self.BD_Union_total.drop(columns=["Avance_2dias_Prom", "Dias_Perforando_Count", "Proyeccion_Dias"], errors='ignore')
        self.BD_Union_total = self.BD_Union_total.merge(promedio_2dias, on="Sondaje", how="left").merge(dias_perforando, on="Sondaje", how="left")

        self.BD_Union_total["Proyeccion_Dias"] = self.BD_Union_total.apply(
            lambda r: round(r["Metro_Faltante"] / r["Avance_2dias_Prom"], 1)
            if (r["Estado"] == "Perforando" and r.get("Avance_2dias_Prom", 0) > 0) else np.nan, axis=1
        )

        df_plot = self.BD_Union_total.copy().reset_index(drop=True)

        # --- OPTIMIZACIÓN PARA GEMELOS (JITTERING) ---
        # Identificamos duplicados en coordenadas para desplazarlos visualmente
        df_plot['coord_group'] = df_plot.groupby(['ESTE', 'NORTE']).cumcount()
        # Desplazamos 6 metros en X por cada duplicado para que no se solapen los puntos de 30s
        df_plot['ESTE_PLOT'] = df_plot['ESTE'] + (df_plot['coord_group'] * 6)
        df_plot['NORTE_PLOT'] = df_plot['NORTE']

        # Formateo para Hover
        df_plot["Proyeccion_fmt"] = df_plot["Proyeccion_Dias"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "-")
        df_plot["Longitud_fmt"] = df_plot["LENGTH"].apply(lambda x: f"{x:.1f}" if pd.notna(x) and x > 0 else "-")
        df_plot["Avance_fmt"] = df_plot["Avance_Actual"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "0.0")
        df_plot["Faltante_fmt"] = df_plot["Metro_Faltante"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "0.0")
        df_plot["Sondaje_fmt"] = df_plot["Sondaje"].fillna("-").astype(str)
        df_plot["Maquina_fmt"] = df_plot["Maquina"].fillna("S/M").astype(str)

        def get_despl(row):
            if pd.isna(row["Este_Rep"]) or pd.isna(row["ESTE"]): return "-"
            dist = np.sqrt((row["ESTE"]-row["Este_Rep"])**2 + (row["NORTE"]-row["Norte_Rep"])**2)
            de, dn = row["Este_Rep"] - row["ESTE"], row["Norte_Rep"] - row["NORTE"]
            ang = (np.degrees(np.arctan2(de, dn)) + 360) % 360
            dirs = [(22.5, 67.5, "NE"), (67.5, 112.5, "E"), (112.5, 157.5, "SE"),
                    (157.5, 202.5, "S"), (202.5, 247.5, "SW"), (247.5, 292.5, "W"), (292.5, 337.5, "NW")]
            d_label = next((lab for low, high, lab in dirs if low <= ang < high), "N")
            return f"{dist:.2f} m ({d_label})"

        df_plot["Despl_fmt"] = df_plot.apply(get_despl, axis=1)
        return df_plot

    def _color_map_por_categoria(self, categorias, saturacion=0.45, brillo=0.88):
        """Genera un color distinto por categoría (p.ej. por sondaje) avanzando el matiz en
        pasos de la proporción áurea en vez de pasos lineales (i/n). Con pasos lineales,
        categorías con IDs correlativos (p.ej. los sondajes de una misma máquina, que suelen
        numerarse en secuencia) terminan adyacentes en la lista y reciben matices casi idénticos
        (todo en la franja roja/magenta, como se vio en el dashboard). Los pasos áureos evitan
        ese amontonamiento incluso entre categorías consecutivas. Se usa saturación/brillo
        moderados (en vez de HSV puro) para obtener tonos pastel más apropiados para un reporte,
        en vez de colores primarios muy saturados."""
        GOLDEN_RATIO_CONJUGATE = 0.6180339887498949
        color_map = {}
        hue = 0.0
        for categoria in categorias:
            color_map[categoria] = colorsys.hsv_to_rgb(hue, saturacion, brillo)
            hue = (hue + GOLDEN_RATIO_CONJUGATE) % 1.0
        return color_map

    def _dibujar_topografia(self, ax):
        """Dibuja las líneas de topografía (DXF) sobre `ax`, si hay una carpeta configurada."""
        if hasattr(self, 'ruta_topo') and self.ruta_topo and os.path.exists(self.ruta_topo):
            try:
                for archivo in os.listdir(self.ruta_topo):
                    if archivo.lower().endswith(".dxf"):
                        try:
                            ruta_completa = os.path.join(self.ruta_topo, archivo)
                            doc = ezdxf.readfile(ruta_completa)
                            for e in doc.modelspace():
                                if e.dxftype() in ("POLYLINE", "LWPOLYLINE"):
                                    if e.dxftype() == "LWPOLYLINE":
                                        p = e.get_points()
                                    else:
                                        p = [v.dxf.location for v in e.vertices]
                                    if p:
                                        ax.plot([pt[0] for pt in p], [pt[1] for pt in p],
                                                color="#E0E0E0", lw=0.5, zorder=1)
                        except Exception as e_dxf:
                            print(f"Error leyendo archivo DXF {archivo}: {e_dxf}")
                            continue
            except Exception as e_topo:
                print(f"No se pudo procesar la carpeta de topografía: {e_topo}")
        else:
            print("Topografía omitida: Ruta no definida o no encontrada.")

    def _etiquetar_puntos_plano(self, ax, sub, fontsize, offset_paso=25):
        """Etiqueta TODOS los puntos de `sub` (ya filtrado a un solo estado) con su
        plataforma y, si hay una máquina asociada, el nombre de la máquina — separando
        verticalmente los que comparten coordenadas (gemelos perforados por una falla
        operativa, ver _build_resumen_sondajes) para que no se superpongan las etiquetas.
        Antes esto solo se hacía para "Perforando"; "Preinicio" quedaba sin ninguna
        etiqueta porque el agrupamiento de coordenadas duplicadas se calculaba una sola vez
        de forma GLOBAL en _build_df_plot_plano (coord_group), así que un gemelo en
        Preinicio con las mismas coordenadas que su par en otro estado (ej. uno ya
        "Perforando") nunca caía en el primer lugar del grupo y no se etiquetaba. Aquí el
        agrupamiento se recalcula DENTRO del propio subconjunto de `sub`, así que cada
        estado etiqueta sus propios puntos sin depender de qué otro estado haya tomado el
        primer lugar en esa coordenada. Devuelve la lista de Text creados (para poder
        mostrarlos/ocultarlos junto con la capa, ver PlotPlano)."""
        sub = sub.copy()
        sub['n_en_plataforma'] = sub.groupby(['ESTE', 'NORTE'])['Sondaje_fmt'].transform('count')
        sub['idx_en_plataforma'] = sub.groupby(['ESTE', 'NORTE']).cumcount()
        textos = []
        for _, r in sub.iterrows():
            texto = r['Plataforma'] if r['n_en_plataforma'] == 1 else f"{r['Plataforma']} ({r['Sondaje_fmt']})"
            if r['Maquina_fmt'] not in ('S/M', '-', ''):
                texto += f"\n{r['Maquina_fmt']}"
            offset_y = r['idx_en_plataforma'] * offset_paso
            txt = ax.text(r["ESTE"] - 12, r["NORTE"] + offset_y, texto, fontsize=fontsize,
                          ha='right', va='center', color='#333333', zorder=5, fontweight='bold')
            txt.set_clip_on(True)
            textos.append(txt)
        return textos

    def PlotPlano(self):
            self.Mensajes.clear()
            try:
                df_plot = self._build_df_plot_plano()

                # =====================================================
                # 2. FIGURA Y CONFIGURACIÓN
                # =====================================================
                fig, ax = plt.subplots(figsize=(16, 9))
                # Ajuste de márgenes: más espacio abajo para la leyenda y derecha para progreso
                plt.subplots_adjust(right=0.72, top=0.88, left=0.06, bottom=0.15) 
                ax.set_facecolor('white')
                color_map = {"Programado": "gray", "Ejecutado": "green", "Perforando": "red", "Plataforma": "orange", "Preinicio": "magenta"}
                
                ax.grid(True, which='major', color='#EEEEEE', linestyle='-', linewidth=0.8, zorder=0)
                ax.minorticks_on()
                # --- 3.1 TOPOGRAFÍA ---
                self._dibujar_topografia(ax)
                # Capas de Sondajes
                capas_dict = {}
                for estado, color in color_map.items():
                    sub = df_plot[df_plot["Estado"] == estado].copy()
                    if not sub.empty:
                        # Se grafican los puntos usando las coordenadas con desplazamiento (PLOT)
                        sc = ax.scatter(sub["ESTE_PLOT"], sub["NORTE_PLOT"], c=color, s=30, lw=0.5, zorder=4, label=estado)
                        
                        # Las etiquetas de texto se mantienen en la coordenada REAL para no duplicar nombres
                        lbls = []
                        if estado in ("Perforando", "Preinicio"):
                            # Muy pocos sondajes están "Perforando"/"Preinicio" a la vez, así
                            # que aquí SÍ etiquetamos todos con plataforma + máquina (ver
                            # _etiquetar_puntos_plano), en vez de solo el resto de estados
                            # (Programado/Ejecutado/Plataforma, con decenas de puntos) que se
                            # dejan sin etiqueta para que el mapa no se sature.
                            lbls = self._etiquetar_puntos_plano(ax, sub, fontsize=6)
                        else:
                            for _, r in sub.iterrows():
                                # Solo dibujar etiqueta si es el primer sondaje del grupo en esa coordenada para evitar amontonamiento
                                if r["coord_group"] == 0:
                                    txt = ax.text(r["ESTE"]-12, r["NORTE"], r["Plataforma"], fontsize=6,
                                                  ha='right', va='center', color='#333333', zorder=5, fontweight='bold')
                                    txt.set_clip_on(True)
                                    lbls.append(txt)

                        capas_dict[sc] = {'estado': estado, 'labels': lbls, 'df': sub}
                
                linea_real, = ax.plot([], [], 'gray', linestyle='--', lw=1, zorder=3, visible=False)
                punto_real, = ax.plot([], [], 'ro', markersize=5, markeredgecolor='white', zorder=6, visible=False)
                
                annot = ax.annotate("", xy=(0,0), xytext=(20,20), textcoords="offset points",
                                    bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="green", lw=1.5),
                                    arrowprops=dict(arrowstyle="->", color='gray'), zorder=10, fontsize=10)
                annot.set_visible(False)
                # =====================================================
                # 3. INTERACTIVIDAD Y LEYENDA HORIZONTAL
                # =====================================================
                def update_annot(sc, ind, estado):
                    idx = ind["ind"][0]
                    row = capas_dict[sc]['df'].iloc[idx]
                    pos = sc.get_offsets()[idx]
                    text = (f"Plataforma: {row['Plataforma']}\nSondaje: {row['Sondaje_fmt']}\n"
                            f"Desplazamiento: {row['Despl_fmt']}\nMetraje Programado: {row['Longitud_fmt']} m\n"
                            f"Avance: {row['Avance_fmt']} m\n"
                            f"Faltante: {row['Faltante_fmt']} m\nEstado: {estado}\nMaquina: {row['Maquina_fmt']}")
                    annot.xy = pos
                    annot.set_text(text)
                    annot.get_bbox_patch().set_edgecolor(color_map[estado])
                    if pd.notna(row["Este_Rep"]):
                        linea_real.set_data([pos[0], row["Este_Rep"]], [pos[1], row["Norte_Rep"]])
                        punto_real.set_data([row["Este_Rep"]], [row["Norte_Rep"]])
                        linea_real.set_visible(True); punto_real.set_visible(True)
                
                def hover(event):
                    if event.inaxes == ax:
                        for sc, data in capas_dict.items():
                            if not sc.get_visible(): continue
                            cont, ind = sc.contains(event)
                            if cont:
                                update_annot(sc, ind, data['estado'])
                                annot.set_visible(True); fig.canvas.draw_idle(); return
                    if annot.get_visible():
                        annot.set_visible(False); linea_real.set_visible(False); punto_real.set_visible(False); fig.canvas.draw_idle()
                
                # --- LEYENDA ABAJO Y HORIZONTAL ---
                leg = ax.legend(loc='upper center', bbox_to_anchor=(0.38, 0.08), bbox_transform=fig.transFigure,
                                ncol=4, title="Estado (Clic para ocultar)", frameon=True, shadow=True, fontsize=10)
                map_leg = {leg_line: sc for leg_line, sc in zip(leg.legend_handles, capas_dict.keys())}
                
                def on_pick(event):
                    if event.artist in map_leg:
                        sc = map_leg[event.artist]
                        vis = not sc.get_visible()
                        sc.set_visible(vis)
                        for t in capas_dict[sc]['labels']: t.set_visible(vis)
                        event.artist.set_alpha(1.0 if vis else 0.2); fig.canvas.draw_idle()
                
                fig.canvas.mpl_connect("motion_notify_event", hover)
                fig.canvas.mpl_connect("pick_event", on_pick)
                for line in leg.legend_handles: line.set_picker(True)
                
                # =====================================================
                # 4. PANEL DE PROGRESO CORREGIDO
                # =====================================================
                df_bar = df_plot[df_plot["Estado"] == "Perforando"].head(5)
                fig.text(0.86, 0.91, 'PROGRESO DE PERFORACIÓN', fontsize=11, fontweight='bold', ha='center')
                
                for i in range(5):
                    ax_card = fig.add_axes([0.74, 0.76 - (i * 0.16), 0.24, 0.13])
                    ax_card.axis('off')
                    
                    card_bg = patches.FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02", 
                                                     facecolor='#F9F9F9', edgecolor='#CCCCCC', linewidth=0.8, transform=ax_card.transAxes)
                    ax_card.add_patch(card_bg)
                
                    if i < len(df_bar):
                        r = df_bar.iloc[i]
                        total = r["LENGTH"] if r["LENGTH"] > 0 else (r["Avance_Actual"] + r["Metro_Faltante"])
                        avance, porc = r["Avance_Actual"], (r["Avance_Actual"] / total * 100) if total > 0 else 0
                        col = "#2E7D32" if porc >= 95 else "#E53935"
                        
                        ax_inner = fig.add_axes([0.76, 0.795 - (i * 0.16), 0.20, 0.035])
                        ax_inner.barh(0, total, color="#EEEEEE", height=0.8)
                        ax_inner.barh(0, avance, color=col, height=0.8)
                        ax_inner.set_xlim(0, total * 1.05)
                        ax_inner.axis('off')
                
                        ax_card.text(0.08, 0.82, f"{r['Plataforma']} / {r['Sondaje_fmt']}", fontsize=8.5, fontweight='bold', va='top')
                        ax_card.text(0.92, 0.82, f"{porc:.1f}%", fontsize=8.5, fontweight='bold', ha='right', va='top', color=col)
                        
                        # --- SOLUCIÓN AL ERROR DE NAN ---
                        # Validamos si el valor es nulo antes de convertir a entero
                        dias_val = int(r['Dias_Perforando_Count']) if pd.notna(r['Dias_Perforando_Count']) else 0
                        
                        footer = f"Máquina: {r['Maquina_fmt']} | Días: {dias_val} | Dias Restantes (Proy): {r['Proyeccion_fmt']} d"
                        ax_card.text(0.08, 0.22, footer, fontsize=7.5, color='#555555', va='top')
                    else:
                        ax_card.text(0.5, 0.5, "Espacio Disponible", ha='center', va='center', alpha=0.3, fontsize=8)
                # Título y Ejes
                fig.suptitle("ESTADO DE PERFORACIÓN - MREDP 2026", fontsize=16, fontweight='bold', y=0.96)
                ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%0.0f'))
                ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%0.0f'))
                ax.set_xlabel("Este (m)", fontweight='bold'); ax.set_ylabel("Norte (m)", fontweight='bold')
                ax.set_xlim(374400, 376800); ax.set_ylim(8715600, 8718580); ax.set_aspect('equal')
                
                plt.show()
                self.Mensajes.setText("Ploteo de plano Completa")
                self.PlanPerf.setEnabled(False)
                self.CtrlProceso.setEnabled(False)
    
            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error al Graficar", f"Error crítico: {str(e)}")
                self.Mensajes.setText("Ocurrio algun error")

    def PlotDash(self):
            self.Mensajes.clear()
            try:
            
                # ======================================================
                # 1. PREPARACIÓN DE DATA
                # ======================================================
                # Aseguramos el formato de fecha y normalizamos a medianoche
                self.df['Fecha'] = pd.to_datetime(self.df['Fecha'])
                self.df['Dia_dt'] = self.df['Fecha'].dt.normalize()
                
                # Variables globales para el dashboard
                maquinas = self._maquinas_validas()
                num_maquinas = len(maquinas)
                sondajes_unicos = self.df['Sondaje'].unique()
                
                # Límites numéricos para el slider (Matplotlib usa floats para fechas)
                date_nums = mdates.date2num(self.df['Dia_dt'].unique())
                d_min, d_max = date_nums.min(), date_nums.max()
                
                # Mapa de colores persistente para los sondajes (ver _color_map_por_categoria).
                color_map_sondaje = self._color_map_por_categoria(sondajes_unicos)
                
                # Función para el formato de fecha en la etiqueta del slider
                def formato_fecha_slider(val):
                    d1 = mdates.num2date(val[0]).strftime('%d/%m/%y')
                    d2 = mdates.num2date(val[1]).strftime('%d/%m/%y')
                    return f'{d1} a {d2}'
                
                # ======================================================
                # 2. FIGURA Y LOGICA DEL DASHBOARD
                # ======================================================
                fig1, axes = plt.subplots(num_maquinas, 1, figsize=(15, 2.5 * num_maquinas + 1.5), sharex=True)
                if num_maquinas == 1: axes = [axes]
                
                # Ajuste de márgenes para leyendas y sliders
                plt.subplots_adjust(bottom=0.18, hspace=0.5, right=0.82, top=0.92)
                
                def actualizar_avance(val):
                    f_min = mdates.num2date(val[0]).replace(tzinfo=None)
                    f_max = mdates.num2date(val[1]).replace(tzinfo=None)
                    
                    # Rango de días que se mostrarán en el eje X
                    rango_visible = pd.date_range(start=f_min, end=f_max).normalize()
                    
                    for i, maquina in enumerate(maquinas):
                        ax = axes[i]
                        ax.clear()
                        
                        # Filtro de datos por máquina y fecha
                        df_m = self.df[(self.df['Maquina'] == maquina) & (self.df['Dia_dt'] >= f_min) & (self.df['Dia_dt'] <= f_max)]
                        
                        if not df_m.empty:
                            # Pivotar datos: Filas=Días, Columnas=Sondajes
                            data_pivot = df_m.groupby(['Dia_dt', 'Sondaje'])['Avance Turno'].sum().unstack(fill_value=0)
                            data_pivot = data_pivot.reindex(rango_visible, fill_value=0)
                            
                            bottom = np.zeros(len(data_pivot))
                            sondajes_en_rango = data_pivot.columns[data_pivot.sum() > 0]
                            
                            # Dibujar barras apiladas
                            for sondaje in data_pivot.columns:
                                values = data_pivot[sondaje].values
                                ax.bar(data_pivot.index, values, bottom=bottom, 
                                       color=color_map_sondaje.get(sondaje), 
                                       label=sondaje if sondaje in sondajes_en_rango else None, 
                                       edgecolor='black', lw=0.3, width=0.8)
                                bottom += values
                            
                            # Etiquetas de totales (Metros acumulados por día)
                            totales = data_pivot.sum(axis=1)
                            for x, total in zip(data_pivot.index, totales):
                                if total > 0:
                                    ax.text(x, total + 1, f'{total:.1f}', ha='center', va='bottom', 
                                            fontsize=7, fontweight='bold', rotation=90, clip_on=True)
                
                            # --- AJUSTE DE ESCALA DINÁMICA ---
                            # Buscamos el máximo valor en este rango para que la barra no se corte
                            max_avance = totales.max()
                            # Margen del 15% para que quepa el texto del total arriba de la barra
                            ax.set_ylim(0, max(60, max_avance * 1.15)) 
                            
                            # Leyenda filtrada por máquina
                            handles, labels = ax.get_legend_handles_labels()
                            if handles:
                                ax.legend(handles, labels, title="Sondajes", bbox_to_anchor=(1.01, 1), 
                                          loc='upper left', fontsize=7, frameon=False)
                
                        # Configuración estética por subplot
                        ax.set_title(f'Máquina: {maquina}', fontweight='bold', loc='left', fontsize=10)
                        ax.axhline(35, color='red', linestyle='--', alpha=0.4, label="Meta (35m)")
                        ax.grid(True, axis='y', linestyle=':', alpha=0.6)
                
                    # Formateo del eje X (Fechas)
                    axes[-1].xaxis.set_major_formatter(mdates.DateFormatter('%d-%m'))
                    # Evitar saturación de etiquetas en el eje X
                    paso = max(1, len(rango_visible) // 35)
                    axes[-1].set_xticks(rango_visible[::paso])
                    plt.setp(axes[-1].get_xticklabels(), rotation=90, fontsize=8)
                    
                    fig1.canvas.draw_idle()
                
                # ======================================================
                # 3. CONFIGURACIÓN DEL SLIDER
                # ======================================================
                ax_slider_av = plt.axes([0.3, 0.05, 0.4, 0.025])
                # Iniciamos mostrando los últimos 30 días
                self.slider_av = RangeSlider(ax_slider_av, 'Rango Fecha ', d_min, d_max, valinit=(d_max-30, d_max))
                
                # Inyección de la función de formato para mostrar fechas reales
                self.slider_av._format = formato_fecha_slider
                self.slider_av.valtext.set_text(formato_fecha_slider(self.slider_av.val))
                self.slider_av.on_changed(actualizar_avance)
                
                # Ejecución inicial
                actualizar_avance((d_max-30, d_max))
                plt.show()
                self.Mensajes.setText("Resumen grafico Ploteada")


                # 1. PREPARACIÓN Y CONFIGURACIÓN
                self.df['Fecha'] = pd.to_datetime(self.df['Fecha'])
                self.df['Dia_dt'] = self.df['Fecha'].dt.normalize()
                
                # Fechas de inicio detectadas automáticamente (self.machine_lifecycle)
                start_dates_config = {m: cfg['start'] for m, cfg in self.machine_lifecycle.items()}

                maquinas = self._maquinas_validas()
                date_nums = mdates.date2num(self.df['Dia_dt'].unique())
                d_min, d_max = date_nums.min(), date_nums.max()
                
                colores_m = plt.colormaps['Set1'](np.linspace(0, 1, len(maquinas)))
                color_map_maquina = {m: colores_m[i] for i, m in enumerate(maquinas)}
        
                # 2. CÁLCULO DE % DE CUMPLIMIENTO (KPI)
                stats_labels = []
                for maq in maquinas:
                    f_inicio = pd.to_datetime(start_dates_config.get(maq, '2000-01-01'))
                    df_maq = self.df[(self.df['Maquina'] == maq) & (self.df['Dia_dt'] >= f_inicio) & (self.df['Avance Turno'] > 0)]
                    if not df_maq.empty:
                        diario = df_maq.groupby('Dia_dt')['Avance Turno'].sum()
                        cumple = (diario > 35).sum()
                        porcentaje = (cumple / len(diario)) * 100
                        stats_labels.append(f"{maq} ({porcentaje:.1f}%)")
                    else:
                        stats_labels.append(f"{maq} (0%)")
        
                # 3. FIGURA
                fig, ax = plt.subplots(figsize=(15, 8))
                plt.subplots_adjust(bottom=0.2, right=0.72, left=0.07)
                
                self.lineas_dict = {}
                self.scatters_dict = {}
                
                annot = ax.annotate("", xy=(0,0), xytext=(10,10), textcoords="offset points",
                                    bbox=dict(boxstyle="round", fc="w", alpha=0.9, ec="gray"),
                                    arrowprops=dict(arrowstyle="->"))
                annot.set_visible(False)
                
                def actualizar_tendencia(val):
                    f_min_dt = mdates.num2date(val[0]).replace(tzinfo=None)
                    f_max_dt = mdates.num2date(val[1]).replace(tzinfo=None)
                    
                    ax.clear()
                    ax.axhline(35, color='red', linestyle='--', alpha=0.4)
                    
                    # Ajuste CRÍTICO: Forzar el límite del eje X al rango del slider
                    ax.set_xlim(f_min_dt, f_max_dt)
                    
                    for i, maq in enumerate(maquinas):
                        f_init_maq = pd.to_datetime(start_dates_config.get(maq, '2000-01-01'))
                        
                        # FILTRADO: Rango de fecha Y Avance > 0 para evitar ruidos
                        mask = (self.df['Maquina'] == maq) & \
                               (self.df['Dia_dt'] >= f_min_dt) & \
                               (self.df['Dia_dt'] <= f_max_dt) & \
                               (self.df['Dia_dt'] >= f_init_maq) & \
                               (self.df['Avance Turno'] > 0)
                        
                        df_plot = self.df[mask].groupby('Dia_dt')['Avance Turno'].sum().reset_index()
                        
                        if not df_plot.empty:
                            x_vals = mdates.date2num(df_plot['Dia_dt'])
                            y_vals = df_plot['Avance Turno'].values
                            color = color_map_maquina[maq]
                            
                            if len(x_vals) > 3:
                                x_new = np.linspace(x_vals.min(), x_vals.max(), 300)
                                spline = make_interp_spline(x_vals, y_vals, k=3)
                                y_smooth = spline(x_new).clip(min=0)
                                ln, = ax.plot(mdates.num2date(x_new), y_smooth, color=color, lw=2.5, alpha=0.8)
                            else:
                                ln, = ax.plot(df_plot['Dia_dt'], y_vals, color=color, lw=2.5, alpha=0.8)
                            
                            sc = ax.scatter(df_plot['Dia_dt'], y_vals, color=color, s=40, zorder=3, edgecolors='white')
                            
                            if hasattr(self, 'check_maquinas'):
                                vis = self.check_maquinas.get_status()[i]
                                ln.set_visible(vis); sc.set_visible(vis)
                            
                            self.lineas_dict[maq] = ln
                            self.scatters_dict[maq] = sc
                    
                    # Configuración de Ejes
                    ax.set_title('Tendencia de Avance Diario y % de Cumplimiento (>35m)', fontweight='bold', fontsize=14, loc='left')
                    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
                    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m'))
                    plt.setp(ax.get_xticklabels(), rotation=0, ha='center')
                    ax.grid(True, linestyle=':', alpha=0.5)
                    ax.add_artist(annot)
                    fig.canvas.draw_idle()
        
                # 4. CONTROLES
                rax = plt.axes([0.73, 0.3, 0.25, 0.5], frameon=False)
                self.check_maquinas = CheckButtons(rax, stats_labels, [True]*len(maquinas))
                for i, m in enumerate(maquinas):
                    self.check_maquinas.labels[i].set_color(color_map_maquina[m])
                    self.check_maquinas.labels[i].set_fontweight('bold')
        
                def func_visibilidad(label):
                    nombre_maq = label.split(' (')[0]
                    if nombre_maq in self.lineas_dict:
                        idx = [s.split(' (')[0] for s in stats_labels].index(nombre_maq)
                        vis = self.check_maquinas.get_status()[idx]
                        self.lineas_dict[nombre_maq].set_visible(vis)
                        self.scatters_dict[nombre_maq].set_visible(vis)
                        fig.canvas.draw_idle()
        
                self.check_maquinas.on_clicked(func_visibilidad)
        
                ax_slider = plt.axes([0.25, 0.05, 0.4, 0.03])
                self.slider = RangeSlider(ax_slider, 'Rango Fecha ', d_min, d_max, valinit=(d_max-30, d_max))
                self.slider.on_changed(actualizar_tendencia)
        
                actualizar_tendencia(self.slider.val)
                plt.show()

                # ======================================================
                # 5. SCATTERPLOT: HORAS PERFORACIÓN NORMAL vs METROS PERFORADOS
                # ======================================================
                fig3, ax3, scatters_maquina = self._draw_scatter_perfnormal_figure()

                cursor_scatter = mplcursors.cursor(list(scatters_maquina.keys()), hover=True)
                @cursor_scatter.connect("add")
                def on_add_scatter(sel):
                    maquina = scatters_maquina.get(sel.artist, '')
                    sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="black")
                    sel.annotation.set_text(
                        f"{maquina}\nHoras Perf. Normal: {sel.target[0]:.1f}\nMetros: {sel.target[1]:.1f}"
                    )

                plt.show()

            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error al Graficar", f"Error crítico: {str(e)}")
                self.Mensajes.setText("Ocurrio algun error")

    def _draw_scatter_perfnormal_figure(self, figsize=(11, 7), dias_recientes=30):
        """Un punto por día+máquina: cuánto se avanzó (Avance Turno) frente a cuántas horas
        de perforación efectiva (Horas Perforacion Normal, U14) tomó lograrlo, para ver de
        un vistazo si una máquina rinde más o menos metros por hora que las demás. Los
        últimos `dias_recientes` días se resaltan (punto más grande, opaco y con borde)
        sobre el resto (tenue) para ver de un vistazo si el rendimiento reciente mejoró o
        empeoró frente al histórico. Devuelve (fig, ax, {scatter_artist: maquina}) —
        reutilizado por PlotDash (interactivo, con hover) y _render_scatter_perfnormal_png
        (estático para el reporte Word)."""
        self.df['Dia_dt'] = pd.to_datetime(self.df['Fecha']).dt.normalize()
        resumen_dia_maq = self.df.groupby(['Dia_dt', 'Maquina']).agg(
            Metros=('Avance Turno', 'sum'),
            Horas_Perf_Normal=('Horas Perforacion Normal', 'sum'),
        ).reset_index()
        resumen_dia_maq = resumen_dia_maq[
            (resumen_dia_maq['Horas_Perf_Normal'] > 0) | (resumen_dia_maq['Metros'] > 0)
        ]

        fecha_max = resumen_dia_maq['Dia_dt'].max()
        fecha_corte = fecha_max - pd.Timedelta(days=dias_recientes - 1)
        es_reciente = resumen_dia_maq['Dia_dt'] >= fecha_corte

        fig, ax = plt.subplots(figsize=figsize)
        maquinas_scatter = self._maquinas_validas(resumen_dia_maq)
        colores_scatter = plt.colormaps['Set1'](np.linspace(0, 1, len(maquinas_scatter)))
        color_map_scatter = {m: colores_scatter[i] for i, m in enumerate(maquinas_scatter)}

        scatters_maquina = {}
        for maquina in maquinas_scatter:
            datos_m = resumen_dia_maq[resumen_dia_maq['Maquina'] == maquina]
            datos_viejos = datos_m[~es_reciente.loc[datos_m.index]]
            datos_recientes = datos_m[es_reciente.loc[datos_m.index]]

            if not datos_viejos.empty:
                sc_v = ax.scatter(datos_viejos['Horas_Perf_Normal'], datos_viejos['Metros'],
                                   label=maquina, color=color_map_scatter[maquina],
                                   alpha=0.25, s=22, edgecolors='none', zorder=2)
                scatters_maquina[sc_v] = maquina
            if not datos_recientes.empty:
                sc_r = ax.scatter(datos_recientes['Horas_Perf_Normal'], datos_recientes['Metros'],
                                   label=maquina if datos_viejos.empty else None,
                                   color=color_map_scatter[maquina],
                                   alpha=0.9, s=55, edgecolors='black', lw=0.6, zorder=3)
                scatters_maquina[sc_r] = maquina

        ax.set_title('Horas de Perforación Normal vs Metros Perforados (por día y máquina)',
                     fontweight='bold', fontsize=13)
        ax.set_xlabel('Horas Perforación Normal', fontweight='bold')
        ax.set_ylabel('Metros Perforados (Avance Turno)', fontweight='bold')
        ax.grid(True, linestyle=':', alpha=0.5)
        for s in ['top', 'right']:
            ax.spines[s].set_visible(False)

        leyenda_maquinas = ax.legend(title='Máquina', frameon=False, fontsize=9, loc='upper left')
        ax.add_artist(leyenda_maquinas)
        leyenda_periodo = [
            Line2D([0], [0], marker='o', color='none', markerfacecolor='gray',
                   markeredgecolor='black', markeredgewidth=0.6, markersize=8, alpha=0.9,
                   label=f'Últimos {dias_recientes} días'),
            Line2D([0], [0], marker='o', color='none', markerfacecolor='gray',
                   markeredgecolor='none', markersize=6, alpha=0.3, label='Anterior'),
        ]
        ax.legend(handles=leyenda_periodo, title='Periodo', frameon=False, fontsize=9, loc='lower right')

        fig.tight_layout()
        return fig, ax, scatters_maquina

    def _render_scatter_perfnormal_png(self):
        fig, _, _ = self._draw_scatter_perfnormal_figure()
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def PlotParadas(self):
            self.Mensajes.clear()
            try:
                resumen_paradas = self._build_resumen_paradas()
                self.resumen_paradas = resumen_paradas

                num_maquinas = resumen_paradas['Maquina'].nunique()
                self._draw_paradas_figure(resumen_paradas, figsize=(15, 3.2 * num_maquinas + 1.5))
                plt.show()

                resumen_disp = self._build_resumen_disponibilidad()
                self.resumen_disponibilidad = resumen_disp
                num_maquinas_disp = resumen_disp['Maquina'].nunique()
                self._draw_disponibilidad_figure(resumen_disp, figsize=(15, 3.2 * num_maquinas_disp + 1.5))
                plt.show()

                self._plot_eficiencia_detallada()

                self.Mensajes.setText("Gráficos de Paradas / Eficiencia y Disponibilidad Ploteados")

            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error al Graficar", f"Error crítico: {str(e)}")
                self.Mensajes.setText("Ocurrio algun error")

    def _dibujar_eficiencia_detallada_barras(self, axes, resumen_completo, maquinas, fecha_min, fecha_max):
        """Dibuja, en los `axes` dados (uno por máquina), las barras apiladas al 100% de las
        3 categorías (Horas Cargables / No Cargables sin Perforación Normal / Perforación
        Normal) para el rango [fecha_min, fecha_max]. Cuerpo compartido entre la ventana
        interactiva (_plot_eficiencia_detallada) y la versión estática para Word
        (_render_eficiencia_detallada_png). Devuelve el rango de fechas dibujado."""
        colores = {'cargables': '#5B84C1', 'no_cargables': '#C62828', 'perf_normal': '#2E7D32'}
        rango_visible = pd.date_range(start=fecha_min, end=fecha_max).normalize()

        for ax, maquina in zip(axes, maquinas):
            ax.clear()
            datos_m = resumen_completo[resumen_completo['Maquina'] == maquina]
            datos_m = pd.DataFrame({'Fecha': rango_visible}).merge(datos_m, on='Fecha', how='left')
            cols_pct = ['Pct_Cargables', 'Pct_No_Cargables', 'Pct_Perforacion_Normal']
            datos_m[cols_pct] = datos_m[cols_pct].fillna(0.0)

            ax.bar(datos_m['Fecha'], datos_m['Pct_Cargables'], color=colores['cargables'],
                   width=0.8, label='Horas Cargables')
            ax.bar(datos_m['Fecha'], datos_m['Pct_No_Cargables'], bottom=datos_m['Pct_Cargables'],
                   color=colores['no_cargables'], width=0.8, label='No Cargables (sin Perf. Normal)')
            ax.bar(datos_m['Fecha'], datos_m['Pct_Perforacion_Normal'],
                   bottom=datos_m['Pct_Cargables'] + datos_m['Pct_No_Cargables'],
                   color=colores['perf_normal'], width=0.8, label='Perforación Normal')

            # Etiqueta encima de cada barra con el % que representa la Perforación Normal
            # frente al total del turno (la categoría que más le importa a EDP a simple
            # vista). Rotada 90° porque con ~60 barras por subplot no entra horizontal sin
            # que se encimen entre sí; para que no se corte hay que darle bastante más
            # margen superior del que ocupa la barra (el texto rotado "gasta" a lo alto lo
            # que gastaría a lo ancho si no estuviera rotado).
            alturas = datos_m['Pct_Cargables'] + datos_m['Pct_No_Cargables'] + datos_m['Pct_Perforacion_Normal']
            for x, pct_pn, alto in zip(datos_m['Fecha'], datos_m['Pct_Perforacion_Normal'], alturas):
                if alto > 0:
                    ax.text(x, alto + 3, f'{pct_pn:.0f}%', ha='center', va='bottom',
                            fontsize=5, rotation=90, clip_on=True)

            ax.set_ylim(0, 150)
            ax.set_yticks([0, 25, 50, 75, 100])
            ax.yaxis.set_major_formatter(ticker.PercentFormatter())
            ax.set_title(f'Máquina: {maquina}', fontweight='bold', loc='left', fontsize=10)
            ax.grid(axis='y', linestyle=':', alpha=0.5)
            ax.label_outer()

        axes[-1].xaxis.set_major_formatter(mdates.DateFormatter('%d-%m'))
        paso = max(1, len(rango_visible) // 35)
        axes[-1].set_xticks(rango_visible[::paso])
        plt.setp(axes[-1].get_xticklabels(), rotation=90, fontsize=8)
        return rango_visible

    def _render_eficiencia_detallada_png(self, dias=60):
        """Versión estática (últimos `dias` días, sin slider ni botón) del gráfico de
        eficiencia diaria por máquina, para el reporte Word."""
        resumen_completo = self._build_resumen_paradas_detallado()
        if resumen_completo.empty:
            return None
        maquinas = self._maquinas_validas(resumen_completo)
        num_maquinas = len(maquinas)

        f_max = resumen_completo['Fecha'].max()
        f_min = f_max - pd.Timedelta(days=dias - 1)

        # Tamaño acotado (igual que el resto de imágenes estáticas del Word, ver
        # _figsize_estatico): sin el tope, con 4-5 máquinas la imagen salía tan alta que no
        # cabía debajo del encabezado en la página y Word la empujaba entera a la siguiente,
        # dejando la primera casi en blanco.
        fig, axes = plt.subplots(num_maquinas, 1, figsize=self._figsize_estatico(num_maquinas), sharex=True)
        if num_maquinas == 1:
            axes = [axes]
        plt.subplots_adjust(hspace=0.55, top=0.88)

        self._dibujar_eficiencia_detallada_barras(axes, resumen_completo, maquinas, f_min, f_max)
        handles, labels = axes[0].get_legend_handles_labels()
        fig.legend(handles, labels, loc='upper center', bbox_to_anchor=(0.5, 1.04),
                   ncol=3, frameon=False, fontsize=9)
        fig.suptitle(f'Eficiencia Operativa Diaria por Máquina — Últimos {dias} días (%)',
                     fontweight='bold', fontsize=13, y=1.1)

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def _plot_eficiencia_detallada(self):
        """Gráfico DIARIO por máquina, en %, de las 3 categorías pedidas (Horas Cargables /
        Horas No Cargables sin Perforación Normal / Horas Perforación Normal), con un
        slider Desde-Hasta para acotar el rango visible y un botón para exportar ese mismo
        rango a Excel (reporte por categoría + resumen de 3 categorías)."""
        resumen_completo = self._build_resumen_paradas_detallado()
        if resumen_completo.empty:
            return

        maquinas = self._maquinas_validas(resumen_completo)
        num_maquinas = len(maquinas)

        date_nums = mdates.date2num(resumen_completo['Fecha'].unique())
        d_min, d_max = date_nums.min(), date_nums.max()

        fig, axes = plt.subplots(num_maquinas, 1, figsize=(15, 2.6 * num_maquinas + 2), sharex=True)
        if num_maquinas == 1:
            axes = [axes]
        plt.subplots_adjust(bottom=0.24, hspace=0.5, right=0.82, top=0.88)

        def actualizar(val):
            f_min = mdates.num2date(val[0]).replace(tzinfo=None)
            f_max = mdates.num2date(val[1]).replace(tzinfo=None)
            self._dibujar_eficiencia_detallada_barras(axes, resumen_completo, maquinas, f_min, f_max)

            handles, labels = axes[0].get_legend_handles_labels()
            fig.legend(handles, labels, loc='upper center', bbox_to_anchor=(0.5, 0.95),
                       ncol=3, frameon=False, fontsize=9)
            fig.canvas.draw_idle()

        ax_slider = plt.axes([0.22, 0.08, 0.4, 0.025])
        self.slider_eficiencia = RangeSlider(ax_slider, 'Rango Fecha ', d_min, d_max,
                                              valinit=(max(d_min, d_max - 30), d_max))
        self.slider_eficiencia.on_changed(actualizar)

        self.ax_btn_export_eficiencia = plt.axes([0.85, 0.075, 0.12, 0.045])
        self.btn_export_eficiencia = Button(self.ax_btn_export_eficiencia, 'Exportar Excel',
                                             color='#f0f0f0', hovercolor='skyblue')

        def exportar(event):
            val = self.slider_eficiencia.val
            f_min = mdates.num2date(val[0]).replace(tzinfo=None)
            f_max = mdates.num2date(val[1]).replace(tzinfo=None)
            self._exportar_excel_eficiencia(f_min, f_max)

        self.btn_export_eficiencia.on_clicked(exportar)

        fig.suptitle('Eficiencia Operativa Diaria por Máquina (%)', fontweight='bold', fontsize=14)
        actualizar(self.slider_eficiencia.val)
        plt.show()

    def _exportar_excel_eficiencia(self, fecha_inicio, fecha_fin):
        """Exporta a Excel, para el rango de fechas seleccionado en el slider, el reporte
        detallado de paradas por categoría (medición, tormenta eléctrica, mantenimiento,
        etc.) y el resumen de las 3 categorías en horas y %."""
        try:
            reporte_categorias = self._build_reporte_categorias_paradas(fecha_inicio, fecha_fin)
            reporte_cargables = self._build_reporte_cargables_paradas(fecha_inicio, fecha_fin)
            resumen_3cat = self._build_resumen_paradas_detallado(fecha_inicio, fecha_fin)

            default_name = (
                f"Reporte_Paradas_{fecha_inicio.strftime('%Y-%m-%d')}_a_"
                f"{fecha_fin.strftime('%Y-%m-%d')}.xlsx"
            )
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Guardar Reporte de Paradas", default_name, "Excel Files (*.xlsx)"
            )
            if not file_path:
                return

            with pd.ExcelWriter(file_path) as writer:
                reporte_categorias.to_excel(writer, sheet_name='Detalle por Categoria', index=False)
                reporte_cargables.to_excel(writer, sheet_name='Detalle Cargables', index=False)
                resumen_3cat.to_excel(writer, sheet_name='Resumen 3 Categorias', index=False)

            QMessageBox.information(self, "Éxito", f"Reporte de paradas exportado a:\n{file_path}")
            self.Mensajes.setText("Reporte de paradas exportado")
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "Error al exportar", f"Error crítico: {str(e)}")
            self.Mensajes.setText("Ocurrio algun error al exportar el reporte de paradas")

    def _dibujar_curva_s_serie(self, ax, df, col_y, color, ultima_fecha, es_real, es_porcentaje, meta_total):
        """Dibuja UNA serie de la Curva S (spline suavizado + marcadores + etiquetas semanales
        de los miércoles + etiqueta destacada del punto de corte actual), replicando el estilo
        de la ventana interactiva de PlotS. Reutilizado por las versiones estática en metros y
        en porcentaje que se embeben en el reporte Word."""
        temp = df.dropna(subset=[col_y])
        if temp.empty:
            return
        x_num = mdates.date2num(temp['Fecha'])
        y_val = temp[col_y].values.astype(float)
        if es_porcentaje:
            y_val = y_val / meta_total * 100

        if len(x_num) > 3:
            x_new = np.linspace(x_num.min(), x_num.max(), 300)
            spl = make_interp_spline(x_num, y_val, k=3)
            ax.plot(mdates.num2date(x_new), spl(x_new), color=color, lw=2.5, alpha=0.8, zorder=2)
        else:
            ax.plot(temp['Fecha'], y_val, color=color, lw=2.5, alpha=0.8, zorder=2)

        max_fecha = temp['Fecha'].max()
        sizes = []
        for f in temp['Fecha']:
            if f == ultima_fecha: sizes.append(80)
            elif f == max_fecha and not es_real: sizes.append(60)
            elif f.weekday() == 2: sizes.append(40)
            else: sizes.append(15)
        colores_pts = ['black' if f == ultima_fecha else color for f in temp['Fecha']]
        ax.scatter(temp['Fecha'], y_val, s=sizes, c=colores_pts, edgecolors='white', lw=0.6, zorder=3)

        fmt = (lambda v: f'{v:.0f}%') if es_porcentaje else (lambda v: f'{v:,.0f}')
        for f, v in zip(temp['Fecha'], y_val):
            if f == ultima_fecha:
                offset = (-5 if es_porcentaje else -1400) if not es_real else (4 if es_porcentaje else 1000)
                lbl_color = 'red' if es_real else 'blue'
                ax.text(f, v + offset, fmt(v), color=lbl_color, fontweight='bold', ha='center', fontsize=11)
            elif f == max_fecha and not es_real:
                offset = 3 if es_porcentaje else 900
                ax.text(f, v + offset, fmt(v), color='blue', fontweight='bold', ha='center', fontsize=9)
            elif f.weekday() == 2:
                offset = 3 if es_porcentaje else 800
                ax.text(f, v + offset, fmt(v), color='black', fontsize=8, ha='center', alpha=0.7)

    def _render_curva_s_png(self):
        """Versión estática (sin slider) de la Curva S en metros, con el mismo estilo de
        etiquetas semanales que la ventana interactiva, para el reporte Word."""
        df = self.curva_s
        df_real_valid = df.dropna(subset=["Real_Acumulado"])
        ultima_fecha = df_real_valid["Fecha"].max()
        plan_val = df.loc[df["Fecha"] == ultima_fecha, "Plan_Acumulado"].values[0]
        real_val = df.loc[df["Fecha"] == ultima_fecha, "Real_Acumulado"].values[0]
        cumplimiento = (real_val / plan_val * 100) if plan_val > 0 else 0
        avance_total_pct = (real_val / META_TOTAL_PROYECTO * 100) if META_TOTAL_PROYECTO > 0 else 0

        fig, ax = plt.subplots(figsize=(13, 7))
        self._dibujar_curva_s_serie(ax, df, "Plan_Acumulado", "blue", ultima_fecha,
                                     es_real=False, es_porcentaje=False, meta_total=META_TOTAL_PROYECTO)
        self._dibujar_curva_s_serie(ax, df, "Real_Acumulado", "#d62728", ultima_fecha,
                                     es_real=True, es_porcentaje=False, meta_total=META_TOTAL_PROYECTO)

        texto_box = (
            f"Estado al {ultima_fecha.strftime('%d/%m/%y')}\n"
            f"Plan Acum.: {plan_val:,.0f} mts\n"
            f"Real Acum.: {real_val:,.0f} mts\n"
            f"Cumplimiento: {cumplimiento:.1f}%"
        )
        ax.annotate(texto_box, xy=(0.98, 0.05), xycoords='axes fraction', fontsize=10,
                    ha='right', va='bottom', multialignment='left',
                    bbox=dict(boxstyle="round,pad=0.6", facecolor="white", edgecolor="black", lw=1))

        ax.set_title("Curva S — Avance de Perforación Diamantina", fontweight='bold', fontsize=15, pad=15)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m-%Y'))
        plt.setp(ax.get_xticklabels(), rotation=45)
        ax.set_ylabel("Metros perforados", fontweight='bold')
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
        ax.grid(True, linestyle='--', alpha=0.5)
        for s in ['top', 'right']:
            ax.spines[s].set_visible(False)
        legend_elements = [
            Line2D([0], [0], color='blue', lw=2, label='Plan acumulado', marker='o', markersize=5),
            Line2D([0], [0], color='#d62728', lw=2, label='Real acumulado', marker='o', markersize=5),
        ]
        ax.legend(handles=legend_elements, loc='upper left', frameon=False, fontsize=9)
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf, ultima_fecha, plan_val, real_val, cumplimiento, avance_total_pct

    def _render_curva_s_pct_png(self):
        """Versión estática de la Curva S en porcentaje de la meta total del proyecto
        (META_TOTAL_PROYECTO), con el mismo estilo de etiquetas semanales, para el reporte Word."""
        df = self.curva_s
        df_real_valid = df.dropna(subset=["Real_Acumulado"])
        ultima_fecha = df_real_valid["Fecha"].max()
        plan_val = df.loc[df["Fecha"] == ultima_fecha, "Plan_Acumulado"].values[0]
        real_val = df.loc[df["Fecha"] == ultima_fecha, "Real_Acumulado"].values[0]
        cumplimiento = (real_val / plan_val * 100) if plan_val > 0 else 0

        fig, ax = plt.subplots(figsize=(13, 7))
        self._dibujar_curva_s_serie(ax, df, "Plan_Acumulado", "blue", ultima_fecha,
                                     es_real=False, es_porcentaje=True, meta_total=META_TOTAL_PROYECTO)
        self._dibujar_curva_s_serie(ax, df, "Real_Acumulado", "#d62728", ultima_fecha,
                                     es_real=True, es_porcentaje=True, meta_total=META_TOTAL_PROYECTO)

        texto_box = (
            f"Estado al {ultima_fecha.strftime('%d/%m/%y')}\n"
            f"Plan Acum.: {plan_val:,.0f} mts\n"
            f"Real Acum.: {real_val:,.0f} mts\n"
            f"Cumplimiento: {cumplimiento:.1f}%"
        )
        ax.annotate(texto_box, xy=(0.98, 0.05), xycoords='axes fraction', fontsize=10,
                    ha='right', va='bottom', multialignment='left',
                    bbox=dict(boxstyle="round,pad=0.6", facecolor="white", edgecolor="black", lw=1))

        ax.set_title(f"Curva S — Porcentaje de Avance (Meta: {META_TOTAL_PROYECTO:,.0f} mts)",
                     fontweight='bold', fontsize=15, pad=15)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m-%Y'))
        plt.setp(ax.get_xticklabels(), rotation=45)
        ax.set_ylabel("Porcentaje de Avance (%)", fontweight='bold')
        ax.set_ylim(-5, 110)
        ax.yaxis.set_major_formatter(ticker.PercentFormatter())
        ax.grid(True, linestyle='--', alpha=0.5)
        for s in ['top', 'right']:
            ax.spines[s].set_visible(False)
        legend_elements = [
            Line2D([0], [0], color='blue', lw=2, label='% Plan acumulado'),
            Line2D([0], [0], color='#d62728', lw=2, label='% Real acumulado'),
        ]
        ax.legend(handles=legend_elements, loc='upper left', frameon=False, fontsize=9)
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf

    def _render_semanal_meta_png(self, resumen_semanal_1):
        fig, ax = plt.subplots(figsize=(9, 4.8))
        ax.bar(resumen_semanal_1['Eje_X'], resumen_semanal_1['Suma_Semanal'], color='#5B84C1', width=0.6)
        ax.plot(resumen_semanal_1['Eje_X'], resumen_semanal_1['Meta_Semanal'], color='#B04848',
                linewidth=2.5, marker='o', label='Meta Semanal')
        ax.set_title('Avance Semanal vs Meta Contractual', fontweight='bold')
        ax.set_ylabel('Metros', fontweight='bold')
        plt.setp(ax.get_xticklabels(), rotation=40, ha='right', fontsize=8)
        ax.grid(axis='y', linestyle='--', alpha=0.4)
        ax.legend(loc='upper left', frameon=True, facecolor='white', edgecolor='#555555', fontsize=9)
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf

    def _render_promedio_maquina_png(self, resumen_semanal_2):
        fig, ax = plt.subplots(figsize=(9, 4.8))
        ax.bar(resumen_semanal_2['Eje_X'], resumen_semanal_2['Promedio_Semanal'], color='#5B84C1', width=0.6)
        ax.axhline(y=OBJETIVO_DIARIO_M, color='#B04848', linewidth=2.5, label=f'Objetivo ({OBJETIVO_DIARIO_M} m)')
        ax.set_title('Promedio Semanal por Máquina', fontweight='bold')
        ax.set_ylabel('Metros/día', fontweight='bold')
        plt.setp(ax.get_xticklabels(), rotation=40, ha='right', fontsize=8)
        ax.grid(axis='y', linestyle='--', alpha=0.4)
        ax.legend(loc='upper left', frameon=True, facecolor='white', edgecolor='#555555', fontsize=9)
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf

    def _render_dashboard_por_maquina(self, dias=30):
        """Versión estática del 'Avance metros x Día' (Dashboard), acotada a los últimos `dias`
        días, para el reporte Word. Devuelve una lista [(maquina, BytesIO), ...] con UNA imagen
        por máquina en vez de un solo mosaico con todas juntas: las máquinas ya retiradas no
        tienen actividad en la ventana y antes quedaban como subplots en blanco, restándole
        espacio y claridad al gráfico de las máquinas que sí están operando."""
        df = self.df.copy()
        df['Fecha'] = pd.to_datetime(df['Fecha'])
        df['Dia_dt'] = df['Fecha'].dt.normalize()

        f_max = df['Dia_dt'].max()
        f_min = f_max - pd.Timedelta(days=dias - 1)
        rango_visible = pd.date_range(start=f_min, end=f_max).normalize()

        df_ventana = df[(df['Dia_dt'] >= f_min) & (df['Dia_dt'] <= f_max)]
        maquinas_activas = sorted(df_ventana.loc[df_ventana['Avance Turno'] > 0, 'Maquina'].unique())

        resultados = []
        for maquina in maquinas_activas:
            df_m = df_ventana[df_ventana['Maquina'] == maquina]
            sondajes_unicos = df_m['Sondaje'].unique()
            color_map_sondaje = self._color_map_por_categoria(sondajes_unicos)

            data_pivot = df_m.groupby(['Dia_dt', 'Sondaje'])['Avance Turno'].sum().unstack(fill_value=0)
            data_pivot = data_pivot.reindex(rango_visible, fill_value=0)

            fig, ax = plt.subplots(figsize=(9, 3.4))
            bottom = np.zeros(len(data_pivot))
            sondajes_en_rango = data_pivot.columns[data_pivot.sum() > 0]
            for sondaje in data_pivot.columns:
                values = data_pivot[sondaje].values
                ax.bar(data_pivot.index, values, bottom=bottom,
                       color=color_map_sondaje.get(sondaje),
                       label=sondaje if sondaje in sondajes_en_rango else None,
                       edgecolor='#555555', lw=0.3, width=0.8)
                bottom += values

            totales = data_pivot.sum(axis=1)
            max_avance = totales.max()
            ax.set_ylim(0, max(60, max_avance * 1.15))

            for dia, total in zip(data_pivot.index, totales):
                if total > 0:
                    ax.text(dia, total + max_avance * 0.015, f'{total:.1f}', ha='center', va='bottom',
                            fontsize=6.5, fontweight='bold', rotation=90, clip_on=True)

            handles, labels = ax.get_legend_handles_labels()
            if handles:
                ax.legend(handles, labels, title="Sondajes", bbox_to_anchor=(1.01, 1),
                          loc='upper left', fontsize=8, frameon=False)

            ax.set_title(f'Máquina: {maquina} (últimos {dias} días)', fontweight='bold', fontsize=11)
            ax.axhline(35, color='#C62828', linestyle='--', alpha=0.5)
            ax.grid(True, axis='y', linestyle=':', alpha=0.6)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%m'))
            plt.setp(ax.get_xticklabels(), rotation=45, fontsize=7)
            fig.tight_layout()

            buf = io.BytesIO()
            fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            plt.close(fig)
            buf.seek(0)
            resultados.append((maquina, buf))
        return resultados

    def _render_plano_png(self, df_plot):
        """Versión estática de la Vista en Planta para el reporte Word. A diferencia de la
        ventana interactiva (que usa hover + tarjetas de progreso), aquí solo se etiquetan los
        sondajes en estado 'Perforando' para que el mapa siga siendo legible con muchos puntos;
        el resto de estados se distinguen solo por color y la leyenda. El detalle de avance de
        cada sondaje en ejecución se agrega como tabla junto a la imagen (ver GenerateWordReport)."""
        color_map = {"Programado": "gray", "Ejecutado": "green", "Perforando": "red",
                     "Plataforma": "orange", "Preinicio": "magenta"}

        fig, ax = plt.subplots(figsize=(9, 6))
        plt.subplots_adjust(right=0.97, top=0.90, left=0.10, bottom=0.16)
        ax.set_facecolor('white')
        ax.grid(True, which='major', color='#EEEEEE', linestyle='-', linewidth=0.8, zorder=0)
        ax.minorticks_on()

        self._dibujar_topografia(ax)

        for estado, color in color_map.items():
            sub = df_plot[df_plot["Estado"] == estado]
            if sub.empty:
                continue
            ax.scatter(sub["ESTE_PLOT"], sub["NORTE_PLOT"], c=color, s=25, lw=0.5, zorder=4, label=estado)
            if estado in ("Perforando", "Preinicio"):
                # Etiquetamos TODOS los sondajes en Perforando/Preinicio (normalmente muy
                # pocos a la vez), con plataforma + máquina (ver _etiquetar_puntos_plano);
                # el resto de estados (Programado/Ejecutado/Plataforma) se dejan sin
                # etiqueta para que el mapa siga siendo legible con decenas de puntos.
                self._etiquetar_puntos_plano(ax, sub, fontsize=6.5)

        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.13), ncol=5, title="Estado",
                  frameon=True, fontsize=8)
        fig.suptitle("Estado de Perforación - MREDP 2026", fontsize=13, fontweight='bold')
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%0.0f'))
        ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%0.0f'))
        ax.set_xlabel("Este (m)", fontweight='bold')
        ax.set_ylabel("Norte (m)", fontweight='bold')
        ax.set_xlim(374400, 376800)
        ax.set_ylim(8715600, 8718580)
        ax.set_aspect('equal')

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def _render_progreso_perforando_png(self, df_perforando):
        """Barra de avance por sondaje en perforación (mismo lenguaje visual que las tarjetas de
        progreso de la vista en planta interactiva: barra gris de fondo = metraje total,
        relleno de color = avance actual), para incluir en el reporte Word."""
        df_bar = df_perforando.copy()
        df_bar["Total_m"] = df_bar.apply(
            lambda r: r["LENGTH"] if r["LENGTH"] > 0 else (r["Avance_Actual"] + r["Metro_Faltante"]),
            axis=1
        )
        df_bar["Pct"] = np.where(df_bar["Total_m"] > 0, df_bar["Avance_Actual"] / df_bar["Total_m"] * 100, 0)
        df_bar = df_bar.sort_values("Pct", ascending=True).reset_index(drop=True)
        df_bar["Etiqueta"] = df_bar["Plataforma"] + " / " + df_bar["Sondaje_fmt"]

        fig, ax = plt.subplots(figsize=(9, max(1.2, 0.6 * len(df_bar) + 0.8)))
        y_pos = np.arange(len(df_bar))
        ax.barh(y_pos, df_bar["Total_m"], color="#EEEEEE", height=0.6, label="Metraje Programado")
        colores = ["#2E7D32" if p >= 95 else "#E53935" for p in df_bar["Pct"]]
        ax.barh(y_pos, df_bar["Avance_Actual"], color=colores, height=0.6, label="Avance")

        for i, (avance, total, pct) in enumerate(zip(df_bar["Avance_Actual"], df_bar["Total_m"], df_bar["Pct"])):
            ax.text(total + max(total, 1) * 0.02, i, f"{avance:.1f} / {total:.1f} m ({pct:.1f}%)",
                    va='center', fontsize=8)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(df_bar["Etiqueta"], fontsize=8)
        ax.set_xlabel("Metros", fontweight='bold')
        ax.set_title("Progreso de Perforación — Sondajes Activos", fontweight='bold', fontsize=12)
        ax.legend(loc='lower right', fontsize=8, frameon=True, edgecolor='#555555', facecolor='white')
        ax.grid(axis='x', linestyle='--', alpha=0.4)
        fig.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def GenerateWordReport(self):
            self.Mensajes.clear()
            if not DOCX_AVAILABLE:
                QMessageBox.critical(self, "Falta dependencia",
                    "No se encontró la librería 'python-docx'.\nInstálala con:\n\npip install python-docx")
                return

            if not hasattr(self, 'curva_s') or not hasattr(self, 'df_final'):
                QMessageBox.warning(self, "Faltan datos", "Primero carga el Plan de Perforación (Curva S).")
                return

            try:
                df_clean, machine_cols = self._build_df_clean()
                resumen_semanal_1, resumen_semanal_2 = self._build_resumen_semanal(df_clean, machine_cols)
                img_curva, ultima_fecha, plan_val, real_val, cumplimiento, avance_total_pct = self._render_curva_s_png()
                img_curva_pct = self._render_curva_s_pct_png()
                img_semanal = self._render_semanal_meta_png(resumen_semanal_1)
                img_promedio = self._render_promedio_maquina_png(resumen_semanal_2)
                img_eficiencia_detallada = self._render_eficiencia_detallada_png()
                img_scatter_perfnormal = self._render_scatter_perfnormal_png()

                doc = Document()
                titulo = doc.add_heading('Reporte Semanal de Avance de Perforación', level=0)
                titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sub = doc.add_paragraph(f"MREDP 2026 — Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}")
                sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

                doc.add_heading('Resumen General', level=1)
                maquinas_activas = ', '.join(m for m, cfg in self.machine_lifecycle.items() if cfg['end'] is None)
                tabla = doc.add_table(rows=1, cols=2)
                tabla.style = 'Light Grid Accent 1'
                hdr = tabla.rows[0].cells
                hdr[0].text = 'Indicador'
                hdr[1].text = 'Valor'
                for k, v in [
                    ('Fecha de corte', ultima_fecha.strftime('%d/%m/%Y')),
                    ('Plan acumulado', f'{plan_val:,.0f} m'),
                    ('Real acumulado', f'{real_val:,.0f} m'),
                    ('% Cumplimiento (Real vs Plan a la fecha)', f'{cumplimiento:.1f}%'),
                    ('Avance Total del Proyecto', f'{avance_total_pct:.1f}% ({real_val:,.0f} de {META_TOTAL_PROYECTO:,.0f} m)'),
                    ('Máquinas activas', maquinas_activas or '-'),
                ]:
                    row = tabla.add_row().cells
                    row[0].text = k
                    row[1].text = v

                resumen_taladros = self._resumen_taladros()
                if resumen_taladros is not None:
                    perforados, planeados = resumen_taladros
                    pct_taladros = (perforados / planeados * 100) if planeados > 0 else 0
                    row = tabla.add_row().cells
                    row[0].text = 'Taladros Perforados'
                    row[1].text = f'{perforados} de {planeados} ({pct_taladros:.0f}%)'

                doc.add_heading('Curva S — Plan vs Real', level=1)
                doc.add_picture(img_curva, width=Inches(6.3))
                doc.add_paragraph(
                    f"Al {ultima_fecha.strftime('%d/%m/%Y')} se registran {real_val:,.0f} m reales frente a "
                    f"{plan_val:,.0f} m planificados ({cumplimiento:.1f}% de cumplimiento respecto al plan a la fecha). "
                    f"Frente a la meta total del proyecto ({META_TOTAL_PROYECTO:,.0f} m), el avance representa un "
                    f"{avance_total_pct:.1f}% de perforación."
                )
                doc.add_picture(img_curva_pct, width=Inches(6.3))

                doc.add_heading('Avance Semanal vs Meta Contractual', level=1)
                doc.add_picture(img_semanal, width=Inches(6.3))
                ultima_semana = resumen_semanal_1.iloc[-1]
                cumpl_semana = (ultima_semana['Suma_Semanal'] / ultima_semana['Meta_Semanal'] * 100) \
                    if ultima_semana['Meta_Semanal'] > 0 else 0
                doc.add_paragraph(
                    f"En la última semana ({ultima_semana['Eje_X']}) se perforaron "
                    f"{ultima_semana['Suma_Semanal']:,.1f} m frente a una meta de "
                    f"{ultima_semana['Meta_Semanal']:,.0f} m ({cumpl_semana:.1f}% de la meta semanal)."
                )

                tabla2 = doc.add_table(rows=1, cols=3)
                tabla2.style = 'Light Grid Accent 1'
                hdr2 = tabla2.rows[0].cells
                hdr2[0].text, hdr2[1].text, hdr2[2].text = 'Semana', 'Metros Perforados', 'Meta (m)'
                for _, r in resumen_semanal_1.tail(8).iterrows():
                    row = tabla2.add_row().cells
                    row[0].text = str(r['Eje_X'])
                    row[1].text = f"{r['Suma_Semanal']:,.1f}"
                    row[2].text = f"{r['Meta_Semanal']:,.0f}"

                doc.add_heading('Promedio Semanal por Máquina', level=1)
                doc.add_picture(img_promedio, width=Inches(6.3))
                doc.add_paragraph(
                    f"El promedio semanal de perforación por máquina activa fue de "
                    f"{resumen_semanal_2['Promedio_Semanal'].mean():.1f} m/día frente al objetivo de "
                    f"{OBJETIVO_DIARIO_M} m/día."
                )

                doc.add_heading('Avance Diario por Máquina (Último Mes)', level=1)
                doc.add_paragraph(
                    "Metros perforados por día y por sondaje en cada máquina durante los últimos "
                    "30 días. La línea roja marca la meta diaria de referencia (35 m). Solo se "
                    "incluyen las máquinas con actividad en este periodo."
                )
                dashboards_por_maquina = self._render_dashboard_por_maquina(dias=30)
                for maquina, img_maquina in dashboards_por_maquina:
                    doc.add_heading(maquina, level=2)
                    doc.add_picture(img_maquina, width=Inches(6.3))

                doc.add_heading('Eficiencia Operativa Diaria por Máquina', level=1)
                if img_eficiencia_detallada is not None:
                    doc.add_picture(img_eficiencia_detallada, width=Inches(6.3))
                    resumen_detallado_word = self._build_resumen_paradas_detallado()
                    prom = resumen_detallado_word[
                        ['Pct_Cargables', 'Pct_No_Cargables', 'Pct_Perforacion_Normal']
                    ].mean()
                    doc.add_paragraph(
                        f"Del total de horas reportadas por turno en todo el periodo, en promedio "
                        f"{prom['Pct_Cargables']:.1f}% corresponden a paradas facturables (Horas "
                        f"Cargables: Pruebas de Medición, Espera por tormenta eléctrica, etc.), "
                        f"{prom['Pct_No_Cargables']:.1f}% a tiempo muerto no facturable y "
                        f"{prom['Pct_Perforacion_Normal']:.1f}% a Perforación Normal efectiva. El "
                        f"gráfico muestra el detalle día a día por máquina de los últimos 60 días."
                    )
                else:
                    doc.add_paragraph("No se encontraron datos suficientes para este gráfico.")

                doc.add_heading('Horas de Perforación Normal vs Metros Perforados', level=1)
                doc.add_picture(img_scatter_perfnormal, width=Inches(6.3))
                doc.add_paragraph(
                    "Cada punto es un día de una máquina: cuántas horas de Perforación Normal tomó "
                    "lograr los metros perforados ese día. Una nube de puntos más arriba (o más a la "
                    "izquierda) que las demás indica una máquina que rinde más metros por hora."
                )

                if hasattr(self, 'BD_Union_total'):
                    doc.add_heading('Vista en Planta — Sondajes en Ejecución', level=1)
                    df_plot_plano = self._build_df_plot_plano()
                    img_plano = self._render_plano_png(df_plot_plano)
                    doc.add_picture(img_plano, width=Inches(6.3))
                    doc.add_paragraph(
                        "Solo se etiquetan los sondajes en estado 'Perforando'; el resto de "
                        "plataformas se distingue por color (ver leyenda) para mantener el mapa "
                        "legible. El detalle de avance de cada sondaje en ejecución se muestra "
                        "en la tabla siguiente."
                    )

                    df_perforando = df_plot_plano[df_plot_plano["Estado"] == "Perforando"].copy()
                    if not df_perforando.empty:
                        df_perforando["Total_m"] = df_perforando.apply(
                            lambda r: r["LENGTH"] if r["LENGTH"] > 0 else (r["Avance_Actual"] + r["Metro_Faltante"]),
                            axis=1
                        )
                        df_perforando["Pct_Avance"] = np.where(
                            df_perforando["Total_m"] > 0,
                            df_perforando["Avance_Actual"] / df_perforando["Total_m"] * 100, 0
                        )

                        img_progreso = self._render_progreso_perforando_png(df_perforando)
                        doc.add_picture(img_progreso, width=Inches(6.3))

                        tabla3 = doc.add_table(rows=1, cols=6)
                        tabla3.style = 'Light Grid Accent 1'
                        hdr3 = tabla3.rows[0].cells
                        for c, titulo_col in zip(hdr3, ['Plataforma', 'Sondaje', 'Máquina', '% Avance', 'Días', 'Proy. Restante (d)']):
                            c.text = titulo_col
                        for _, r in df_perforando.sort_values("Pct_Avance", ascending=False).iterrows():
                            row = tabla3.add_row().cells
                            dias_val = int(r['Dias_Perforando_Count']) if pd.notna(r['Dias_Perforando_Count']) else 0
                            row[0].text = str(r['Plataforma'])
                            row[1].text = str(r['Sondaje_fmt'])
                            row[2].text = str(r['Maquina_fmt'])
                            row[3].text = f"{r['Pct_Avance']:.1f}%"
                            row[4].text = str(dias_val)
                            row[5].text = str(r['Proyeccion_fmt'])
                    else:
                        doc.add_paragraph("No hay sondajes en estado 'Perforando' actualmente.")

                    doc.add_heading('Movimientos de Máquina entre Sondajes', level=1)
                    movimientos = self._build_movimientos_maquina()
                    if not movimientos.empty:
                        doc.add_paragraph(
                            "Cada fila es un cambio de sondaje detectado para una máquina (terminó "
                            "uno y pasó a otro). La distancia es un APROXIMADO: se calcula en línea "
                            "recta entre las coordenadas de collar de ambos sondajes, no es la "
                            "distancia real que recorrió la máquina (que depende del camino/accesos "
                            "de la plataforma)."
                        )
                        resumen_mov = movimientos.groupby('Maquina').agg(
                            Movimientos=('Distancia_Aprox_m', 'size'),
                            Distancia_Total_Aprox_m=('Distancia_Aprox_m', 'sum'),
                        ).reset_index()

                        tabla_mov_resumen = doc.add_table(rows=1, cols=3)
                        tabla_mov_resumen.style = 'Light Grid Accent 1'
                        hdr_mr = tabla_mov_resumen.rows[0].cells
                        for c, titulo_col in zip(hdr_mr, ['Máquina', 'N° Movimientos', 'Distancia Total Aprox. (m)']):
                            c.text = titulo_col
                        for _, r in resumen_mov.iterrows():
                            row = tabla_mov_resumen.add_row().cells
                            row[0].text = str(r['Maquina'])
                            row[1].text = str(int(r['Movimientos']))
                            row[2].text = f"{r['Distancia_Total_Aprox_m']:,.0f}"

                        tabla_mov = doc.add_table(rows=1, cols=5)
                        tabla_mov.style = 'Light Grid Accent 1'
                        hdr_m = tabla_mov.rows[0].cells
                        for c, titulo_col in zip(hdr_m, ['Fecha', 'Máquina', 'De (Sondaje)', 'A (Sondaje)', 'Distancia Aprox. (m)']):
                            c.text = titulo_col
                        for _, r in movimientos.iterrows():
                            row = tabla_mov.add_row().cells
                            row[0].text = r['Fecha'].strftime('%d/%m/%Y')
                            row[1].text = str(r['Maquina'])
                            row[2].text = str(r['Sondaje_Origen'])
                            row[3].text = str(r['Sondaje_Destino'])
                            row[4].text = f"{r['Distancia_Aprox_m']:,.0f}" if pd.notna(r['Distancia_Aprox_m']) else "s/coordenadas"
                    else:
                        doc.add_paragraph("No se detectaron cambios de sondaje en el periodo cargado.")

                if hasattr(self, 'almacen_data'):
                    doc.add_heading('Avance de Trabajos en Almacén', level=1)
                    doc.add_paragraph(
                        "Metros (o número de muestras, en el caso de Muestreo) procesados por "
                        "etapa: recepción, fotografía, logueo geológico, logueo geotécnico, "
                        "corte y muestreo."
                    )

                    resumen_total_almacen = self._build_almacen_resumen_total()
                    doc.add_picture(self._render_almacen_resumen_png(resumen_total_almacen), width=Inches(6.3))
                    doc.add_paragraph(
                        "Vista rápida del metraje acumulado en cada etapa a la fecha. El "
                        "'WIP' bajo cada flecha es la diferencia entre etapas consecutivas: "
                        "cuánto ya pasó por una etapa pero todavía no llega a la siguiente."
                    )

                    doc.add_heading('Diagrama de Flujo Acumulativo (CFD)', level=2)
                    doc.add_paragraph(
                        "Cada línea es el metraje acumulado de una etapa; el área sombreada "
                        "entre dos líneas consecutivas es el WIP entre ellas — mientras más "
                        "ancha la banda, más metraje esperando pasar de una etapa a la "
                        "siguiente."
                    )
                    cfd_sem = self._build_almacen_cfd(freq='W-TUE')
                    if not cfd_sem.empty:
                        doc.add_picture(self._render_almacen_cfd_png(cfd_sem, freq='W-TUE'), width=Inches(6.3))
                    cfd_mes = self._build_almacen_cfd(freq='M')
                    if not cfd_mes.empty:
                        doc.add_picture(self._render_almacen_cfd_png(cfd_mes, freq='M'), width=Inches(6.3))

                    resumen_almacen_sem = self._build_resumen_almacen(freq='W-TUE')
                    if not resumen_almacen_sem.empty:
                        doc.add_heading('Vista Semanal (miércoles a martes)', level=2)
                        img_almacen_sem = self._render_almacen_png(resumen_almacen_sem, freq='W-TUE')
                        doc.add_picture(img_almacen_sem, width=Inches(6.3))

                        doc.add_heading('Vista Mensual', level=2)
                        resumen_almacen_mes = self._build_resumen_almacen(freq='M')
                        img_almacen_mes = self._render_almacen_png(resumen_almacen_mes, freq='M')
                        doc.add_picture(img_almacen_mes, width=Inches(6.3))
                    else:
                        doc.add_paragraph("No se encontraron datos de avance de almacén para graficar.")

                    wip_sem = self._build_almacen_wip(freq='W-TUE')
                    if not wip_sem.empty:
                        doc.add_heading('Trabajo en Proceso (WIP) por Transición', level=2)
                        doc.add_paragraph(
                            "Metros que ya pasaron por una etapa pero aún no llegan a la "
                            "siguiente (acumulado, no solo lo del periodo). Una transición con "
                            "WIP creciente indica un cuello de botella formándose ahí."
                        )
                        doc.add_picture(self._render_almacen_wip_png(wip_sem, freq='W-TUE'), width=Inches(6.3))
                        wip_mes = self._build_almacen_wip(freq='M')
                        doc.add_picture(self._render_almacen_wip_png(wip_mes, freq='M'), width=Inches(6.3))

                    ciclo_sem = self._build_almacen_ciclo(freq='W-TUE')
                    if not ciclo_sem.empty:
                        doc.add_heading('Tiempo de Ciclo entre Etapas', level=2)
                        doc.add_paragraph(
                            "Días promedio entre la última fecha reportada de un sondaje en "
                            "una etapa y en la siguiente. Un tiempo de ciclo que aumenta con el "
                            "tiempo en una transición específica es otra señal de cuello de "
                            "botella, independiente del WIP acumulado."
                        )
                        doc.add_picture(self._render_almacen_ciclo_png(ciclo_sem, freq='W-TUE'), width=Inches(6.3))
                        ciclo_mes = self._build_almacen_ciclo(freq='M')
                        doc.add_picture(self._render_almacen_ciclo_png(ciclo_mes, freq='M'), width=Inches(6.3))

                doc.add_heading('Brocas — Metros por Diámetro y Tramo de Vida Útil', level=1)
                resumen_brocas = self._build_resumen_brocas()
                if not resumen_brocas.empty:
                    doc.add_picture(self._render_brocas_bar_png(resumen_brocas), width=Inches(6.3))
                    doc.add_paragraph(
                        "Cada turno se atribuye al tramo de 100 m en que se encontraba el metraje "
                        "acumulado de la broca usada ese día (BROCAS, A29:F32 de cada hoja de turno), "
                        "prorrateando los metros si el turno cruza el límite entre dos tramos. El "
                        "Tricono se muestra como categoría aparte, sin desglose por tramo. Los meses "
                        "van del día 20 del mes anterior al 19 del mes indicado (excepto Diciembre, "
                        "que se extiende hasta el 31, y Enero, que empieza el 01)."
                    )

                    def _agregar_tabla_brocas(titulo, tabla_pivote):
                        doc.add_heading(titulo, level=2)
                        if tabla_pivote.empty:
                            doc.add_paragraph("Sin datos de brocas para esta máquina en el periodo.")
                            return
                        tabla_doc = doc.add_table(rows=1, cols=len(tabla_pivote.columns) + 1)
                        tabla_doc.style = 'Light Grid Accent 1'
                        hdr = tabla_doc.rows[0].cells
                        hdr[0].text = 'Categoría'
                        for j, col in enumerate(tabla_pivote.columns, start=1):
                            hdr[j].text = str(col)
                        for categoria, fila_tabla in tabla_pivote.iterrows():
                            row = tabla_doc.add_row().cells
                            row[0].text = categoria
                            for j, valor in enumerate(fila_tabla, start=1):
                                row[j].text = f'{valor:,.1f}' if valor else '-'

                    _agregar_tabla_brocas('Global (todas las máquinas)', self._tabla_brocas_pivote(resumen_brocas))
                    for maquina in self._maquinas_validas():
                        _agregar_tabla_brocas(
                            f'Máquina: {maquina}', self._tabla_brocas_pivote(resumen_brocas, maquina=maquina)
                        )
                else:
                    doc.add_paragraph("No se encontraron datos de brocas (tabla BROCAS) en los reportes cargados.")

                fecha_hoy = datetime.now().strftime("%Y-%m-%d")
                default_name = f"Reporte_Semanal_{fecha_hoy}.docx"
                file_path, _ = QFileDialog.getSaveFileName(
                    self, "Guardar Reporte Semanal", default_name, "Word Files (*.docx)"
                )
                if not file_path:
                    return

                doc.save(file_path)
                QMessageBox.information(self, "Éxito", f"Reporte guardado en:\n{file_path}")
                self.Mensajes.setText("Reporte Semanal en Word generado")

            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error al generar reporte", f"Error crítico: {str(e)}")
                self.Mensajes.setText("Ocurrio algun error al generar el reporte")

    def LoadDesv(self):
            self.GrafDesv.setEnabled(False)
            self.Mensajes.clear()
            """
            Selecciona un archivo CSV y realiza el procesamiento de trayectorias
            comparando los datos de entrada y salida (In/Out).
            """
            # 1. Seleccionar el archivo CSV
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Seleccionar archivo de Control Standard Rep", "", "CSV Files (*.csv)"
            )
            
            if not file_path:
                return
    
            try:
                # 2. Cargar el dataframe
                df = pd.read_csv(file_path)
                
                # --- PROCESAMIENTO DE DATOS ---
                # Conversión de hora y ordenamiento
                df["Hora_dt"] = pd.to_datetime(df["Hora"], format="%H:%M:%S", errors="coerce")
                df = df.sort_values(by="Hora_dt", ascending=True).reset_index(drop=True)
    
                # Redondeo para asegurar coincidencia en el merge
                df['Profundidad'] = df['Profundidad'].round(1)
    
                # Identificar punto de retorno (Profundidad Máxima)
                idx_max = df['Profundidad'].idxmax()
                df_entrada = df.iloc[:idx_max + 1].copy()
                df_salida = df.iloc[idx_max:].copy()
    
                # Cruce de datos (Merge) para comparar entrada vs salida
                self.df_comp = pd.merge(
                    df_entrada[['Profundidad', 'Norte', 'Este', 'Elevación']], 
                    df_salida[['Profundidad', 'Norte', 'Este', 'Elevación']], 
                    on='Profundidad', 
                    suffixes=('_in', '_out')
                )
    
                # Cálculo de distancia (Error 3D Euclidiano)
                self.df_comp['Dist'] = np.sqrt(
                    (self.df_comp['Norte_in'] - self.df_comp['Norte_out'])**2 +
                    (self.df_comp['Este_in'] - self.df_comp['Este_out'])**2 +
                    (self.df_comp['Elevación_in'] - self.df_comp['Elevación_out'])**2
                )
    
                # --- CÁLCULO DE MÉTRICAS CON VALIDACIÓN ---
    
                # 1. Error de Cierre (Metraje 0)
                filas_cierre = self.df_comp[self.df_comp['Profundidad'] == 0]
    
                if not filas_cierre.empty:
                    self.error_cierre = filas_cierre['Dist'].iloc[0]
                    msg_cierre = f"{self.error_cierre:.3f} m"
                else:
                    self.error_cierre = 0.0
                    msg_cierre = "No encontrado"
                    print("⚠️ Aviso: No se encontró metraje '0'.")
    
                # 2. Desviación Máxima
                if not self.df_comp.empty:
                    self.error_max = self.df_comp['Dist'].max()
                    msg_max = f"{self.error_max:.3f} m"
                else:
                    self.error_max = 0.0
                    msg_max = "Sin datos"
    
                # Éxito: Guardamos el dataframe original por si se necesita
                self.df_csv = df
                
                # Mensaje de confirmación
                mensaje_exito = (f"Procesamiento completado:\n\n"
                                f"Archivo: {Path(file_path).name}\n"
                                f"Error de Cierre: {msg_cierre}\n"
                                f"Desviación Máxima: {msg_max}")
                
                QMessageBox.information(self, "Proceso Exitoso", mensaje_exito)
                self.Mensajes.setText(f"Archivo cargado {mensaje_exito}")
                self.GrafDesv.setEnabled(True)
    
            except Exception as e:
                print("\n" + "!"*50)
                print("ERROR EN LOADCSV:")
                traceback.print_exc()
                print("!"*50 + "\n")
                QMessageBox.critical(self, "Error de Procesamiento", f"No se pudo procesar el CSV:\n{str(e)}")
                self.Mensajes.setText("Ocurrio algun error, revisar que el archivo sea el correcto")
                
    def PlotDesv(self):
            self.Mensajes.clear()
            try:
                # 1. VALIDACIÓN DE DATOS
                if not hasattr(self, 'df_csv') or self.df_csv.empty:
                    QMessageBox.warning(self, "Faltan Datos", "Primero debe cargar y procesar un archivo CSV.")
                    return
    
                df = self.df_csv
                # Recuperamos los dataframes de entrada/salida calculados en LoadCSV
                idx_max = df['Profundidad'].idxmax()
                df_entrada = df.iloc[:idx_max + 1].copy()
                df_salida = df.iloc[idx_max:].copy()
                
                # Identificador del pozo
                id_hoyo = df['ID del Hoyo'].iloc[0] if 'ID del Hoyo' in df.columns else "N/A"
    
                # ======================================================
                # FIGURA 1: TRAYECTORIA 3D (COMPARATIVA)
                # ======================================================
                fig1 = plt.figure(figsize=(10, 8))
                ax3d = fig1.add_subplot(111, projection='3d')
    
                # Trayectoria Planeada (Ideal vertical)
                z_planeado = -df_entrada['Profundidad']
                ax3d.plot(np.zeros(len(z_planeado)), np.zeros(len(z_planeado)), z_planeado, 
                        label='Planeado (Dip -90°)', color='black', linestyle='--', linewidth=2, alpha=0.7)
    
                # Entrada y Salida
                ax3d.plot(df_entrada['Este'], df_entrada['Norte'], df_entrada['Elevación'], 
                        label='Entrada (Bajada)', color='#2ECC71', linewidth=2, marker='o', markersize=2)
                ax3d.plot(df_salida['Este'], df_salida['Norte'], df_salida['Elevación'], 
                        label='Salida (Subida)', color='#E74C3C', linewidth=2, marker='o', markersize=2)
    
                ax3d.set_xlabel('Este (m)')
                ax3d.set_ylabel('Norte (m)')
                ax3d.set_zlabel('Elevación (m)')
                ax3d.set_title(f"Trayectoria 3D - Pozo: {id_hoyo}", pad=20)
                ax3d.set_box_aspect([1, 1, 2])
    
                # Lógica de Botones (Guardados en self para evitar que desaparezcan)
                self.ax_btn_iso = plt.axes([0.55, 0.05, 0.1, 0.04])
                self.ax_btn_plan = plt.axes([0.66, 0.05, 0.1, 0.04])
                self.ax_btn_der = plt.axes([0.77, 0.05, 0.1, 0.04])
    
                self.btn_iso = Button(self.ax_btn_iso, 'ISO', color='#f0f0f0', hovercolor='skyblue')
                self.btn_plan = Button(self.ax_btn_plan, 'Planta', color='#f0f0f0', hovercolor='skyblue')
                self.btn_der = Button(self.ax_btn_der, 'Derecha', color='#f0f0f0', hovercolor='skyblue')
    
                self.btn_iso.on_clicked(lambda e: (ax3d.view_init(elev=30, azim=45), fig1.canvas.draw_idle()))
                self.btn_plan.on_clicked(lambda e: (ax3d.view_init(elev=90, azim=-90), fig1.canvas.draw_idle()))
                self.btn_der.on_clicked(lambda e: (ax3d.view_init(elev=0, azim=0), fig1.canvas.draw_idle()))
    
                info_txt = f"Error Cierre: {getattr(self, 'error_cierre', 0):.3f} m\nDesv. Máx: {getattr(self, 'error_max', 0):.3f} m"
                fig1.text(0.12, 0.12, info_txt, fontsize=9, bbox=dict(facecolor='white', alpha=0.8, boxstyle='round'))
                ax3d.legend()
    
                # ======================================================
                # FIGURA 2: ANÁLISIS QA/QC INCLINACIÓN (2D)
                # ======================================================
                fig2, ax2d = plt.subplots(figsize=(12, 6))
    
                def get_smooth_data(sub_df):
                    clean_df = sub_df.drop_duplicates(subset=['Profundidad']).sort_values('Profundidad')
                    x = clean_df['Profundidad'].values
                    y = clean_df['Inclinación'].values
                    x_new = np.linspace(x.min(), x.max(), 300)
                    spl = make_interp_spline(x, y, k=3)
                    return x_new, spl(x_new)
    
                # Alertas
                df_alerts = df_entrada.copy().reset_index(drop=True)
                df_alerts['diff_10m'] = df_alerts['Inclinación'].diff().abs()
                df_alerts['diff_100m'] = df_alerts['Inclinación'].diff(periods=10).abs()
                al_10m = df_alerts[df_alerts['diff_10m'] > 1.0]
                al_100m = df_alerts[df_alerts['diff_100m'] > 1.0]
    
                ax2d.axhline(y=-90, color='black', linestyle='--', alpha=0.3, label='Target -90°')
                
                # Graficar líneas suavizadas
                xi, yi = get_smooth_data(df_entrada)
                xo, yo = get_smooth_data(df_salida)
                line_in, = ax2d.plot(xi, yi, color='#2ECC71', lw=3, label='Entrada')
                line_out, = ax2d.plot(xo, yo, color='#E74C3C', lw=2, ls=':', label='Salida')
    
                # Scatter de alertas
                s10 = ax2d.scatter(al_10m['Profundidad'], al_10m['Inclinación'], color='orange', marker='x', label='Alerta >1° (10m)', zorder=5)
                s100 = ax2d.scatter(al_100m['Profundidad'], al_100m['Inclinación'], color='red', marker='x', s=100, lw=3, label='Alerta >1° (100m)', zorder=6)
    
                # Cursor Interactivo
                cursor = mplcursors.cursor([line_in, line_out, s10, s100], hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="black")
                    sel.annotation.set_text(f"Prof: {sel.target[0]:.2f}m\nInc: {sel.target[1]:.2f}°")
    
                ax2d.set_title(f"QA/QC Inclinación - {id_hoyo}")
                ax2d.set_xlabel("Profundidad (m)"); ax2d.set_ylabel("Inclinación (°)")
                ax2d.set_ylim(-95, -85)
                ax2d.grid(True, alpha=0.3)
                ax2d.legend(loc='lower right', fontsize=8)
    
                plt.tight_layout()
                plt.show()
                self.Mensajes.setText("Ploteo de desviacion realizada")
    
            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error", f"Error al graficar desviaciones: {e}")
                self.Mensajes.setText("ocurrio algun error")

    def _emparejar_nombres_maquina(self, nombres_a, nombres_b):
        """Empareja nombres de máquina entre dos fuentes que no siempre coinciden exactamente
        (ej. "JINSHI YN-1500" en los reportes de turno vs "YN-1500" en el Excel de Almacén).
        Primero intenta coincidencia exacta; lo que no calza, lo empareja si el nombre
        normalizado (sin tildes/mayúsculas/espacios) de uno contiene al del otro. Devuelve un
        dict nombre_a -> nombre_b (o None si no se encontró equivalencia)."""
        resultado = {}
        usados_b = set()
        for a in nombres_a:
            if a in nombres_b:
                resultado[a] = a
                usados_b.add(a)
        for a in nombres_a:
            if a in resultado:
                continue
            a_norm = _strip_accents(a).upper().replace(" ", "")
            mejor = None
            for b in nombres_b:
                if b in usados_b:
                    continue
                b_norm = _strip_accents(b).upper().replace(" ", "")
                if a_norm and b_norm and (a_norm in b_norm or b_norm in a_norm):
                    mejor = b
                    break
            resultado[a] = mejor
            if mejor:
                usados_b.add(mejor)
        return resultado

    def _validar_metros_vs_almacen(self, tolerancia=0.05):
        """Compara, para cada Fecha+Máquina, los metros calculados desde los reportes de
        turno (self.df_final) contra los reportados independientemente en la hoja
        'Perforación por Equipo' del Excel de Almacén (self.almacen_perforacion_equipo).
        Marca "DIFERENCIA" cuando la diferencia absoluta supera `tolerancia` (m)."""
        maquinas_reportes = [c for c in self.df_final.columns if c not in ('Fecha', 'Total Dia')]
        maquinas_almacen = sorted(self.almacen_perforacion_equipo['Maquina'].dropna().unique())
        mapa = self._emparejar_nombres_maquina(maquinas_reportes, maquinas_almacen)

        filas = []
        for maquina_rep, maquina_alm in mapa.items():
            serie = self.df_final[['Fecha', maquina_rep]].rename(columns={maquina_rep: 'Metros_Reportes'})
            if maquina_alm is None:
                serie['Metros_Almacen'] = np.nan
                serie['Estado'] = 'Sin equivalencia en Almacén'
            else:
                serie_alm = self.almacen_perforacion_equipo[
                    self.almacen_perforacion_equipo['Maquina'] == maquina_alm
                ][['Fecha', 'Metros']].rename(columns={'Metros': 'Metros_Almacen'})
                serie = serie.merge(serie_alm, on='Fecha', how='outer')
                serie['Metros_Reportes'] = serie['Metros_Reportes'].fillna(0.0)
                serie['Metros_Almacen'] = serie['Metros_Almacen'].fillna(0.0)
                diferencia = (serie['Metros_Reportes'] - serie['Metros_Almacen']).abs()
                serie['Estado'] = np.where(diferencia > tolerancia, 'DIFERENCIA', 'OK')
            serie['Maquina_Reportes'] = maquina_rep
            serie['Maquina_Almacen'] = maquina_alm if maquina_alm is not None else '(sin equivalencia)'
            filas.append(serie)

        columnas = ['Fecha', 'Maquina_Reportes', 'Maquina_Almacen',
                    'Metros_Reportes', 'Metros_Almacen', 'Estado']
        if not filas:
            return pd.DataFrame(columns=columnas + ['Diferencia'])

        validacion = pd.concat(filas, ignore_index=True)
        validacion['Diferencia'] = validacion['Metros_Reportes'] - validacion['Metros_Almacen']
        validacion = validacion[columnas[:-1] + ['Diferencia', 'Estado']]
        return validacion.sort_values(['Maquina_Reportes', 'Fecha']).reset_index(drop=True)

    def _build_matriz_diferencias_almacen(self, validacion):
        """Matriz Fecha (filas) x Máquina (columnas) con la Diferencia (Metros_Reportes -
        Metros_Almacen) de cada combinación, para escanear de un vistazo dónde corregir
        metros en vez de recorrer la hoja larga fila por fila (ver
        _aplicar_formato_matriz_diferencias, que resalta en rojo lo que está fuera de
        tolerancia directamente sobre esta matriz)."""
        if validacion.empty:
            return pd.DataFrame()
        matriz = validacion.pivot_table(
            index='Fecha', columns='Maquina_Reportes', values='Diferencia', aggfunc='sum'
        )
        return matriz.sort_index()

    def _aplicar_formato_matriz_diferencias(self, hoja, matriz_export, tolerancia=0.05):
        """Resalta en rojo las celdas de la matriz Fecha x Máquina cuya diferencia (Reportes
        - Almacén) supera `tolerancia` en valor absoluto, y en verde las que sí cuadran —
        para poder ver de un vistazo dónde hay que corregir metros, sin leer la hoja larga
        'Validacion vs Almacen' fila por fila."""
        from openpyxl.styles import PatternFill
        from openpyxl.formatting.rule import FormulaRule
        from openpyxl.utils import get_column_letter

        n_filas, n_cols = matriz_export.shape
        if n_filas == 0 or n_cols <= 1:
            return

        primera_fila, ultima_fila = 2, n_filas + 1
        col_b, col_ultima = get_column_letter(2), get_column_letter(n_cols)
        rango = f"{col_b}{primera_fila}:{col_ultima}{ultima_fila}"
        celda_ref = f"{col_b}{primera_fila}"

        rojo = PatternFill(start_color='F8CBCB', end_color='F8CBCB', fill_type='solid')
        verde = PatternFill(start_color='C6E7C8', end_color='C6E7C8', fill_type='solid')
        hoja.conditional_formatting.add(
            rango, FormulaRule(formula=[f'ABS({celda_ref})>{tolerancia}'], fill=rojo)
        )
        hoja.conditional_formatting.add(
            rango, FormulaRule(formula=[f'AND({celda_ref}<>"",ABS({celda_ref})<={tolerancia})'], fill=verde)
        )

        hoja.column_dimensions['A'].width = 12
        for col_idx in range(2, n_cols + 1):
            hoja.column_dimensions[get_column_letter(col_idx)].width = 15
        for row in hoja.iter_rows(min_row=primera_fila, max_row=ultima_fila, min_col=2, max_col=n_cols):
            for celda in row:
                if isinstance(celda.value, (int, float)):
                    celda.number_format = '0.00'

    def _actualizar_revision_con_validacion_almacen(self):
        """(Re)escribe revision_<fecha>.xlsx con la hoja de metros por día/máquina y, si ya
        se cargó el Excel de Almacén (self.almacen_perforacion_equipo), dos hojas más con la
        validación cruzada contra 'Perforación por Equipo': el detalle largo y una matriz
        Fecha x Máquina con formato condicional (roja = fuera de tolerancia) para corregir
        rápido. Se llama tanto al terminar LoadReport como al terminar LoadAlmacen, para que
        el archivo quede actualizado sin importar en qué orden se cargaron ambos archivos."""
        if not hasattr(self, 'df_final'):
            return
        try:
            fecha_hoy = datetime.now().strftime("%Y-%m-%d")
            nombre_archivo = f"revision_{fecha_hoy}.xlsx"
            with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
                self.df_final.to_excel(writer, sheet_name='Metros por Dia', index=False)
                if hasattr(self, 'almacen_perforacion_equipo'):
                    validacion = self._validar_metros_vs_almacen()
                    validacion.to_excel(writer, sheet_name='Validacion vs Almacen', index=False)

                    matriz = self._build_matriz_diferencias_almacen(validacion)
                    if not matriz.empty:
                        matriz_export = matriz.reset_index()
                        matriz_export.to_excel(writer, sheet_name='Validacion Matriz', index=False)
                        self._aplicar_formato_matriz_diferencias(
                            writer.sheets['Validacion Matriz'], matriz_export
                        )

                    n_diferencias = int((validacion['Estado'] == 'DIFERENCIA').sum())
                    print(f"'{nombre_archivo}' actualizado con la validación vs Almacén "
                          f"({n_diferencias} diferencia(s) encontradas).")
                    if n_diferencias and hasattr(self, 'Mensajes'):
                        texto_actual = self.Mensajes.text()
                        self.Mensajes.setText(
                            (texto_actual + "\n" if texto_actual else "") +
                            f"Validación vs Almacén: {n_diferencias} diferencia(s) en '{nombre_archivo}' "
                            f"(ver hoja 'Validacion Matriz')."
                        )
        except Exception as e:
            traceback.print_exc()
            print(f"No se pudo actualizar '{nombre_archivo}' con la validación vs Almacén: {e}")

    def LoadAlmacen(self):
            self.Mensajes.clear()
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Seleccionar archivo de Almacén (GeoCore)", "", "Excel Files (*.xlsx; *.xlsm)"
            )
            if not file_path:
                return

            try:
                xls = pd.ExcelFile(file_path)
                almacen_data = {}
                etapas_faltantes = []
                for nombre, hoja, columna_valor, _ in ETAPAS_ALMACEN:
                    if hoja not in xls.sheet_names:
                        etapas_faltantes.append(nombre)
                        continue
                    df_etapa = pd.read_excel(xls, sheet_name=hoja)
                    if 'Fecha' not in df_etapa.columns or columna_valor not in df_etapa.columns:
                        etapas_faltantes.append(nombre)
                        continue
                    df_etapa['Fecha'] = pd.to_datetime(df_etapa['Fecha'], dayfirst=True, errors='coerce')
                    df_etapa = df_etapa.dropna(subset=['Fecha'])
                    almacen_data[nombre] = df_etapa

                if not almacen_data:
                    QMessageBox.warning(self, "Aviso",
                        "No se encontraron hojas reconocibles de avance de almacén en el archivo.")
                    return

                self.almacen_data = almacen_data

                mensaje = f"Se cargaron {len(almacen_data)} etapas de almacén."
                if etapas_faltantes:
                    mensaje += f"\nNo se encontraron o no coinciden las columnas de: {', '.join(etapas_faltantes)}."

                # Hoja "Perforación por Equipo": metros por día y máquina reportados de forma
                # independiente en Almacén (formato ancho Fecha+una columna por máquina, no
                # el formato largo de ETAPAS_ALMACEN), para cruzar contra revision_*.xlsx.
                if 'Perforación por Equipo' in xls.sheet_names:
                    df_perf = pd.read_excel(xls, sheet_name='Perforación por Equipo')
                    if 'Fecha' in df_perf.columns:
                        df_perf['Fecha'] = pd.to_datetime(df_perf['Fecha'], errors='coerce')
                        df_perf = df_perf.dropna(subset=['Fecha'])
                        cols_maquina = [c for c in df_perf.columns
                                        if c not in ('Fecha', 'TOTAL DIA', 'Total Dia', 'Total')]
                        self.almacen_perforacion_equipo = df_perf.melt(
                            id_vars='Fecha', value_vars=cols_maquina,
                            var_name='Maquina', value_name='Metros'
                        )
                        self.almacen_perforacion_equipo['Metros'] = pd.to_numeric(
                            self.almacen_perforacion_equipo['Metros'], errors='coerce'
                        ).fillna(0.0)
                        mensaje += "\nSe cargó 'Perforación por Equipo' para validar contra los reportes de turno."
                        self._actualizar_revision_con_validacion_almacen()
                    else:
                        mensaje += "\n'Perforación por Equipo' no tiene columna 'Fecha'; no se pudo usar para validar."

                QMessageBox.information(self, "Éxito", mensaje)
                self.Mensajes.setText("Archivo de Almacén cargado")
                self.PlotAlmacenBtn.setEnabled(True)

            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error", f"Ocurrió un fallo al procesar el archivo de Almacén: {e}")
                self.Mensajes.setText("Ocurrio algun error al cargar el archivo de Almacén")

    def _build_resumen_almacen(self, freq='W-TUE'):
        """Agrega el avance de cada etapa de almacén (self.almacen_data) por semana (miércoles
        a martes, freq='W-TUE') o por mes calendario (freq='M'). Devuelve un DataFrame largo
        con columnas Periodo/Eje_X/Etapa/Unidad/Valor, listo para graficar con un subplot por
        etapa (mismo patrón que _build_resumen_paradas)."""
        filas = []
        for nombre, _, columna_valor, unidad in ETAPAS_ALMACEN:
            df_etapa = self.almacen_data.get(nombre)
            if df_etapa is None or df_etapa.empty:
                continue
            df_etapa = df_etapa.copy()
            df_etapa['Periodo'] = df_etapa['Fecha'].dt.to_period(freq)
            agg = df_etapa.groupby('Periodo')[columna_valor].sum().reset_index()
            agg = agg.rename(columns={columna_valor: 'Valor'})
            agg['Etapa'] = nombre
            agg['Unidad'] = unidad
            if freq == 'W-TUE':
                agg['Eje_X'] = (agg['Periodo'].dt.start_time.dt.strftime('%d%b') + '-' +
                                agg['Periodo'].dt.end_time.dt.strftime('%d%b'))
            else:
                agg['Eje_X'] = agg['Periodo'].dt.strftime('%b-%y')
            filas.append(agg[['Periodo', 'Eje_X', 'Etapa', 'Unidad', 'Valor']])

        if not filas:
            return pd.DataFrame(columns=['Periodo', 'Eje_X', 'Etapa', 'Unidad', 'Valor'])
        return pd.concat(filas, ignore_index=True)

    def _draw_barras_por_categoria(self, resumen, col_categoria, col_valor, freq, figsize,
                                    titulo_figura, titulo_prefijo='', ylabel_func=None,
                                    fmt_valor=lambda v: f'{v:,.0f}'):
        """Base común: un subplot de barras por categoría (etapa de almacén, transición entre
        etapas, etc.), compartiendo el mismo eje de periodo entre todas (ver _periodos_comunes),
        para que la misma columna sea siempre el mismo periodo calendario. Reutilizado por los
        gráficos de avance, WIP (trabajo en proceso) y tiempo de ciclo del módulo de Almacén,
        tanto en su versión interactiva como en la estática para el reporte Word."""
        categorias = list(dict.fromkeys(resumen[col_categoria]))
        num_cat = len(categorias)
        fig, axes = plt.subplots(num_cat, 1, figsize=figsize, sharex=(freq == 'W-TUE'))
        if num_cat == 1:
            axes = [axes]

        meses_recientes = 2 if freq == 'W-TUE' else None
        periodos_comunes = self._periodos_comunes(resumen, col_periodo='Periodo', meses_recientes=meses_recientes)
        colores_cat = self._color_map_por_categoria(categorias, saturacion=0.55, brillo=0.8)

        for ax, categoria in zip(axes, categorias):
            datos_c = periodos_comunes.merge(
                resumen[resumen[col_categoria] == categoria], on=['Periodo', 'Eje_X'], how='left'
            )
            ax.bar(datos_c['Eje_X'], datos_c[col_valor], color=colores_cat[categoria],
                   width=0.6, edgecolor='#555555', lw=0.3)

            for x, v in zip(datos_c['Eje_X'], datos_c[col_valor]):
                if pd.notna(v) and v != 0:
                    ax.text(x, v, fmt_valor(v), ha='center', va='bottom', fontsize=7)

            ax.set_title(f'{titulo_prefijo}{categoria}', fontweight='bold', loc='left', fontsize=10)
            if ylabel_func:
                ax.set_ylabel(ylabel_func(categoria), fontsize=9)
            ax.grid(axis='y', linestyle='--', alpha=0.4)
            if freq == 'W-TUE':
                ax.label_outer()
            else:
                plt.setp(ax.get_xticklabels(), rotation=40, ha='right', fontsize=7)

        if freq == 'W-TUE':
            plt.setp(axes[-1].get_xticklabels(), rotation=40, ha='right', fontsize=7)

        titulo_periodo = 'Semanal (miércoles a martes)' if freq == 'W-TUE' else 'Mensual'
        fig.suptitle(f'{titulo_figura} — Vista {titulo_periodo}', fontsize=14, fontweight='bold', y=1.02)
        fig.tight_layout()
        return fig

    def _draw_almacen_figure(self, resumen, freq, figsize):
        """Avance por etapa de almacén (metros u otra unidad). Ver _draw_barras_por_categoria."""
        unidad_por_etapa = resumen.drop_duplicates('Etapa').set_index('Etapa')['Unidad']
        return self._draw_barras_por_categoria(
            resumen, col_categoria='Etapa', col_valor='Valor', freq=freq, figsize=figsize,
            titulo_figura='Avance de Almacén', titulo_prefijo='Etapa: ',
            ylabel_func=lambda etapa: unidad_por_etapa.get(etapa, '')
        )

    def _build_almacen_wip(self, freq='W-TUE'):
        """Trabajo en proceso (WIP) entre cada par de etapas consecutivas de almacén: cuántos
        metros ya pasaron por la etapa anterior pero todavía no llegan a la siguiente. Se calcula
        con el ACUMULADO de cada etapa hasta cada periodo (no solo lo hecho en ese periodo), para
        que un periodo sin actividad en una etapa no la haga "perder" el WIP ya acumulado."""
        frames = []
        for nombre, _, columna_valor, _ in ETAPAS_ALMACEN:
            df_etapa = self.almacen_data.get(nombre)
            if df_etapa is None or df_etapa.empty:
                continue
            df_etapa = df_etapa.copy()
            df_etapa['Periodo'] = df_etapa['Fecha'].dt.to_period(freq)
            cum = df_etapa.groupby('Periodo')[columna_valor].sum().sort_index().cumsum()
            cum = cum.reset_index()
            cum.columns = ['Periodo', 'Cumulativo']
            cum['Etapa'] = nombre
            frames.append(cum)

        if not frames:
            return pd.DataFrame(columns=['Periodo', 'Eje_X', 'Transicion', 'WIP'])

        cum_todas = pd.concat(frames, ignore_index=True)
        pivot = cum_todas.pivot(index='Periodo', columns='Etapa', values='Cumulativo').sort_index()
        pivot = pivot.ffill().fillna(0)

        nombres_orden = [n for n, _, _, _ in ETAPAS_ALMACEN if n in pivot.columns]
        filas = []
        for actual, siguiente in zip(nombres_orden, nombres_orden[1:]):
            wip = (pivot[actual] - pivot[siguiente]).clip(lower=0)
            df_t = wip.reset_index()
            df_t.columns = ['Periodo', 'WIP']
            df_t['Transicion'] = f"{actual} → {siguiente}"
            filas.append(df_t)

        resultado = pd.concat(filas, ignore_index=True)
        if freq == 'W-TUE':
            resultado['Eje_X'] = (resultado['Periodo'].dt.start_time.dt.strftime('%d%b') + '-' +
                                  resultado['Periodo'].dt.end_time.dt.strftime('%d%b'))
        else:
            resultado['Eje_X'] = resultado['Periodo'].dt.strftime('%b-%y')
        return resultado

    def _draw_almacen_wip_figure(self, resumen_wip, freq, figsize):
        """WIP (backlog) por transición entre etapas de almacén. Ver _draw_barras_por_categoria."""
        return self._draw_barras_por_categoria(
            resumen_wip, col_categoria='Transicion', col_valor='WIP', freq=freq, figsize=figsize,
            titulo_figura='Trabajo en Proceso (WIP) por Transición', ylabel_func=lambda t: 'Metros'
        )

    def _build_almacen_ciclo(self, freq='W-TUE'):
        """Tiempo de ciclo (en días) entre cada par de etapas consecutivas: para cada sondaje
        (DDHID), la diferencia entre la última fecha reportada en la etapa siguiente y la
        última fecha reportada en la etapa anterior, promediada por periodo (agrupada según el
        periodo en que se completó la etapa siguiente). Un tiempo de ciclo creciente en una
        transición específica señala un cuello de botella formándose ahí."""
        ultimas_fechas = {}
        for nombre, _, _, _ in ETAPAS_ALMACEN:
            df_etapa = self.almacen_data.get(nombre)
            if df_etapa is None or df_etapa.empty or 'DDHID' not in df_etapa.columns:
                continue
            ultimas_fechas[nombre] = df_etapa.groupby('DDHID')['Fecha'].max()

        nombres_orden = [n for n, _, _, _ in ETAPAS_ALMACEN if n in ultimas_fechas]
        filas = []
        for anterior, siguiente in zip(nombres_orden, nombres_orden[1:]):
            comunes = ultimas_fechas[anterior].index.intersection(ultimas_fechas[siguiente].index)
            if len(comunes) == 0:
                continue
            dias = (ultimas_fechas[siguiente].loc[comunes] - ultimas_fechas[anterior].loc[comunes]).dt.days
            dias = dias[dias >= 0]
            if dias.empty:
                continue
            periodo = ultimas_fechas[siguiente].loc[dias.index].dt.to_period(freq)
            df_t = pd.DataFrame({'Periodo': periodo.values, 'Dias': dias.values})
            resumen_t = df_t.groupby('Periodo')['Dias'].agg(Dias_Promedio='mean', N_Sondajes='count').reset_index()
            resumen_t['Transicion'] = f"{anterior} → {siguiente}"
            filas.append(resumen_t)

        if not filas:
            return pd.DataFrame(columns=['Periodo', 'Eje_X', 'Transicion', 'Dias_Promedio', 'N_Sondajes'])

        resultado = pd.concat(filas, ignore_index=True)
        if freq == 'W-TUE':
            resultado['Eje_X'] = (resultado['Periodo'].dt.start_time.dt.strftime('%d%b') + '-' +
                                  resultado['Periodo'].dt.end_time.dt.strftime('%d%b'))
        else:
            resultado['Eje_X'] = resultado['Periodo'].dt.strftime('%b-%y')
        return resultado

    def _draw_almacen_ciclo_figure(self, resumen_ciclo, freq, figsize):
        """Días promedio entre etapas consecutivas de almacén. Ver _draw_barras_por_categoria."""
        return self._draw_barras_por_categoria(
            resumen_ciclo, col_categoria='Transicion', col_valor='Dias_Promedio', freq=freq, figsize=figsize,
            titulo_figura='Tiempo de Ciclo entre Etapas', ylabel_func=lambda t: 'Días promedio',
            fmt_valor=lambda v: f'{v:.1f}d'
        )

    def _build_almacen_cfd(self, freq='W-TUE'):
        """Acumulado por etapa y periodo para las 6 etapas de almacén juntas, para armar un
        Diagrama de Flujo Acumulativo (CFD): cada etapa es una línea ascendente, y el área
        sombreada entre dos líneas consecutivas es el WIP entre esas etapas en cada periodo."""
        filas = []
        for nombre, _, columna_valor, _ in ETAPAS_ALMACEN:
            df_etapa = self.almacen_data.get(nombre)
            if df_etapa is None or df_etapa.empty:
                continue
            df_etapa = df_etapa.copy()
            df_etapa['Periodo'] = df_etapa['Fecha'].dt.to_period(freq)
            cum = df_etapa.groupby('Periodo')[columna_valor].sum().sort_index().cumsum().reset_index()
            cum.columns = ['Periodo', 'Cumulativo']
            cum['Etapa'] = nombre
            filas.append(cum)

        if not filas:
            return pd.DataFrame(columns=['Periodo', 'Eje_X', 'Etapa', 'Cumulativo'])

        resultado = pd.concat(filas, ignore_index=True)
        if freq == 'W-TUE':
            resultado['Eje_X'] = (resultado['Periodo'].dt.start_time.dt.strftime('%d%b') + '-' +
                                  resultado['Periodo'].dt.end_time.dt.strftime('%d%b'))
        else:
            resultado['Eje_X'] = resultado['Periodo'].dt.strftime('%b-%y')
        return resultado

    def _draw_almacen_cfd_figure(self, resumen_cfd, freq, figsize, interactivo=False):
        """Dibuja el CFD con las 6 etapas juntas: una línea acumulada por etapa, con el área
        entre cada par de etapas consecutivas sombreada (= WIP visual). La leyenda va FUERA del
        área de trazado (a la derecha) porque con 6 líneas ya no cabe limpiamente adentro.
        Reutilizado por PlotAlmacen (interactivo, con `interactivo=True` para habilitar el hover
        con mplcursors) y _render_almacen_cfd_png (versión estática para el reporte Word)."""
        etapas = [nombre for nombre, _, _, _ in ETAPAS_ALMACEN]
        meses_recientes = 2 if freq == 'W-TUE' else None
        periodos_comunes = self._periodos_comunes(resumen_cfd, col_periodo='Periodo', meses_recientes=meses_recientes)

        pivot = periodos_comunes.merge(resumen_cfd, on=['Periodo', 'Eje_X'], how='left')
        pivot_wide = pivot.pivot_table(index=['Periodo', 'Eje_X'], columns='Etapa', values='Cumulativo')
        etapas_presentes = [e for e in etapas if e in pivot_wide.columns]
        pivot_wide = pivot_wide[etapas_presentes].ffill().fillna(0)
        pivot_wide = pivot_wide.reset_index().sort_values('Periodo')

        fig, ax = plt.subplots(figsize=figsize)
        plt.subplots_adjust(right=0.8)
        colores = self._color_map_por_categoria(etapas_presentes, saturacion=0.55, brillo=0.8)
        x = pivot_wide['Eje_X']

        lineas = []
        for i, etapa in enumerate(etapas_presentes):
            y = pivot_wide[etapa]
            linea, = ax.plot(x, y, color=colores[etapa], lw=2, marker='o', markersize=3, label=etapa, zorder=3)
            lineas.append(linea)
            if i > 0:
                y_prev = pivot_wide[etapas_presentes[i - 1]]
                ax.fill_between(x, y_prev, y, color=colores[etapa], alpha=0.25, zorder=1)

        ax.set_ylabel('Metros acumulados', fontweight='bold')
        ax.grid(axis='y', linestyle='--', alpha=0.4)
        plt.setp(ax.get_xticklabels(), rotation=40, ha='right', fontsize=8)
        ax.legend(loc='upper left', bbox_to_anchor=(1.01, 1), fontsize=8, frameon=False)
        titulo_periodo = 'Semanal (miércoles a martes)' if freq == 'W-TUE' else 'Mensual'
        ax.set_title(f'CFD — Avance de Almacén — Vista {titulo_periodo}', fontweight='bold', fontsize=13, pad=12)
        fig.tight_layout()

        if interactivo:
            cursor = mplcursors.cursor(lineas, hover=True)

            @cursor.connect("add")
            def _on_add(sel):
                idx = int(round(sel.index))
                idx = max(0, min(idx, len(pivot_wide) - 1))
                semana = pivot_wide['Eje_X'].iloc[idx]
                lineas_texto = [f"Semana: {semana}"]
                lineas_texto += [f"{etapa}: {pivot_wide[etapa].iloc[idx]:,.0f} m" for etapa in etapas_presentes]
                sel.annotation.set_text("\n".join(lineas_texto))
                sel.annotation.get_bbox_patch().set(fc="white", alpha=0.95, ec="gray")
            # Mantenemos una referencia para que el Garbage Collector no lo destruya mientras
            # la ventana interactiva sigue abierta.
            self._cfd_cursor = cursor

        return fig

    def _render_almacen_cfd_png(self, resumen_cfd, freq):
        fig = self._draw_almacen_cfd_figure(resumen_cfd, freq, figsize=(11, 5.5))
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def _build_almacen_resumen_total(self):
        """Total acumulado a la fecha de cada etapa de almacén (sin desglose por periodo), para
        el resumen horizontal tipo flujograma que se muestra al inicio de la sección de Almacén
        en el reporte Word."""
        filas = []
        for nombre, _, columna_valor, unidad in ETAPAS_ALMACEN:
            df_etapa = self.almacen_data.get(nombre)
            total = df_etapa[columna_valor].sum() if df_etapa is not None and not df_etapa.empty else 0.0
            filas.append({'Etapa': nombre, 'Total': total, 'Unidad': unidad})
        return pd.DataFrame(filas)

    def _draw_almacen_resumen_figure(self, resumen_total):
        """Flujograma horizontal: una caja por etapa con su metraje total acumulado, y flechas
        entre etapas consecutivas anotadas con el WIP (la diferencia) cuando ambas comparten
        unidad — así de un vistazo se ve dónde se acumula más metraje sin procesar. Reutilizado
        por PlotAlmacen (interactivo) y _render_almacen_resumen_png (versión estática para el
        reporte Word)."""
        etapas = resumen_total['Etapa'].tolist()
        totales = resumen_total['Total'].tolist()
        unidades = resumen_total['Unidad'].tolist()
        n = len(etapas)

        fig, ax = plt.subplots(figsize=(13, 3.2))
        ax.set_xlim(0, n)
        ax.set_ylim(0, 1)
        ax.axis('off')

        colores = self._color_map_por_categoria(etapas, saturacion=0.5, brillo=0.85)
        ancho_caja = 0.72
        for i, (etapa, total, unidad) in enumerate(zip(etapas, totales, unidades)):
            cx = i + 0.5
            caja = patches.FancyBboxPatch((cx - ancho_caja / 2, 0.32), ancho_caja, 0.38,
                                           boxstyle="round,pad=0.02", facecolor=colores[etapa],
                                           edgecolor='#555555', linewidth=1, zorder=2)
            ax.add_patch(caja)
            ax.text(cx, 0.58, etapa, ha='center', va='center', fontsize=8.5, fontweight='bold', zorder=3)
            ax.text(cx, 0.42, f'{total:,.0f} {unidad}', ha='center', va='center', fontsize=9.5, zorder=3)

            if i < n - 1:
                x_ini = cx + ancho_caja / 2 + 0.02
                x_fin = i + 1 + 0.5 - ancho_caja / 2 - 0.02
                ax.annotate('', xy=(x_fin, 0.5), xytext=(x_ini, 0.5),
                            arrowprops=dict(arrowstyle='-|>', color='#555555', lw=1.5), zorder=1)
                if unidad == unidades[i + 1]:
                    wip = total - totales[i + 1]
                    color_wip = '#C62828' if wip > 0 else '#2E7D32'
                    ax.text((x_ini + x_fin) / 2, 0.78, f'WIP\n{wip:,.0f} m',
                            ha='center', va='center', fontsize=7.5, color=color_wip)

        ax.set_title('Resumen de Almacén — Metraje Acumulado por Etapa', fontweight='bold', fontsize=13, pad=10)
        fig.tight_layout()
        return fig

    def _render_almacen_resumen_png(self, resumen_total):
        fig = self._draw_almacen_resumen_figure(resumen_total)
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def PlotAlmacen(self):
            self.Mensajes.clear()
            try:
                # En pantalla solo mostramos el resumen y el CFD (interactivo, con hover). Los
                # gráficos de barras (avance, WIP, tiempo de ciclo) quedan solo en el reporte
                # Word para no abrir demasiadas ventanas en cascada.
                resumen_total = self._build_almacen_resumen_total()
                self._draw_almacen_resumen_figure(resumen_total)
                plt.show()

                cfd_semanal = self._build_almacen_cfd(freq='W-TUE')
                if not cfd_semanal.empty:
                    self._draw_almacen_cfd_figure(cfd_semanal, freq='W-TUE', figsize=(15, 7), interactivo=True)
                    plt.show()

                cfd_mensual = self._build_almacen_cfd(freq='M')
                if not cfd_mensual.empty:
                    self._draw_almacen_cfd_figure(cfd_mensual, freq='M', figsize=(15, 7), interactivo=True)
                    plt.show()

                self.Mensajes.setText("Resumen y CFD de Almacén Ploteados")

            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "Error al Graficar", f"Error crítico: {str(e)}")
                self.Mensajes.setText("Ocurrio algun error")

    def _render_almacen_png(self, resumen, freq):
        num_etapas = resumen['Etapa'].nunique()
        fig = self._draw_almacen_figure(resumen, freq, figsize=self._figsize_estatico(num_etapas))

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def _render_almacen_wip_png(self, resumen_wip, freq):
        num_trans = resumen_wip['Transicion'].nunique()
        fig = self._draw_almacen_wip_figure(resumen_wip, freq, figsize=self._figsize_estatico(num_trans))

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def _render_almacen_ciclo_png(self, resumen_ciclo, freq):
        num_trans = resumen_ciclo['Transicion'].nunique()
        fig = self._draw_almacen_ciclo_figure(resumen_ciclo, freq, figsize=self._figsize_estatico(num_trans))

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return buf

    def toggle_text(self):
        if self.aboutlabel.isHidden():
            self.aboutlabel.setText("Desarrollado por: Felix Recines Trinidad")
            self.aboutlabel.show()
        else:
            self.aboutlabel.clear()
            self.aboutlabel.hide()
if __name__ == "__main__":

    app =  QtWidgets.QApplication(sys.argv)
    sys.excepthook = manejar_excepcion_global
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
    
