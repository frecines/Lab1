# -*- coding: utf-8 -*-
"""
Modulo QAQC - interfaz grafica.

Created on Sat May 18 09:43:03 2024
@author: frecines

Estructura:
    qaqc.py         interfaz (este archivo)
    qaqcbs.ui       diseno de la ventana (Qt Designer)
    qaqc_calc.py    calculos y reportes (sin Qt)
    qaqc_plots.py   graficos
    qaqc_config.py  parametros: limites, estandares, umbrales
"""
import functools
import os
import sys
import traceback
from datetime import datetime

# En el .exe sin consola (--windowed) sys.stdout/stderr son None y cualquier
# print o traceback haria fallar la aplicacion.
for _flujo in ("stdout", "stderr"):
    if getattr(sys, _flujo) is None:
        setattr(sys, _flujo, open(os.devnull, "w"))

from PyQt5 import QtWidgets  # noqa: E402
from PyQt5.QtCore import (QObject, QRegularExpression, QRunnable, Qt,  # noqa: E402
                          QThreadPool, pyqtSignal)
from PyQt5.QtGui import QBrush, QColor, QFont, QRegularExpressionValidator  # noqa: E402
from PyQt5.QtWidgets import (QApplication, QCompleter, QFileDialog,  # noqa: E402
                             QMainWindow, QMessageBox, QTableWidgetItem)

import pandas as pd  # noqa: E402

import qaqc_calc as calc  # noqa: E402
import qaqc_config as cfg  # noqa: E402
import qaqc_estandares as estandares  # noqa: E402

# matplotlib y los graficos se importan recien al abrir el primer grafico
# (_modulos_graficos): asi la ventana principal abre ~0.8 s antes.


def _ruta_recurso(nombre):
    """Ruta junto a este archivo (funciona aunque se ejecute desde otra carpeta
    o empaquetado con PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nombre)


try:  # interfaz compilada con pyuic5 (la que usa el .exe)
    from ui_qaqcbs import Ui_ModuloQAQC as Ui_MainWindow
except ImportError:  # en desarrollo, si no se compilo, se lee el .ui
    from PyQt5 import uic
    Ui_MainWindow, _ = uic.loadUiType(_ruta_recurso("qaqcbs.ui"))

FILTRO_EXCEL = "Archivos Excel (*.xlsx *.xlsm *.xls)"
COLOR_CLASE = {"ACEPTABLE": "green", "CUESTIONABLE": "orange",
               "INACEPTABLE": "red", "No Certificado": "black", "Sin datos": "gray"}
GRIS = QColor("lightgray")
ARCHIVO_ERRORES = os.path.join(os.path.expanduser("~"), "QAQC_errores.log")


@functools.lru_cache(maxsize=None)
def _modulos_graficos():
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_qt5agg import (FigureCanvasQTAgg,
                                                    NavigationToolbar2QT)
    import qaqc_plots
    return Figure, FigureCanvasQTAgg, NavigationToolbar2QT, qaqc_plots


class _GraficosPerezosos:
    """`plots.funcion` importa qaqc_plots (y matplotlib) la primera vez."""

    def __getattr__(self, nombre):
        return getattr(_modulos_graficos()[3], nombre)


plots = _GraficosPerezosos()


def accion(titulo_error):
    """Ejecuta el slot con cursor de espera y muestra cualquier error en un
    mensaje y en el registro, en vez de cerrar la aplicacion."""
    def decorador(funcion):
        @functools.wraps(funcion)
        def envoltura(self, *args):
            try:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                try:
                    return funcion(self, *args)
                finally:
                    QApplication.restoreOverrideCursor()
            except Exception as e:  # noqa: BLE001 - se informa al usuario
                self.mostrar_error(titulo_error, e, traceback.format_exc())
        return envoltura
    return decorador


class _Senales(QObject):
    terminado = pyqtSignal(object)
    fallo = pyqtSignal(object, str)


class Tarea(QRunnable):
    """Ejecuta una funcion en segundo plano (lectura de Excel, procesamiento)
    para que la ventana no quede 'No responde'."""

    def __init__(self, funcion):
        super().__init__()
        self.funcion = funcion
        self.senales = _Senales()

    def run(self):
        try:
            resultado = self.funcion()
        except Exception as e:  # noqa: BLE001 - se reenvia al hilo principal
            self.senales.fallo.emit(e, traceback.format_exc())
        else:
            self.senales.terminado.emit(resultado)


class VentanaGraficos(QtWidgets.QWidget):
    """Una ventana con una pestana por grafico (antes se abrian 6 ventanas
    sueltas por boton)."""

    def __init__(self, titulo, padre):
        super().__init__(padre, Qt.Window)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setWindowTitle(titulo)
        self.resize(1000, 760)
        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setUsesScrollButtons(True)
        QtWidgets.QVBoxLayout(self).addWidget(self.tabs)

    def agregar(self, nombre, fig):
        NavigationToolbar = _modulos_graficos()[2]
        pagina = QtWidgets.QWidget()
        capa = QtWidgets.QVBoxLayout(pagina)
        capa.addWidget(NavigationToolbar(fig.canvas, pagina))
        capa.addWidget(fig.canvas)
        self.tabs.addTab(pagina, nombre)


class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.bd_ini = self.bd_sec = self.bd_cert = self.bd_rma = None
        self.res = self.reporte_rma = None
        self.carpeta_salida = None

        validador = QRegularExpressionValidator(QRegularExpression(r"\d*"), self)
        self.Vmin.setValidator(validador)
        self.Vmax.setValidator(validador)
        for tabla in self.findChildren(QtWidgets.QTableWidget):
            tabla.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            tabla.setAlternatingRowColors(True)
            tabla.horizontalHeader().setStyleSheet(
                "QHeaderView::section { background-color: lightgray; font-weight: bold; }")

        self.botones_qaqc = [self.GENR, self.GBF, self.PGBF, self.GBG, self.PGBG,
                             self.GDC, self.QQDC, self.RDDC, self.GRG, self.QQRG,
                             self.RDRG, self.GDP, self.QQDP, self.RDDP, self.GSTD,
                             self.CHNH, self.GRI]
        self.botones_rma = [self.GRMA, self.QQPRMA, self.RDRMA, self.CUSSL, self.REPRMA]
        for boton in [self.BDQaqc, self.PROC, self.ProcRMA,
                      *self.botones_qaqc, *self.botones_rma]:
            boton.setEnabled(False)

        conexiones = {
            self.BDLeyes: self.importar_leyes, self.BDQaqc: self.importar_reporte,
            self.PROC: self.procesar, self.GENR: self.exportar_reporte,
            self.OutDir: self.elegir_carpeta,
            self.GBF: lambda: self.graficos_blanco("BF"),
            self.GBG: lambda: self.graficos_blanco("BG"),
            self.PGBF: lambda: self.graficos_precedente("BF"),
            self.PGBG: lambda: self.graficos_precedente("BG"),
            self.GDC: lambda: self.graficos_duplicado("DC"),
            self.GRG: lambda: self.graficos_duplicado("RG"),
            self.GDP: lambda: self.graficos_duplicado("DP"),
            self.QQDC: lambda: self.graficos_qq_duplicado("DC"),
            self.QQRG: lambda: self.graficos_qq_duplicado("RG"),
            self.QQDP: lambda: self.graficos_qq_duplicado("DP"),
            self.RDDC: lambda: self.graficos_rd_duplicado("DC"),
            self.RDRG: lambda: self.graficos_rd_duplicado("RG"),
            self.RDDP: lambda: self.graficos_rd_duplicado("DP"),
            self.GSTD: lambda: self.graficos_estandar(self.cmbSTD.currentData()),
            self.ParamSTD: self.parametros_estandares,
            self.CHNH: self.graficos_historico,
            self.ImpRMA: self.importar_externo, self.ProcRMA: self.procesar_externo,
            self.GRMA: self.graficos_rma, self.QQPRMA: self.graficos_qq_rma,
            self.RDRMA: self.graficos_rd_rma, self.CUSSL: self.graficos_cuss,
            self.REPRMA: self.exportar_reporte_externo, self.GRI: self.ratio_insercion,
        }
        for boton, slot in conexiones.items():
            # lambda: descarta el argumento 'checked' que envia clicked
            boton.clicked.connect(lambda _=False, s=slot: s())
        self.actionAcerca.triggered.connect(self.acerca_de)
        self.actionEstandares.triggered.connect(lambda _=False: self.parametros_estandares())
        self.param = estandares.Parametros()

        # Malla (~800 valores) y Scheme (~1100): combos con busqueda por texto
        for combo in (self.VMalla, self.VLote):
            combo.setEditable(True)
            combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
            combo.completer().setFilterMode(Qt.MatchContains)
            combo.completer().setCompletionMode(QCompleter.PopupCompletion)

        self._tareas = set()
        self.progreso = QtWidgets.QProgressBar()
        self.progreso.setRange(0, 0)          # barra "ocupado"
        self.progreso.setMaximumWidth(180)
        self.progreso.hide()
        self.statusbar.addPermanentWidget(self.progreso)
        self._mostrar_carpeta()
        if self.param.aviso:
            self.log(self.param.aviso, error=True)
        self.log(f"Estandares cargados ({', '.join(self.param.estandares)}) - "
                 f"actualizados: {self.param.actualizado}.")
        self.log("Seleccione la BD de Leyes para comenzar.")

    # ------------------------------------------------------------ utilidades
    def log(self, mensaje, error=False):
        hora = datetime.now().strftime("%H:%M:%S")
        color = "red" if error else "#1a4fa0"
        self.Log.appendHtml(f'<span style="color:gray">{hora}</span> '
                            f'<span style="color:{color}">{mensaje}</span>')
        self.statusbar.showMessage(mensaje, 8000)

    def mostrar_error(self, titulo, error, detalle=""):
        mensaje = str(error)
        if isinstance(error, PermissionError):
            mensaje += "\n\nRevisar que el archivo no este abierto en Excel."
        self.log(f"{titulo}: {error}", error=True)
        try:  # el detalle tecnico va a un archivo (en el .exe no hay consola)
            with open(ARCHIVO_ERRORES, "a", encoding="utf-8") as f:
                f.write(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] {titulo}\n{detalle}")
        except OSError:
            pass
        QMessageBox.critical(self, titulo, mensaje)

    def en_segundo_plano(self, mensaje, funcion, al_terminar, titulo_error):
        """Ejecuta `funcion` en otro hilo y luego `al_terminar(resultado)` en
        el hilo de la interfaz. Mientras tanto la ventana queda bloqueada."""
        tarea = Tarea(funcion)
        self._tareas.add(tarea)

        def fin():
            self._tareas.discard(tarea)
            self.tabWidget.setEnabled(True)
            self.progreso.hide()
            QApplication.restoreOverrideCursor()

        def ok(resultado):
            fin()
            try:
                al_terminar(resultado)
            except Exception as e:  # noqa: BLE001
                self.mostrar_error(titulo_error, e, traceback.format_exc())

        def error(e, detalle):
            fin()
            self.mostrar_error(titulo_error, e, detalle)

        tarea.senales.terminado.connect(ok)
        tarea.senales.fallo.connect(error)
        self.tabWidget.setEnabled(False)
        self.progreso.show()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.log(mensaje)
        QThreadPool.globalInstance().start(tarea)

    def parametros_estandares(self):
        from qaqc_dialogo_estandares import DialogoEstandares
        if DialogoEstandares(self.param, self).exec_():
            self.log(f"Parametros de estandares guardados ({self.param.actualizado}).")
            if self.res is not None:
                self.log("Recalculando con los nuevos parametros...")
                self.procesar()

    def acerca_de(self):
        QMessageBox.about(self, "Acerca de", "Modulo QAQC\n\n"
                          "Desarrollado por: Felix Recines Trinidad")

    def _carpeta(self):
        carpeta = self.carpeta_salida or os.getcwd()
        os.makedirs(carpeta, exist_ok=True)
        return carpeta

    def _mostrar_carpeta(self):
        self.lblOutDir.setText(self.carpeta_salida or os.getcwd())

    def elegir_carpeta(self):
        carpeta = QFileDialog.getExistingDirectory(self, "Carpeta de salida", self._carpeta())
        if carpeta:
            self.carpeta_salida = carpeta
            self._mostrar_carpeta()

    def _abrir_excel(self, titulo):
        ruta, _ = QFileDialog.getOpenFileName(self, titulo, "", FILTRO_EXCEL)
        return ruta

    def _rango(self):
        if not self.Vmin.text() or not self.Vmax.text():
            raise ValueError("Ingrese el SAMPLE_ID minimo y maximo.")
        vmin, vmax = int(self.Vmin.text()), int(self.Vmax.text())
        if vmax <= vmin:
            raise ValueError("Verifique que el valor maximo sea mayor al minimo.")
        return vmin, vmax

    @staticmethod
    def _filtro(combo, nombre):
        texto = combo.currentText().strip()
        if not texto or texto == combo.itemText(0):
            return None
        indice = combo.findText(texto, Qt.MatchFixedString)
        if indice < 0:
            raise ValueError(f"{nombre} '{texto}' no existe en el Reporte de Muestras QAQC.")
        return combo.itemText(indice)

    # ---------------------------------------------------------------- tablas
    @staticmethod
    def llenar_tabla(tabla, df, col_estandar=None, gris_primera=True, colores=None):
        tabla.clear()
        tabla.setRowCount(df.shape[0])
        tabla.setColumnCount(df.shape[1])
        tabla.setHorizontalHeaderLabels([str(c) for c in df.columns])
        clases = [j for j, c in enumerate(df.columns) if "Clase" in str(c)]
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                valor = df.iat[i, j]
                texto = "" if not isinstance(valor, str) and pd.isna(valor) else str(valor)
                item = QTableWidgetItem(texto)
                if j in clases and texto in COLOR_CLASE:
                    item.setForeground(QColor(COLOR_CLASE[texto]))
                if j == col_estandar:
                    item.setBackground(QColor((colores or {}).get(texto, "lightgray")))
                elif j == 0 and gris_primera:
                    item.setBackground(GRIS)
                tabla.setItem(i, j, item)
        tabla.resizeColumnsToContents()
        tabla.resizeRowsToContents()

    def _llenar_tablas_qaqc(self):
        r = self.res.reportes
        self.llenar_tabla(self.DFBLF, r["Blanco Fino"])
        self.llenar_tabla(self.DFBLG, r["Blanco Grueso"])
        self.llenar_tabla(self.DFDUP, r["Duplicado"])
        colores = estandares.colores(self.param.estandares)
        self.llenar_tabla(self.DFSTD, r["Estandares"], col_estandar=0, colores=colores)
        self.llenar_tabla(self.DFSTDB, r["Estandares por Analito"], col_estandar=1,
                          colores=colores)
        self.llenar_tabla(self.E_STD, calc.errores_estandares(self.res), col_estandar=0,
                          colores=colores)
        self.cmbSTD.clear()
        for control, conf in self.res.conf_std.items():
            self.cmbSTD.addItem(f"{conf['codigo']} - {conf['titulo']} ({control})", control)
        self.GSTD.setEnabled(bool(self.res.conf_std))
        errores = {self.ErrorBlancoF: "ErrorBlancoFino", self.ErrorBlancoG: "ErrorBlancoGrueso",
                   self.E_DUPC: "ErrorDuplicadoCampo", self.E_RG: "ErrorRechazoGrueso",
                   self.E_PULP: "ErrorDuplicadoPulpa"}
        for tabla, hoja in errores.items():
            self.llenar_tabla(tabla, r[hoja], gris_primera=False)

    # ------------------------------------------------------- importacion
    @staticmethod
    def _leer(lector, ruta, hoja):
        try:
            return lector(ruta)
        except ValueError as e:
            if "not found" in str(e) or "Worksheet" in str(e):
                raise ValueError(f"El excel seleccionado debe tener la hoja '{hoja}'.") from e
            raise

    def importar_leyes(self):
        ruta = self._abrir_excel("Seleccionar archivo BD Leyes")
        if ruta:
            self.en_segundo_plano(
                f"Leyendo {os.path.basename(ruta)}...",
                lambda: self._leer(calc.leer_bd_leyes, ruta, "General"),
                lambda bd: self._leyes_cargadas(ruta, bd), "Error al cargar el archivo")

    def _leyes_cargadas(self, ruta, bd):
        ids = pd.to_numeric(bd.get("SAMPLE_ID"), errors="coerce").dropna()
        if ids.empty:
            raise ValueError("La hoja 'General' no tiene la columna SAMPLE_ID con valores numericos.")
        self.bd_ini = bd
        self.lblLeyes.setText(os.path.basename(ruta))
        self.Vmin.setText(str(int(ids.min())))
        self.Vmax.setText(str(int(ids.max())))
        if self.carpeta_salida is None:
            self.carpeta_salida = os.path.join(os.path.dirname(ruta), "Resultados_QAQC")
            self._mostrar_carpeta()
        self.BDQaqc.setEnabled(True)
        self.PROC.setEnabled(self.bd_sec is not None)
        self._actualizar_filtros()
        self.log(f"BD con Leyes importada ({len(bd)} filas). Cargue el Reporte de Muestras QAQC.")

    def importar_reporte(self):
        ruta = self._abrir_excel("Seleccionar archivo Reporte de Muestras QAQC")
        if ruta:
            self.en_segundo_plano(
                f"Leyendo {os.path.basename(ruta)} (puede tardar unos segundos)...",
                lambda: self._leer(calc.leer_reporte_qaqc, ruta, "LEYES SGS"),
                lambda bd: self._reporte_cargado(ruta, bd), "Error al cargar el archivo")

    def _reporte_cargado(self, ruta, bd):
        self.bd_sec = bd
        self.lblQaqc.setText(os.path.basename(ruta))
        self._actualizar_filtros()
        self.PROC.setEnabled(self.bd_ini is not None)
        self.log(f"Reporte de Muestras QAQC importado ({len(bd)} filas). Revise el rango y "
                 "los filtros y presione 'Procesar Datos'.")

    def _actualizar_filtros(self):
        """Mallas y reportes que tienen muestras en la BD de Leyes cargada."""
        if self.bd_ini is None or self.bd_sec is None:
            return
        ids = pd.to_numeric(self.bd_ini["SAMPLE_ID"], errors="coerce")
        sec = self.bd_sec[pd.to_numeric(self.bd_sec["SAMPLE_ID"], errors="coerce").isin(ids)]
        for combo, columna, todos in ((self.VMalla, "Malla", "(Todas)"),
                                      (self.VLote, "Scheme", "(Todos)")):
            combo.clear()
            combo.addItems([todos, *calc.opciones_filtro(sec, columna)])

    @accion("Error en el procesamiento")
    def procesar(self):
        vmin, vmax = self._rango()
        malla = self._filtro(self.VMalla, "La malla")
        reporte = self._filtro(self.VLote, "El reporte")
        bd_ini, bd_sec = self.bd_ini, self.bd_sec
        parametros_std = dict(self.param.estandares)

        def calcular():
            bd_total, avisos = calc.unir_bases(bd_ini, bd_sec, vmin, vmax, malla, reporte)
            return calc.ResultadosQAQC(bd_total, parametros_std), avisos

        filtros = [f"rango {vmin} - {vmax}"]
        filtros += [f"malla {malla}"] if malla else []
        filtros += [f"reporte {reporte}"] if reporte else []
        self.en_segundo_plano("Procesando...", calcular,
                              lambda r: self._procesado(*r, ", ".join(filtros)),
                              "Error en el procesamiento")

    def _procesado(self, res, avisos, filtros):
        for aviso in avisos:
            self.log(f"Aviso: {aviso}", error=True)
        self.res = res
        self._llenar_tablas_qaqc()
        for boton in self.botones_qaqc:
            boton.setEnabled(True)
        # La union cambio: hay que volver a procesar el control externo
        self.bd_rma = None
        for boton in self.botones_rma:
            boton.setEnabled(False)
        self.ProcRMA.setEnabled(self.bd_cert is not None)

        conteo = self.res.conteo_controles()
        resumen = ", ".join(f"{k}: {v}" for k, v in conteo.items())
        self.log(f"Procesado ({filtros}): {len(res.bd_total)} muestras. "
                 f"Controles -> {resumen or 'ninguno'}.")
        sin_parametros = [c for c in conteo.index
                          if c not in calc.controles_conocidos(self.param.estandares)]
        if sin_parametros:
            self.log(f"Controles sin parametros en qaqc_config.py (no se evaluan): "
                     f"{', '.join(map(str, sin_parametros))}")

    @accion("No se creo el reporte")
    def exportar_reporte(self):
        ruta = os.path.join(self._carpeta(), "Reporte_QAQC.xlsx")
        calc.exportar_excel(self.res.reportes, ruta)
        self.log(f"Reporte exportado: {ruta}")

    # -------------------------------------------------------------- graficos
    def _mostrar(self, titulo, graficos):
        """graficos: lista de (pestana, archivo_png, figsize, dibujar(fig))."""
        Figure, FigureCanvas = _modulos_graficos()[:2]
        guardar = self.chkGuardarPNG.isChecked()
        carpeta = self._carpeta() if guardar else None
        ventana = VentanaGraficos(titulo, self)
        for pestana, archivo, tamano, dibujar in graficos:
            fig = Figure(figsize=tamano)
            FigureCanvas(fig)  # el canvas Qt debe existir antes de crear los tooltips
            dibujar(fig)
            if guardar:
                fig.savefig(os.path.join(carpeta, f"{archivo}.png"), dpi=300,
                            bbox_inches="tight")
            ventana.agregar(pestana, fig)
        ventana.show()
        self.log(f"{titulo}: {len(graficos)} graficos"
                 + (f" guardados en {carpeta}" if guardar else ""))

    def _lista_blanco(self, clave, meses=False, prefijo=""):
        conf, df = cfg.BLANCOS[clave], self.res.blancos[clave]
        sufijo = " Hist" if meses else ""
        return [(f"{prefijo}{cfg.ETIQUETA[e]}", f"{conf['nombre']} {e}{sufijo}", (10, 6),
                 functools.partial(plots.grafico_blanco, df=df, elemento=e,
                                   titulo=f"{conf['nombre']} {e}",
                                   tasa_error=calc.tasa(df, f"{e}_ERR"), meses=meses))
                for e in cfg.ELEMENTOS]

    def _lista_estandar(self, clave, meses=False, prefijo=""):
        conf, df = self.res.conf_std[clave], self.res.estandares[clave]
        sufijo = " Hist" if meses else ""
        return [(f"{prefijo}{cfg.ETIQUETA[e]}", f"{conf['titulo']} {cfg.ETIQUETA[e]}{sufijo}",
                 (10, 6), functools.partial(plots.grafico_estandar, df=df, elemento=e,
                                            conf=conf, meses=meses))
                for e in ["CU", "AG", "AS", "MO", "ZN"]]

    @accion("Error al graficar")
    def graficos_blanco(self, clave):
        self._mostrar(cfg.BLANCOS[clave]["nombre"], self._lista_blanco(clave))

    @accion("Revisar Rango")
    def graficos_precedente(self, clave):
        nombre, df = cfg.BLANCOS[clave]["nombre"], self.res.blancos[clave]
        graficos = []
        for e in cfg.ELEMENTOS:
            titulo = f"Muestra Precedente vs {nombre} {cfg.ETIQUETA[e]}"
            graficos.append((cfg.ETIQUETA[e], titulo, (10, 6),
                             functools.partial(plots.grafico_precedente, df=df,
                                               elemento=e, titulo=titulo)))
        self._mostrar(f"Muestras Precedentes vs {nombre}", graficos)

    @accion("Error al graficar")
    def graficos_duplicado(self, clave):
        conf, df = cfg.DUPLICADOS[clave], self.res.duplicados[clave]
        self._mostrar(conf["nombre"], [
            (cfg.ETIQUETA[e], f"{conf['titulo']} {e}", (8, 8),
             functools.partial(plots.grafico_duplicado, df=df, elemento=e, conf=conf,
                               tasa_error=calc.tasa(df, f"{e}_ERR")))
            for e in cfg.ELEMENTOS])

    @accion("Error al graficar")
    def graficos_qq_duplicado(self, clave):
        df = self.res.duplicados[clave]
        graficos = []
        for e in cfg.ELEMENTOS:
            et = cfg.ETIQUETA[e]
            graficos.append((et, f"Q-Q Plot {e} {clave}", (8, 8), functools.partial(
                plots.grafico_qq, x=df[f"{e}_O"], y=df[f"{e}_D"],
                titulo=f"Q-Q Plot {clave} {e}", xlabel="Original", ylabel="Duplicado",
                nombre_x=f"{et} Orig", nombre_y=f"{et} Dup")))
        self._mostrar(f"Q-Q Plot {cfg.DUPLICADOS[clave]['nombre']}", graficos)

    @accion("Error al graficar")
    def graficos_rd_duplicado(self, clave):
        df = self.res.duplicados[clave]
        graficos = []
        for e in cfg.ELEMENTOS:
            et = cfg.ETIQUETA[e]
            titulo = f"{clave} {et} RD-Plot"
            graficos.append((et, titulo, (10, 8), functools.partial(
                plots.grafico_rd, df=df, col_prom=f"{e}_PROM", col_rd=f"{e}_RD",
                titulo=titulo, etiqueta=et, id_col="SAMPLE_ORIG")))
        self._mostrar(f"RD-Plot {cfg.DUPLICADOS[clave]['nombre']}", graficos)

    @accion("Error al graficar")
    def graficos_estandar(self, clave):
        if not clave:
            raise ValueError("No hay estandares con muestras en el rango procesado.")
        self._mostrar(self.res.conf_std[clave]["titulo"], self._lista_estandar(clave))

    @accion("Error al graficar")
    def graficos_historico(self):
        graficos = []
        for clave in cfg.BLANCOS:
            graficos += self._lista_blanco(clave, meses=True, prefijo=f"{clave} ")
        for clave, conf in self.res.conf_std.items():
            graficos += self._lista_estandar(clave, meses=True, prefijo=f"{conf['codigo']} ")
        self._mostrar("Historico por Mes - Blancos y Estandares", graficos)

    # ------------------------------------------------------ control externo
    def importar_externo(self):
        ruta = self._abrir_excel("Seleccionar reporte del laboratorio externo")
        if ruta:
            self.en_segundo_plano(f"Leyendo {os.path.basename(ruta)}...",
                                  lambda: calc.leer_reporte_externo(ruta),
                                  lambda bd: self._externo_cargado(ruta, bd),
                                  "Error al cargar el archivo")

    def _externo_cargado(self, ruta, bd):
        self.bd_cert = bd
        self.lblRMA.setText(os.path.basename(ruta))
        self.ProcRMA.setEnabled(self.res is not None)
        self.log("Reporte de muestras externas importado."
                 + ("" if self.res else " Procese primero la pestana de Importacion."))

    @accion("Error en el proceso")
    def procesar_externo(self):
        if self.res is None:
            raise ValueError("Primero procese los datos en la pestana de Importacion.")
        vmin, vmax = self._rango()
        self.bd_rma = calc.procesar_externo(self.bd_cert, self.res.bd_total, vmin, vmax)
        reporte = calc.reporte_externo(self.bd_rma)
        self.llenar_tabla(self.DFRMA, reporte)
        self.reporte_rma = reporte
        for boton in self.botones_rma:
            boton.setEnabled(True)
        self.log(f"Control externo procesado: {len(self.bd_rma)} muestras.")

    def _pares_externos(self):
        """(etiqueta, unidad, pares SGS/Certimin) para cada analito."""
        for col_sgs, (_, etiqueta, unidad) in cfg.EXTERNO.items():
            yield etiqueta, unidad, calc.pares_externo(self.bd_rma, col_sgs)

    @accion("Error inesperado")
    def graficos_rma(self):
        graficos = []
        for et, unidad, p in self._pares_externos():
            dibujar = functools.partial(plots.grafico_rma, pares=p, etiqueta=et, unidad=unidad,
                                        sesgo=calc.sesgo_rma(p["X"], p["Y"]))
            graficos.append((et, f"{et.upper()} RMA PLOT", (9, 8), dibujar))
        self._mostrar("RMA - Laboratorio Externo", graficos)

    @accion("Error inesperado")
    def graficos_qq_rma(self):
        graficos = []
        for et, _, p in self._pares_externos():
            dibujar = functools.partial(
                plots.grafico_qq, x=p["X"], y=p["Y"], titulo=f"{et.upper()}% QQ-Plot",
                xlabel="SGS Toromocho", ylabel="Certimin", nombre_x=f"{et} SGS",
                nombre_y=f"{et} Cert", etiqueta=f"{et.upper()} Q-Q Plot")
            graficos.append((et, f"Q-Q Plot {et.upper()} RMA", (8, 8), dibujar))
        self._mostrar("Q-Q Plot Lab Ext vs Int", graficos)

    @accion("Error inesperado")
    def graficos_rd_rma(self):
        graficos = []
        for et, _, p in self._pares_externos():
            titulo = f"RMA {et} RD-Plot"
            dibujar = functools.partial(plots.grafico_rd, df=p, col_prom="PROM",
                                        col_rd="RD", titulo=titulo, etiqueta=et)
            graficos.append((et, titulo, (10, 8), dibujar))
        self._mostrar("RD-Plot Lab Ext vs Int", graficos)

    @accion("Error inesperado")
    def graficos_cuss(self):
        self._mostrar("CuSS Certimin vs SGS", [
            ("CuCi vs CuSS", "LinePlot CUSS CUCI", (12, 8),
             functools.partial(plots.grafico_cuss, bd_rma=self.bd_rma, incluir_cuci=True)),
            ("CuSS Cert vs SGS", "LinePlot CUSS SGS vs CERT", (12, 8),
             functools.partial(plots.grafico_cuss, bd_rma=self.bd_rma, incluir_cuci=False))])

    @accion("No se creo el reporte")
    def exportar_reporte_externo(self):
        ruta = os.path.join(self._carpeta(), "Reporte_QAQC_Ext.xlsx")
        calc.exportar_excel({"RMA SGS-CERT": self.reporte_rma}, ruta)
        self.log(f"Reporte de control externo exportado: {ruta}")

    # --------------------------------------------------- ratio de insercion
    @accion("Inserte un Valor")
    def ratio_insercion(self):
        anio, mes = self.VAno.text().strip(), self.VMes.text().strip()
        if not anio or not mes:
            raise ValueError("Ingrese el Año y el Mes.")
        n_externo = self.bd_rma["CU"].count() if self.bd_rma is not None else 0
        if self.bd_rma is None:
            self.log("Ratio de insercion sin control externo (no se proceso la pestana RMA).")
        df = calc.ratio_insercion(self.res.bd_total, n_externo, anio, mes)
        self.llenar_tabla(self.RATINS, df)
        negrita = QFont()
        negrita.setBold(True)
        gris = QBrush(QColor(200, 200, 200))
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = self.RATINS.item(i, j)
                if i < 2 or i == df.shape[0] - 1:
                    item.setFont(negrita)
                if i < 2 or j == 0:
                    item.setBackground(gris)
        self.RATINS.resizeColumnsToContents()


def main():
    app = QApplication.instance() or QApplication(sys.argv)
    if cfg.FECHA_LIMITE is not None and datetime.now() > cfg.FECHA_LIMITE:
        QMessageBox.warning(None, "Error", "Falta algunos componentes")
        return 1
    window = MyApp()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
