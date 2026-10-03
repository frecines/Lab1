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

from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import QRegularExpression, Qt
from PyQt5.QtGui import QBrush, QColor, QFont, QRegularExpressionValidator
from PyQt5.QtWidgets import (QApplication, QFileDialog, QMainWindow,
                             QMessageBox, QTableWidgetItem)

from matplotlib.figure import Figure
try:
    from matplotlib.backends.backend_qtagg import (
        FigureCanvasQTAgg as FigureCanvas, NavigationToolbar2QT as NavigationToolbar)
except ImportError:  # matplotlib < 3.5
    from matplotlib.backends.backend_qt5agg import (
        FigureCanvasQTAgg as FigureCanvas, NavigationToolbar2QT as NavigationToolbar)

import pandas as pd

import qaqc_calc as calc
import qaqc_config as cfg
import qaqc_plots as plots


def _ruta_recurso(nombre):
    """Ruta junto a este archivo (funciona aunque se ejecute desde otra carpeta
    o empaquetado con PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nombre)


Ui_MainWindow, _ = uic.loadUiType(_ruta_recurso("qaqcbs.ui"))

FILTRO_EXCEL = "Archivos Excel (*.xlsx *.xlsm *.xls)"
COLOR_CLASE = {"ACEPTABLE": "green", "CUESTIONABLE": "orange",
               "INACEPTABLE": "red", "No Certificado": "black", "Sin datos": "gray"}
GRIS = QColor("lightgray")


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
            except PermissionError as e:
                self.log(f"{titulo_error}: {e}", error=True)
                QMessageBox.critical(self, titulo_error,
                                     f"{e}\n\nRevisar que el archivo no este abierto en Excel.")
            except Exception as e:  # noqa: BLE001 - se informa al usuario
                self.log(f"{titulo_error}: {e}", error=True)
                traceback.print_exc()
                QMessageBox.critical(self, titulo_error, str(e))
        return envoltura
    return decorador


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
                             self.RDRG, self.GDP, self.QQDP, self.RDDP, self.CHNB,
                             self.CHNM, self.CHNA, self.CHNH, self.GRI]
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
            self.CHNB: lambda: self.graficos_estandar("B"),
            self.CHNM: lambda: self.graficos_estandar("M"),
            self.CHNA: lambda: self.graficos_estandar("A"),
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
        self._mostrar_carpeta()
        self.log("Seleccione la BD de Leyes para comenzar.")

    # ------------------------------------------------------------ utilidades
    def log(self, mensaje, error=False):
        hora = datetime.now().strftime("%H:%M:%S")
        color = "red" if error else "#1a4fa0"
        self.Log.appendHtml(f'<span style="color:gray">{hora}</span> '
                            f'<span style="color:{color}">{mensaje}</span>')
        self.statusbar.showMessage(mensaje, 8000)

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
    def _filtro(combo):
        return combo.currentText() if combo.currentIndex() > 0 else None

    # ---------------------------------------------------------------- tablas
    @staticmethod
    def llenar_tabla(tabla, df, col_estandar=None, gris_primera=True):
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
                    item.setBackground(QColor(cfg.COLOR_ESTANDAR.get(texto, "lightgray")))
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
        self.llenar_tabla(self.DFSTD, r["Estandares"], col_estandar=0)
        self.llenar_tabla(self.DFSTDB, r["Estandares por Analito"], col_estandar=1)
        errores = {self.ErrorBlancoF: "ErrorBlancoFino", self.ErrorBlancoG: "ErrorBlancoGrueso",
                   self.E_DUPC: "ErrorDuplicadoCampo", self.E_RG: "ErrorRechazoGrueso",
                   self.E_PULP: "ErrorDuplicadoPulpa", self.E_STDB: "ErrorCHNB",
                   self.E_STDM: "ErrorCHNM", self.E_STDA: "ErrorCHNA"}
        for tabla, hoja in errores.items():
            self.llenar_tabla(tabla, r[hoja], gris_primera=False)

    # ------------------------------------------------------- importacion
    @accion("Error al cargar el archivo")
    def importar_leyes(self):
        ruta = self._abrir_excel("Seleccionar archivo BD Leyes")
        if not ruta:
            return
        try:
            bd = calc.leer_bd_leyes(ruta)
        except ValueError as e:
            raise ValueError(f"{e}\nEl excel seleccionado debe tener la hoja 'General'.")
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
        self.log(f"BD con Leyes importada ({len(bd)} filas). Cargue el Reporte de Muestras QAQC.")

    @accion("Error al cargar el archivo")
    def importar_reporte(self):
        ruta = self._abrir_excel("Seleccionar archivo Reporte de Muestras QAQC")
        if not ruta:
            return
        try:
            bd = calc.leer_reporte_qaqc(ruta)
        except ValueError as e:
            raise ValueError(f"{e}\nEl excel seleccionado debe tener la hoja 'LEYES SGS'.")
        self.bd_sec = bd
        self.lblQaqc.setText(os.path.basename(ruta))
        for combo, columna, todos in ((self.VMalla, "Malla", "(Todas)"),
                                      (self.VLote, "Scheme", "(Todos)")):
            combo.clear()
            combo.addItems([todos, *calc.opciones_filtro(bd, columna)])
        self.PROC.setEnabled(True)
        self.log("Reporte de Muestras QAQC importado. Revise el rango y los filtros y "
                 "presione 'Procesar Datos'.")

    @accion("Error en el procesamiento")
    def procesar(self):
        vmin, vmax = self._rango()
        malla, reporte = self._filtro(self.VMalla), self._filtro(self.VLote)
        bd_total, avisos = calc.unir_bases(self.bd_ini, self.bd_sec, vmin, vmax, malla, reporte)
        for aviso in avisos:
            self.log(f"Aviso: {aviso}", error=True)
        self.res = calc.ResultadosQAQC(bd_total)
        self._llenar_tablas_qaqc()
        for boton in self.botones_qaqc:
            boton.setEnabled(True)
        # La union cambio: hay que volver a procesar el control externo
        self.bd_rma = None
        for boton in self.botones_rma:
            boton.setEnabled(False)
        self.ProcRMA.setEnabled(self.bd_cert is not None)

        filtros = [f"rango {vmin} - {vmax}"]
        filtros += [f"malla {malla}"] if malla else []
        filtros += [f"reporte {reporte}"] if reporte else []
        conteo = ", ".join(f"{k}: {v}" for k, v in self.res.conteo_controles().items())
        self.log(f"Procesado ({', '.join(filtros)}): {len(bd_total)} muestras. "
                 f"Controles -> {conteo or 'ninguno'}.")

    @accion("No se creo el reporte")
    def exportar_reporte(self):
        ruta = os.path.join(self._carpeta(), "Reporte_QAQC.xlsx")
        calc.exportar_excel(self.res.reportes, ruta)
        self.log(f"Reporte exportado: {ruta}")

    # -------------------------------------------------------------- graficos
    def _mostrar(self, titulo, graficos):
        """graficos: lista de (pestana, archivo_png, figsize, dibujar(fig))."""
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
        conf, df = cfg.ESTANDARES[clave], self.res.estandares[clave]
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
        self._mostrar(cfg.ESTANDARES[clave]["titulo"], self._lista_estandar(clave))

    @accion("Error al graficar")
    def graficos_historico(self):
        graficos = []
        for clave in cfg.BLANCOS:
            graficos += self._lista_blanco(clave, meses=True, prefijo=f"{clave} ")
        for clave, conf in cfg.ESTANDARES.items():
            graficos += self._lista_estandar(clave, meses=True, prefijo=f"{conf['codigo']} ")
        self._mostrar("Historico por Mes - Blancos y Estandares", graficos)

    # ------------------------------------------------------ control externo
    @accion("Error al cargar el archivo")
    def importar_externo(self):
        ruta = self._abrir_excel("Seleccionar reporte del laboratorio externo")
        if not ruta:
            return
        self.bd_cert = calc.leer_reporte_externo(ruta)
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
