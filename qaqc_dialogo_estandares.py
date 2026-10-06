# -*- coding: utf-8 -*-
"""
Ventana 'Parametros de Estandares'.

Muestra los valores vigentes en modo solo lectura. Para modificarlos hay que
desbloquear con la contrasena; los cambios (manuales o importados desde el
Excel de estandares) solo se aplican al presionar Guardar, que pide
confirmacion y deja registro en el historial.

@author: frecines
"""
import copy
import os

from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QColor, QDesktopServices
from PyQt5.QtWidgets import QFileDialog, QInputDialog, QLineEdit, QMessageBox

import qaqc_config as cfg
import qaqc_estandares as std

COL_STD = ["Control (en CONTROLES)", "Código", "Descripción (título de gráficos)"]
COL_VAL = ["Elemento", "BV - Valor esperado\n(vacío = no certificado)",
           "CV del BV (%)\n(> 5% = valor referencial)", "MH - Media histórica",
           "DE - Desviación estándar"]
CAMPOS_VAL = ("BV", "CV", "MH", "DE")


class DialogoEstandares(QtWidgets.QDialog):
    def __init__(self, parametros, padre=None):
        super().__init__(padre)
        self.param = parametros
        self.trabajo = copy.deepcopy(parametros.estandares)
        self.clave_nueva = None
        self.editable = False
        self._cargando = False
        self.setWindowTitle("Parámetros de Estándares")
        self.resize(900, 640)

        self.lblInfo = QtWidgets.QLabel()
        self.lblInfo.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lblEstado = QtWidgets.QLabel()

        self.tblStd = QtWidgets.QTableWidget(0, len(COL_STD))
        self.tblStd.setHorizontalHeaderLabels(COL_STD)
        self.tblStd.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tblStd.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.tblStd.horizontalHeader().setStretchLastSection(True)
        self.tblVal = QtWidgets.QTableWidget(len(cfg.ELEMENTOS_STD), len(COL_VAL))
        self.tblVal.setHorizontalHeaderLabels(COL_VAL)
        self.tblVal.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tblVal.verticalHeader().hide()

        def boton(texto, slot, edicion=False):
            b = QtWidgets.QPushButton(texto)
            b.clicked.connect(lambda _=False: slot())
            b.setProperty("edicion", edicion)
            return b

        self.btnDesbloquear = boton("Desbloquear edición...", self.desbloquear)
        self.botones_edicion = [
            boton("Importar desde Excel...", self.importar, True),
            boton("Agregar estándar...", self.agregar, True),
            boton("Eliminar estándar", self.eliminar, True),
            boton("Cambiar contraseña...", self.cambiar_clave, True),
        ]
        self.btnGuardar = boton("Guardar", self.guardar, True)
        otros = [boton("Exportar a Excel...", self.exportar),
                 boton("Ver historial de cambios", self.ver_historial)]
        cerrar = boton("Cerrar", self.reject)

        fila1 = QtWidgets.QHBoxLayout()
        for b in [self.btnDesbloquear, *self.botones_edicion]:
            fila1.addWidget(b)
        fila1.addStretch()
        fila2 = QtWidgets.QHBoxLayout()
        for b in otros:
            fila2.addWidget(b)
        fila2.addStretch()
        fila2.addWidget(self.btnGuardar)
        fila2.addWidget(cerrar)

        capa = QtWidgets.QVBoxLayout(self)
        capa.addWidget(self.lblInfo)
        capa.addWidget(self.lblEstado)
        capa.addLayout(fila1)
        capa.addWidget(QtWidgets.QLabel("<b>Estándares</b>"))
        capa.addWidget(self.tblStd, 2)
        capa.addWidget(QtWidgets.QLabel("<b>Valores del estándar seleccionado</b>"))
        capa.addWidget(self.tblVal, 3)
        capa.addLayout(fila2)

        self.tblStd.itemSelectionChanged.connect(self._mostrar_valores)
        self.tblStd.itemChanged.connect(self._std_editado)
        self.tblVal.itemChanged.connect(self._valor_editado)
        self._actualizar_info()
        self._llenar_estandares()
        self._modo_edicion(False)

    # --------------------------------------------------------------- vista
    def _actualizar_info(self):
        self.lblInfo.setText(f"Archivo: {self.param.ruta}\n"
                             f"Última actualización: {self.param.actualizado}")

    def _modo_edicion(self, activo):
        self.editable = activo
        for b in self.botones_edicion + [self.btnGuardar]:
            b.setEnabled(activo)
        self.btnDesbloquear.setEnabled(not activo)
        disparadores = (QtWidgets.QAbstractItemView.DoubleClicked
                        | QtWidgets.QAbstractItemView.EditKeyPressed
                        | QtWidgets.QAbstractItemView.AnyKeyPressed) if activo else \
            QtWidgets.QAbstractItemView.NoEditTriggers
        self.tblStd.setEditTriggers(disparadores)
        self.tblVal.setEditTriggers(disparadores)
        self.lblEstado.setText(
            "<span style='color:#b35c00'><b>Edición habilitada.</b> Los cambios se aplican "
            "al presionar Guardar.</span>" if activo else
            "<span style='color:gray'>Solo lectura. Use 'Desbloquear edición' para "
            "modificar los valores.</span>")

    def _llenar_estandares(self, seleccionar=None):
        self._cargando = True
        self.tblStd.setRowCount(len(self.trabajo))
        for i, (control, conf) in enumerate(self.trabajo.items()):
            for j, texto in enumerate([control, conf["codigo"], conf["titulo"]]):
                item = QtWidgets.QTableWidgetItem(texto)
                if j == 0:  # el control se define al agregar el estandar
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QColor("lightgray"))
                self.tblStd.setItem(i, j, item)
        self.tblStd.resizeColumnsToContents()
        self._cargando = False
        controles = list(self.trabajo)
        fila = controles.index(seleccionar) if seleccionar in controles else 0
        if controles:
            self.tblStd.selectRow(fila)
        self._mostrar_valores()

    def _control_actual(self):
        filas = self.tblStd.selectionModel().selectedRows()
        if not filas:
            return None
        return self.tblStd.item(filas[0].row(), 0).text()

    def _mostrar_valores(self):
        self._cargando = True
        control = self._control_actual()
        conf = self.trabajo.get(control)
        for i, e in enumerate(cfg.ELEMENTOS_STD):
            nombre = QtWidgets.QTableWidgetItem(cfg.ANALITO[e])
            nombre.setFlags(nombre.flags() & ~Qt.ItemIsEditable)
            nombre.setBackground(QColor("lightgray"))
            self.tblVal.setItem(i, 0, nombre)
            for j, campo in enumerate(CAMPOS_VAL, start=1):
                valor = conf.get(campo, {}).get(e) if conf else None
                self.tblVal.setItem(i, j, QtWidgets.QTableWidgetItem(
                    "" if valor is None else repr(valor)))
        self._cargando = False

    # ------------------------------------------------------------- edicion
    def _std_editado(self, item):
        if self._cargando:
            return
        control = self.tblStd.item(item.row(), 0).text()
        campo = "codigo" if item.column() == 1 else "titulo"
        self.trabajo[control][campo] = item.text().strip()

    def _valor_editado(self, item):
        if self._cargando or item.column() == 0:
            return
        control = self._control_actual()
        elemento = cfg.ELEMENTOS_STD[item.row()]
        campo = CAMPOS_VAL[item.column() - 1]
        texto = item.text().strip().replace(",", ".")
        if not texto:
            self.trabajo[control].setdefault(campo, {}).pop(elemento, None)
            return
        try:
            self.trabajo[control].setdefault(campo, {})[elemento] = float(texto)
        except ValueError:
            QMessageBox.warning(self, "Valor inválido", f"'{item.text()}' no es un número.")
            self._mostrar_valores()

    def desbloquear(self):
        clave, ok = QInputDialog.getText(self, "Desbloquear edición", "Contraseña:",
                                         QLineEdit.Password)
        if not ok:
            return
        if not self.param.clave_correcta(clave):
            QMessageBox.warning(self, "Contraseña incorrecta", "La contraseña no es válida.")
            return
        self._modo_edicion(True)

    def importar(self):
        ruta, _ = QFileDialog.getOpenFileName(self, "Excel de estándares", "",
                                              "Archivos Excel (*.xlsx *.xlsm *.xls)")
        if not ruta:
            return
        try:
            nuevos = std.leer_excel(ruta, self.trabajo)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "No se pudo importar", str(e))
            return
        resultado = dict(self.trabajo)
        resultado.update(nuevos)
        cambios = std.diferencias(self.trabajo, resultado)
        texto = "\n".join(f"{c[0]} {c[1]} {c[2]}: {c[3] or '(vacío)'} → {c[4] or '(vacío)'}"
                          for c in cambios[:25])
        if len(cambios) > 25:
            texto += f"\n... y {len(cambios) - 25} cambios más"
        if QMessageBox.question(
                self, "Importar estándares",
                f"Se leyeron {len(nuevos)} estándares ({', '.join(nuevos)}).\n"
                f"{len(cambios)} valores cambian:\n\n{texto or '(ninguno)'}\n\n"
                "¿Aplicar? (luego presione Guardar)") == QMessageBox.Yes:
            self.trabajo = resultado
            self._llenar_estandares(next(iter(nuevos)))

    def agregar(self):
        control, ok = QInputDialog.getText(
            self, "Agregar estándar",
            "Código tal como aparece en la columna CONTROLES (ej. STD-19):")
        control = control.strip()
        if not ok or not control:
            return
        if control in self.trabajo:
            QMessageBox.warning(self, "Ya existe", f"{control} ya está en la lista.")
            return
        self.trabajo[control] = dict(codigo=control, titulo=f"Estandar {control}",
                                     BV={}, CV={}, MH={}, DE={})
        self._llenar_estandares(control)
        QMessageBox.information(self, "Agregar estándar",
                                "Complete MH y DE de los 5 elementos (BV y CV son opcionales) "
                                "y luego presione Guardar.")

    def eliminar(self):
        control = self._control_actual()
        if control and QMessageBox.question(
                self, "Eliminar estándar",
                f"¿Eliminar {control}? Sus muestras dejarán de evaluarse.") == QMessageBox.Yes:
            del self.trabajo[control]
            self._llenar_estandares()

    def cambiar_clave(self):
        nueva, ok = QInputDialog.getText(self, "Cambiar contraseña", "Nueva contraseña:",
                                         QLineEdit.Password)
        if not ok:
            return
        repetida, ok = QInputDialog.getText(self, "Cambiar contraseña", "Repita la contraseña:",
                                            QLineEdit.Password)
        if not ok:
            return
        if len(nueva) < 4 or nueva != repetida:
            QMessageBox.warning(self, "Cambiar contraseña",
                                "Las contraseñas no coinciden o tienen menos de 4 caracteres.")
            return
        self.clave_nueva = std.hash_clave(nueva)
        QMessageBox.information(self, "Cambiar contraseña",
                                "La nueva contraseña se aplicará al presionar Guardar.")

    def _pendientes(self):
        return std.diferencias(self.param.estandares, self.trabajo)

    def guardar(self):
        try:
            std.validar(self.trabajo)
        except ValueError as e:
            QMessageBox.warning(self, "Datos incompletos", str(e))
            return
        cambios = self._pendientes()
        if not cambios and not self.clave_nueva:
            QMessageBox.information(self, "Guardar", "No hay cambios para guardar.")
            return
        if QMessageBox.question(
                self, "Guardar",
                f"Se guardarán {len(cambios)} cambios y quedarán registrados en el "
                "historial. ¿Continuar?") != QMessageBox.Yes:
            return
        try:
            self.param.guardar(self.trabajo, self.clave_nueva)
        except OSError as e:
            QMessageBox.critical(self, "No se pudo guardar", str(e))
            return
        self.clave_nueva = None
        self._actualizar_info()
        self._modo_edicion(False)
        self.accept()

    def reject(self):
        """Cerrar, Esc o la X de la ventana: avisar si hay cambios sin guardar."""
        if self.editable and (self._pendientes() or self.clave_nueva):
            if QMessageBox.question(self, "Cambios sin guardar",
                                    "Hay cambios sin guardar. ¿Descartarlos?") != QMessageBox.Yes:
                return
        super().reject()

    # -------------------------------------------------------------- otros
    def exportar(self):
        ruta, _ = QFileDialog.getSaveFileName(self, "Exportar estándares",
                                              "Estandares_vigentes.xlsx",
                                              "Archivos Excel (*.xlsx)")
        if ruta:
            try:
                std.exportar_excel(self.trabajo, ruta)
            except OSError as e:
                QMessageBox.critical(self, "No se pudo exportar",
                                     f"{e}\n\nRevisar que el archivo no esté abierto.")

    def ver_historial(self):
        if not os.path.exists(self.param.ruta_historial):
            QMessageBox.information(self, "Historial", "Todavía no hay cambios registrados.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.param.ruta_historial))
