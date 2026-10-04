# -*- mode: python ; coding: utf-8 -*-
# Compilacion:  pyinstaller qaqc.spec --noconfirm   (o ejecutar build_exe.bat)
#
# UN_SOLO_ARCHIVO = True  -> dist/Modulo QAQC.exe (un solo archivo, como antes).
# UN_SOLO_ARCHIVO = False -> dist/Modulo QAQC/Modulo QAQC.exe (carpeta; abre mas
#                            rapido porque no descomprime todo en cada inicio).
UN_SOLO_ARCHIVO = True
NOMBRE = "Modulo QAQC"
ICONO = "Grid.ico"  # si el archivo no existe se usa el icono por defecto

import os

from PyInstaller.utils.hooks import collect_data_files

# Librerias que el entorno suele tener instaladas pero el modulo no usa:
# excluirlas reduce el tamano y el tiempo de inicio.
EXCLUIR = [
    "tkinter", "_tkinter", "IPython", "jupyter", "notebook", "nbformat",
    "scipy", "statsmodels", "sklearn", "numba", "sympy", "pytest",
    "PySide2", "PySide6", "PyQt6",
    "PyQt5.QtWebEngine", "PyQt5.QtWebEngineCore", "PyQt5.QtWebEngineWidgets",
    "PyQt5.QtMultimedia", "PyQt5.QtQml", "PyQt5.QtQuick", "PyQt5.QtBluetooth",
    "PyQt5.QtSql", "PyQt5.QtNetwork", "PyQt5.QtDesigner", "PyQt5.QtTest",
    "matplotlib.backends.backend_tkagg", "matplotlib.backends.backend_webagg",
    "matplotlib.backends.backend_gtk3agg", "matplotlib.backends.backend_wxagg",
    "pandas.io.formats.style", "jinja2", "pyarrow", "tables", "sqlalchemy",
]

a = Analysis(
    ["qaqc.py"],
    pathex=[],
    binaries=[],
    # El .ui solo se usa si falta ui_qaqcbs.py (respaldo). python-docx necesita
    # sus plantillas (encabezado/pie) para generar el informe Word.
    # include_py_files: crea la carpeta docx/parts que usa la ruta "parts/../templates".
    datas=[("qaqcbs.ui", ".")] + collect_data_files("docx", include_py_files=True),
    # pandas carga los motores de Excel de forma dinamica
    hiddenimports=["python_calamine", "openpyxl", "xlsxwriter",
                   "matplotlib.backends.backend_qt5agg", "qaqc_plots",
                   "qaqc_dialogo_estandares", "qaqc_informe", "docx"],
    excludes=EXCLUIR,
    noarchive=False,
)
# Quitar traducciones de Qt y modulos QML/Quick que arrastran los plugins
NO_USADOS = ("translations", "Qt5Qml", "Qt5Quick", "Qt5WebSockets", "Qt5Pdf")
a.binaries = [b for b in a.binaries if not any(n in b[0] for n in NO_USADOS)]
a.datas = [d for d in a.datas if not any(n in d[0] for n in NO_USADOS)]

pyz = PYZ(a.pure)
icono = ICONO if os.path.exists(ICONO) else None

# UPX desactivado: comprimir las DLL de Qt con UPX suele dañarlas y genera
# falsos positivos de antivirus.
if UN_SOLO_ARCHIVO:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=NOMBRE,
              console=False, upx=False, icon=icono)
else:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=NOMBRE,
              console=False, upx=False, icon=icono)
    coll = COLLECT(exe, a.binaries, a.datas, name=NOMBRE, upx=False)
