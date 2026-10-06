# -*- coding: utf-8 -*-
"""
Parametros de los estandares (materiales de referencia).

Los valores ya no estan fijos en el codigo: se guardan en `estandares.json`
(junto al .exe) y se editan desde la aplicacion (Herramientas > Parametros de
estandares) o importando el Excel de estandares con el formato habitual:

    Estandar | Elemento | Valor Esperado BV (Media, CV%) | Valor Historico (Media, Std. Dev.)

Cada cambio guardado se registra en `historial_estandares.csv`
(fecha, usuario, estandar, elemento, campo, valor anterior, valor nuevo).

@author: frecines
"""
import csv
import getpass
import hashlib
import json
import os
import re
import sys
from datetime import datetime

import numpy as np
import pandas as pd

import qaqc_config as cfg

ARCHIVO = "estandares.json"
HISTORIAL = "historial_estandares.csv"
CLAVE_INICIAL = "qaqc"          # se recomienda cambiarla en el primer uso
CAMPOS = ("BV", "CV", "MH", "DE")   # CV: coeficiente de variacion (%) del BV
_SAL = "ModuloQAQC-estandares"

# Valores de fabrica. STD-16/17/18: archivo Estandares_2026.xlsx
# (BV = media del valor esperado, CV = CV (%) del valor esperado,
#  MH = media historica AAS, DE = Std. Dev.).
# STD-13/14/15 no tienen CV cargado: se grafican todos sus elementos.
PREDETERMINADOS = {
    "STD-13": dict(
        codigo="CHNB13", titulo="Estandar Baja Ley CHNB-13",
        MH={"CU": 0.266184914841849, "AG": 2.6863779342723,
            "MO": 0.00419716312056736, "AS": 0.0072113022113022,
            "ZN": 0.0253397590361446},
        BV={"CU": 0.266031884057971, "MO": 0.00423220338983051,
            "ZN": 0.0263231884057971},
        DE={"CU": 0.0036884613991108, "AG": 0.184056641873746,
            "MO": 0.000121177044741277, "AS": 0.000701560557845971,
            "ZN": 0.000902350628509791}),
    "STD-14": dict(
        codigo="CHNM14", titulo="Estandar Media Ley CHNM-14",
        MH={"CU": 0.427990543735225, "AG": 4.6302027972028,
            "MO": 0.00978690476190473, "AS": 0.016616091954023,
            "ZN": 0.0456080760095013},
        BV={"CU": 0.420332352941176, "MO": 0.00996842105263158,
            "AS": 0.0162754385964912, "ZN": 0.0453863636363636},
        DE={"CU": 0.00745666186049466, "AG": 0.177246972191448,
            "MO": 0.000271115331782672, "AS": 0.000827086603439742,
            "ZN": 0.00114052258444863}),
    "STD-15": dict(
        codigo="CHNA15", titulo="Estandar Alta Ley CHNA-15",
        MH={"CU": 0.662866666666667, "AG": 10.3921145584725,
            "MO": 0.0104511961722488, "AS": 0.0685093896713616,
            "ZN": 0.155243309002433},
        BV={"CU": 0.676945762711864, "AG": 10.140350877193,
            "MO": 0.0105706896551724, "AS": 0.07035, "ZN": 0.149810294117647},
        DE={"CU": 0.0119090159918189, "AG": 0.181529114416518,
            "MO": 0.000309809262563219, "AS": 0.00221782216510573,
            "ZN": 0.00304505560834989}),
    "STD-16": dict(
        codigo="CHNB16", titulo="Estandar Baja Ley CHNB-16",
        MH={"CU": 0.2454732824427479, "AG": 2.3889445843828705,
            "MO": 0.0041423558897243015, "AS": 0.006857506361323132,
            "ZN": 0.019901763224181256},
        BV={"CU": 0.241131884057971, "AG": 2.16432051765731,
            "MO": 0.00415828573173739, "AS": 0.00576315789473684,
            "ZN": 0.0199415384615385},
        CV={"CU": 2.029311707960067, "AG": 5.800195734231601,
            "MO": 3.8542348273505223, "AS": 3.404654345017692,
            "ZN": 3.123980532837728},
        DE={"CU": 0.004019077422834363, "AG": 0.20294772664318628,
            "MO": 9.660292066263082e-05, "AS": 0.000717884427758113,
            "ZN": 0.0007405443768598027}),
    "STD-17": dict(
        codigo="CHNM17", titulo="Estandar Media Ley CHNM-17",
        MH={"CU": 0.3985618556701031, "AG": 4.261017721518988,
            "MO": 0.009817179487179462, "AS": 0.014188295165394363,
            "ZN": 0.04493540051679585},
        BV={"CU": 0.384075, "AG": 3.89402242493882,
            "MO": 0.0100505386682992, "AS": 0.0128643868798608,
            "ZN": 0.0451897058823529},
        CV={"CU": 2.0752733140195456, "AG": 5.684625055019337,
            "MO": 2.410471781991903, "AS": 4.7575557626144915,
            "ZN": 3.7119747884539906},
        DE={"CU": 0.007961391893926327, "AG": 0.21137565422214796,
            "MO": 0.00010868337347859412, "AS": 0.0008807263899075752,
            "ZN": 0.0011599814775632804}),
    "STD-18": dict(
        codigo="CHNA18", titulo="Estandar Alta Ley CHNA-18",
        MH={"CU": 0.6043254593175857, "AG": 7.947753886010362,
            "MO": 0.016067684478371494, "AS": 0.029104999999999905,
            "ZN": 0.12881889763779542},
        BV={"CU": 0.596746268656716, "AG": 7.398615536568704,
            "MO": 0.0165168392857143, "AS": 0.0296482142857143,
            "ZN": 0.128048571428571},
        CV={"CU": 2.0695605732018865, "AG": 5.95927653584465,
            "MO": 1.8251161745541933, "AS": 2.9954632398196135,
            "ZN": 3.2972805108297796},
        DE={"CU": 0.004256772636084097, "AG": 0.24325494946538112,
            "MO": 0.0002930295349540777, "AS": 0.0006820101499329248,
            "ZN": 0.0027930854176019585}),
}


# ------------------------------------------------------------- archivos
def carpeta_datos():
    """Carpeta del .exe (o del script). Si no se puede escribir alli (p. ej.
    'Archivos de programa'), se usa %APPDATA%/ModuloQAQC."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    if os.access(base, os.W_OK):
        return base
    alterna = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ModuloQAQC")
    os.makedirs(alterna, exist_ok=True)
    return alterna


def hash_clave(clave):
    return hashlib.sha256(f"{_SAL}:{clave}".encode("utf-8")).hexdigest()


def _copia(estandares):
    return {c: dict(codigo=v["codigo"], titulo=v["titulo"],
                    **{k: dict(v.get(k, {})) for k in CAMPOS})
            for c, v in estandares.items()}


class Parametros:
    """Estandares vigentes + datos del archivo donde se guardan."""

    def __init__(self, carpeta=None):
        self.carpeta = carpeta or carpeta_datos()
        self.ruta = os.path.join(self.carpeta, ARCHIVO)
        self.ruta_historial = os.path.join(self.carpeta, HISTORIAL)
        self.estandares = _copia(PREDETERMINADOS)
        self.clave = hash_clave(CLAVE_INICIAL)
        self.actualizado = "valores de fabrica"
        self.aviso = None
        if os.path.exists(self.ruta):
            try:
                self._leer()
            except (OSError, ValueError, KeyError, TypeError) as e:
                self.aviso = (f"No se pudo leer {self.ruta} ({e}); se usan los "
                              "valores de fabrica.")

    def _leer(self):
        with open(self.ruta, encoding="utf-8") as f:
            datos = json.load(f)
        estandares = {}
        for e in datos["estandares"]:
            estandares[e["control"]] = dict(
                codigo=e.get("codigo") or e["control"],
                titulo=e.get("titulo") or f"Estandar {e['control']}",
                **{k: {el: float(v) for el, v in e.get(k, {}).items() if v is not None}
                   for k in CAMPOS})
        validar(estandares)
        self.estandares = estandares
        self.clave = datos.get("clave", self.clave)
        self.actualizado = f"{datos.get('fecha', '?')} por {datos.get('usuario', '?')}"

    def clave_correcta(self, clave):
        return hash_clave(clave) == self.clave

    def guardar(self, estandares, clave_hash=None):
        """Guarda los estandares y registra las diferencias en el historial.
        Devuelve la lista de cambios."""
        validar(estandares)
        cambios = diferencias(self.estandares, estandares)
        usuario, fecha = getpass.getuser(), datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        datos = {"version": 1, "fecha": fecha, "usuario": usuario,
                 "clave": clave_hash or self.clave,
                 "estandares": [{"control": c, **v} for c, v in estandares.items()]}
        temporal = self.ruta + ".tmp"
        with open(temporal, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
        os.replace(temporal, self.ruta)       # escritura atomica
        if cambios or clave_hash:
            nuevo = not os.path.exists(self.ruta_historial)
            with open(self.ruta_historial, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                if nuevo:
                    w.writerow(["Fecha", "Usuario", "Estandar", "Elemento", "Campo",
                                "Valor anterior", "Valor nuevo"])
                for cambio in cambios:
                    w.writerow([fecha, usuario, *cambio])
                if clave_hash:
                    w.writerow([fecha, usuario, "", "", "Contrasena", "", "cambiada"])
        self.estandares = _copia(estandares)
        self.clave = datos["clave"]
        self.actualizado = f"{fecha} por {usuario}"
        return cambios


def validar(estandares):
    for control, conf in estandares.items():
        if not re.fullmatch(r"\S+", control or ""):
            raise ValueError(f"Nombre de control invalido: '{control}'")
        for e in cfg.ELEMENTOS_STD:
            for campo in ("MH", "DE"):
                valor = conf[campo].get(e)
                if valor is None or not np.isfinite(valor) or valor <= 0:
                    raise ValueError(f"{control} {cfg.ETIQUETA[e]}: falta {campo} "
                                     "(debe ser un numero mayor que 0).")
            bv = conf["BV"].get(e)
            if bv is not None and (not np.isfinite(bv) or bv <= 0):
                raise ValueError(f"{control} {cfg.ETIQUETA[e]}: BV debe ser mayor que 0 "
                                 "o quedar vacio (no certificado).")
            cv = conf.get("CV", {}).get(e)
            if cv is not None and (not np.isfinite(cv) or cv < 0):
                raise ValueError(f"{control} {cfg.ETIQUETA[e]}: CV debe ser un numero >= 0 "
                                 "o quedar vacio.")


def diferencias(antes, despues):
    """[(control, elemento, campo, anterior, nuevo)] entre dos versiones."""
    cambios = []
    for control in sorted(set(antes) | set(despues)):
        a, d = antes.get(control), despues.get(control)
        if a is None or d is None:
            cambios.append((control, "", "Estandar", "" if a is None else "existia",
                            "agregado" if a is None else "eliminado"))
            if a is None:
                a = {"codigo": "", "titulo": "", **{k: {} for k in CAMPOS}}
            else:
                continue
        for campo in ("codigo", "titulo"):
            if a[campo] != d[campo]:
                cambios.append((control, "", campo, a[campo], d[campo]))
        for campo in CAMPOS:
            for e in cfg.ELEMENTOS_STD:
                va, vd = a.get(campo, {}).get(e), d.get(campo, {}).get(e)
                if va != vd and not (va is not None and vd is not None and np.isclose(va, vd, rtol=0, atol=1e-15)):
                    cambios.append((control, cfg.ETIQUETA[e], campo,
                                    "" if va is None else repr(va), "" if vd is None else repr(vd)))
    return cambios


# --------------------------------------------------------- Excel de estandares
_ELEMENTO = {"cu": "CU", "ag": "AG", "mo": "MO", "as": "AS", "zn": "ZN"}


def _texto(v):
    return "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()


def leer_excel(ruta, actuales=None):
    """Lee el Excel de estandares (formato Estandares_2026.xlsx).

    Columnas reconocidas por su encabezado: 'Estandar', 'Elemento',
    'Valor Esperado BV' (Media), 'Valor Historico' (Media y Std. Dev.).
    Opcionales: 'Codigo' y 'Descripcion'. Si no estan, se conservan los del
    estandar existente o se usa el nombre del control.
    """
    crudo = pd.read_excel(ruta, header=None, engine="calamine" if _hay_calamine() else None)
    fila_titulo = next((i for i in range(min(len(crudo), 15))
                        if any("estandar" in _sin_tilde(_texto(v)) for v in crudo.iloc[i])), None)
    if fila_titulo is None:
        raise ValueError("No se encontro el encabezado 'Estandar' en el Excel.")
    titulos = [_sin_tilde(_texto(v)) for v in crudo.iloc[fila_titulo]]
    sub = [_sin_tilde(_texto(v)) for v in crudo.iloc[fila_titulo + 1]]
    # Encabezados combinados: propagar el grupo hacia la derecha
    grupo = []
    for t in titulos:
        grupo.append(t or (grupo[-1] if grupo else ""))

    def columna(condicion, nombre):
        for j, (g, s) in enumerate(zip(grupo, sub)):
            if condicion(g, s):
                return j
        raise ValueError(f"No se encontro la columna '{nombre}' en el Excel.")

    c_std = columna(lambda g, s: "estandar" in g, "Estandar")
    c_el = columna(lambda g, s: "elemento" in g, "Elemento")
    c_bv = columna(lambda g, s: ("esperado" in g or "bv" in g) and "media" in s, "Valor Esperado BV - Media")
    try:  # opcional
        c_cv = columna(lambda g, s: ("esperado" in g or "bv" in g) and "cv" in s, "CV")
    except ValueError:
        c_cv = None
    c_mh = columna(lambda g, s: "historic" in g and "media" in s, "Valor Historico - Media")
    c_de = columna(lambda g, s: "historic" in g and ("dev" in s or "desv" in s), "Valor Historico - Std. Dev.")
    opcional = {}
    for clave, palabras in (("codigo", ("codigo",)), ("titulo", ("descripcion", "nombre"))):
        opcional[clave] = next((j for j, g in enumerate(titulos)
                                if any(p in g for p in palabras)), None)

    actuales = actuales or {}
    nuevos, elemento = {}, None
    for _, fila in crudo.iloc[fila_titulo + 2:].iterrows():
        control = _texto(fila.iloc[c_std])
        el_texto = _texto(fila.iloc[c_el])
        if el_texto:
            elemento = _ELEMENTO.get(re.split(r"[\s(]", el_texto.lower())[0])
        if not control or elemento is None:
            continue
        previo = actuales.get(control, {})
        conf = nuevos.setdefault(control, dict(
            codigo=previo.get("codigo", control), titulo=previo.get("titulo", f"Estandar {control}"),
            BV={}, CV={}, MH={}, DE={}))
        for clave, j in opcional.items():
            if j is not None and _texto(fila.iloc[j]):
                conf[clave] = _texto(fila.iloc[j])
        for campo, j in (("BV", c_bv), ("CV", c_cv), ("MH", c_mh), ("DE", c_de)):
            if j is None:
                continue
            valor = pd.to_numeric(fila.iloc[j], errors="coerce")
            if pd.notna(valor):
                conf[campo][elemento] = float(valor)
    if not nuevos:
        raise ValueError("El Excel no tiene filas de estandares reconocibles.")
    return nuevos


def exportar_excel(estandares, ruta):
    """Exporta en el mismo formato que se puede volver a importar."""
    filas = []
    for e in cfg.ELEMENTOS_STD:
        for control, conf in estandares.items():
            filas.append([control, f"{cfg.ETIQUETA[e]} ({'g/T' if e == 'AG' else '%'})",
                          conf["BV"].get(e), conf.get("CV", {}).get(e),
                          conf["MH"].get(e), conf["DE"].get(e),
                          conf["codigo"], conf["titulo"]])
    with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
        hoja = w.book.add_worksheet("Estandares")
        negrita = w.book.add_format({"bold": True, "align": "center", "border": 1})
        for j, t in enumerate(["Estándar", "Elemento", "Valor Esperado BV", "",
                               "Valor Histórico (AAS)", "", "Código", "Descripción"]):
            hoja.write(0, j, t, negrita)
        for j, t in enumerate(["", "", "Media", "CV (%)", "Media", "Std. Dev.", "", ""]):
            hoja.write(1, j, t, negrita)
        for i, fila in enumerate(filas, start=2):
            for j, v in enumerate(fila):
                if v is not None:
                    hoja.write(i, j, v)
        hoja.set_column(0, 1, 12)
        hoja.set_column(2, 5, 16)
        hoja.set_column(6, 7, 28)


def _sin_tilde(texto):
    return (texto.lower().replace("á", "a").replace("é", "e").replace("í", "i")
            .replace("ó", "o").replace("ú", "u"))


def _hay_calamine():
    try:
        import python_calamine  # noqa: F401
        return True
    except ImportError:
        return False


def colores(estandares):
    """Color de fondo para el nombre de cada estandar en las tablas: ley baja
    (CHNB..) amarillo, media (CHNM..) naranja, alta (CHNA..) verde, como en los
    informes anteriores; otros codigos toman colores de una paleta."""
    por_ley = {"CHNB": "yellow", "CHNM": "orange", "CHNA": "lightgreen"}
    paleta = ["lightblue", "pink", "khaki", "plum", "wheat"]
    resultado, k = {}, 0
    for conf in estandares.values():
        color = por_ley.get(conf["codigo"][:4].upper())
        if color is None:
            color, k = paleta[k % len(paleta)], k + 1
        resultado[conf["codigo"]] = color
    return resultado


def categoria_cv(cv):
    """Certificado / Provisional / Indicativo segun el CV (%) del BV; None si no hay CV."""
    if cv is None or pd.isna(cv):
        return None
    if cv <= cfg.UMBRAL_CV_ESTANDAR:
        return "Certificado"
    return "Provisional" if cv <= cfg.UMBRAL_CV_INDICATIVO else "Indicativo"


def es_referencial(conf, elemento):
    """True si el BV del elemento tiene CV > 5%: se grafica como valor
    referencial y no se informa su sesgo."""
    return categoria_cv(conf.get("CV", {}).get(elemento)) in ("Provisional", "Indicativo")


def elementos_referenciales(conf):
    """[(elemento, cv, categoria)] con CV del BV mayor a 5%."""
    return [(e, conf["CV"][e], categoria_cv(conf["CV"][e]))
            for e in cfg.ELEMENTOS_STD if es_referencial(conf, e)]
