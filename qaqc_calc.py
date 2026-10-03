# -*- coding: utf-8 -*-
"""
Calculos QAQC (sin interfaz grafica).

Contiene la lectura de archivos, la union de bases, el procesamiento de
blancos, duplicados, estandares, control externo y la generacion de los
cuadros de reporte. Al no depender de Qt se puede probar y reutilizar
desde scripts o notebooks.

@author: frecines
"""
import numpy as np
import pandas as pd

import qaqc_config as cfg

COLS_REPORTE_QAQC = ["SAMPLE_ID", "CONTROLES", "MES ", "Malla", "Scheme"]


# --------------------------------------------------------------- utilidades
def _requiere(df, columnas, origen):
    faltan = [c for c in columnas if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan las columnas {faltan} en {origen}. "
                         "Revisar si se modifico el encabezado del excel.")


def a_numero(serie, reemplazos):
    """Convierte una columna de leyes a float.

    Aplica primero la tabla de reemplazos ('<0.001' -> 0.0005, ...). Un
    valor '<x' no listado se toma como x/2 y un '>x' no listado como x.
    Cualquier otro texto queda como NaN en vez de romper los calculos.
    """
    if pd.api.types.is_numeric_dtype(serie):
        return serie.astype(float)

    def convertir(valor):
        if not isinstance(valor, str):
            return np.nan if pd.isna(valor) else float(valor)
        texto = valor.strip().replace(",", ".")
        if texto in reemplazos:
            return float(reemplazos[texto])
        try:
            return float(texto)
        except ValueError:
            pass
        if texto[:1] in ("<", ">"):
            try:
                numero = float(texto[1:])
            except ValueError:
                return np.nan
            return numero / 2 if texto[0] == "<" else numero
        return np.nan

    return serie.map(convertir).astype(float)


def _ids_enteros(df, origen):
    df = df.copy()
    df["SAMPLE_ID"] = pd.to_numeric(df["SAMPLE_ID"], errors="coerce")
    df = df.dropna(subset=["SAMPLE_ID"])
    if df.empty:
        raise ValueError(f"No hay SAMPLE_ID numericos en {origen}.")
    df["SAMPLE_ID"] = df["SAMPLE_ID"].astype("int64")
    return df


def _buscar(base, ids, columnas):
    """Valores de `columnas` para cada SAMPLE_ID de `ids` (NaN si no existe)."""
    tabla = base.drop_duplicates("SAMPLE_ID").set_index("SAMPLE_ID")
    return tabla.reindex(np.asarray(ids))[columnas].reset_index(drop=True)


def _bandera(condicion, valido):
    """1.0 / 0.0 segun la condicion, NaN cuando el dato no es valido."""
    return condicion.astype(float).where(valido)


def tasa(df, columna):
    """% de fallos de una columna bandera (ignora los pares sin dato)."""
    if df.empty or df[columna].count() == 0:
        return np.nan
    return df[columna].mean() * 100


def clasificar_sesgo(sesgo, certificado=True):
    if not certificado:
        return "No Certificado"
    if pd.isna(sesgo):
        return "Sin datos"
    if abs(sesgo) <= cfg.SESGO_ACEPTABLE:
        return "ACEPTABLE"
    if abs(sesgo) <= cfg.SESGO_CUESTIONABLE:
        return "CUESTIONABLE"
    return "INACEPTABLE"


def _clase_umbral(serie, umbral):
    return np.where(serie.isna(), "Sin datos",
                    np.where(serie > umbral, "INACEPTABLE", "ACEPTABLE"))


def _redondear(df, decimales=3):
    numericas = df.select_dtypes(include="number").columns
    df[numericas] = df[numericas].round(decimales)
    return df


# ---------------------------------------------------------- lectura/union
def leer_bd_leyes(ruta):
    return pd.read_excel(ruta, sheet_name="General")


def leer_reporte_qaqc(ruta):
    return pd.read_excel(ruta, sheet_name="LEYES SGS", header=1)


def opciones_filtro(bd_sec, columna):
    """Valores distintos (como texto) para llenar los combos de filtro."""
    if columna not in bd_sec.columns:
        return []
    valores = bd_sec[columna].dropna().astype(str).str.strip().unique()
    return sorted(v for v in valores if v)


def unir_bases(bd_ini, bd_sec, vmin, vmax, malla=None, reporte=None):
    """Une la BD de leyes con el reporte QAQC y aplica los filtros.

    Devuelve (bd_total, avisos).
    """
    _requiere(bd_ini, ["SAMPLE_ID", "LAB_REPORT", *cfg.COLUMNA_BD.values()],
              "la BD de Leyes (hoja 'General')")
    _requiere(bd_sec, COLS_REPORTE_QAQC,
              "el Reporte de Muestras QAQC (hoja 'LEYES SGS')")
    avisos = []

    leyes = _ids_enteros(bd_ini, "la BD de Leyes")
    for col in [*cfg.COLUMNA_BD.values(), "CUOX"]:
        if col in leyes.columns:
            leyes[col] = a_numero(leyes[col], cfg.REEMPLAZO_LD)

    columnas = COLS_REPORTE_QAQC + [c for c in ["FECHA DE ENVIO"] if c in bd_sec]
    sec = _ids_enteros(bd_sec[columnas], "el Reporte de Muestras QAQC")
    repetidos = int(sec["SAMPLE_ID"].duplicated().sum())
    if repetidos:
        avisos.append(f"{repetidos} SAMPLE_ID repetidos en el reporte QAQC; "
                      "se usa la primera aparicion.")
        sec = sec.drop_duplicates("SAMPLE_ID")
    sec["CONTROLES"] = sec["CONTROLES"].astype("string").str.strip()

    total = leyes.merge(sec, on="SAMPLE_ID", how="left")
    total = total[(total["SAMPLE_ID"] >= vmin) & (total["SAMPLE_ID"] <= vmax)]
    for columna, valor in (("Malla", malla), ("Scheme", reporte)):
        if valor is not None:
            ids = sec.loc[sec[columna].astype(str).str.strip() == str(valor),
                          "SAMPLE_ID"]
            total = total[total["SAMPLE_ID"].isin(ids)]
    if total.empty:
        raise ValueError("No hay muestras con el rango y los filtros seleccionados.")
    return total.sort_values("SAMPLE_ID").reset_index(drop=True), avisos


def _ids_control(bd_total, codigo):
    ids = bd_total.loc[bd_total["CONTROLES"] == codigo, "SAMPLE_ID"]
    return np.sort(ids.to_numpy())


# --------------------------------------------------------------- blancos
def procesar_blanco(bd_total, conf):
    ids = _ids_control(bd_total, conf["control"])
    cols = list(cfg.COLUMNA_BD.values())
    propio = _buscar(bd_total, ids, cols + ["LAB_REPORT", "MES "])
    previo = _buscar(bd_total, ids - conf["desfase_prec"], cols)

    df = pd.DataFrame({"SAMPLE_ID": ids,
                       "SAMPLE_PREC": ids - conf["desfase_prec"],
                       "LAB_REPORT": propio["LAB_REPORT"],
                       "MES": propio["MES "]})
    for e in cfg.ELEMENTOS:
        col = cfg.COLUMNA_BD[e]
        mav = conf["factor"] * conf["ld"][e]
        df[e] = propio[col]
        df[f"{e}_P"] = previo[col]
        df[f"{e}_MAV"] = mav
        df[f"{e}_ERR"] = _bandera(df[e] > mav, df[e].notna())
    return df


# ------------------------------------------------------------ duplicados
def procesar_duplicado(bd_total, conf):
    ids = _ids_control(bd_total, conf["control"])
    par = ids + conf["desfase_par"]
    orig, dup = (ids, par) if conf["desfase_par"] > 0 else (par, ids)
    cols = list(cfg.COLUMNA_BD.values())
    v_orig = _buscar(bd_total, orig, cols)
    v_dup = _buscar(bd_total, dup, cols)
    # Igual que el original: LAB_REPORT de la muestra pareja
    lab = _buscar(bd_total, par, ["LAB_REPORT"])["LAB_REPORT"]

    df = pd.DataFrame({"SAMPLE_ID": ids, "SAMPLE_ORIG": orig,
                       "SAMPLE_DUP": dup, "LAB_REPORT": lab})
    m = conf["pendiente"]
    for e in cfg.ELEMENTOS:
        col = cfg.COLUMNA_BD[e]
        o, d = v_orig[col], v_dup[col]
        valido = o.notna() & d.notna()
        b = conf["factor"] * cfg.LD_DUPLICADO[e]
        df[f"{e}_O"] = o
        df[f"{e}_D"] = d
        df[f"{e}_MIN"] = np.fmin(o, d).where(valido)
        df[f"{e}_MAX"] = np.fmax(o, d).where(valido)
        df[f"{e}_LIM"] = np.sqrt(m**2 * df[f"{e}_MIN"]**2 + b**2)
        df[f"{e}_ERR"] = _bandera(df[f"{e}_MAX"] > df[f"{e}_LIM"], valido)
        df[f"{e}_PROM"] = (o + d) / 2
        suma = (o + d).replace(0, np.nan)
        df[f"{e}_RD"] = 200 * (o - d) / suma
    return df


def curva_hiperbolica(conf, elemento, xmax, n=200):
    """Funcion hiperbolica y = sqrt(m^2 x^2 + b^2) entre 0 y xmax."""
    b = conf["factor"] * cfg.LD_DUPLICADO[elemento]
    x = np.linspace(0, xmax if np.isfinite(xmax) and xmax > 0 else 1, n)
    return x, np.sqrt(conf["pendiente"]**2 * x**2 + b**2)


# ------------------------------------------------------------ estandares
def procesar_estandar(bd_total, conf):
    ids = _ids_control(bd_total, conf["control"])
    cols = [cfg.COLUMNA_BD[e] for e in cfg.ELEMENTOS_STD]
    propio = _buscar(bd_total, ids, cols + ["LAB_REPORT", "MES "])
    df = pd.DataFrame({"SAMPLE_ID": ids, "LAB_REPORT": propio["LAB_REPORT"],
                       "MES": propio["MES "]})
    for e in cfg.ELEMENTOS_STD:
        v = propio[cfg.COLUMNA_BD[e]]
        mh, de = conf["MH"][e], conf["DE"][e]
        df[e] = v
        df[f"{e}_OUT"] = _bandera((v > mh + 3 * de) | (v < mh - 3 * de), v.notna())
    return df


def sesgo_estandar(df, conf, elemento):
    bv = conf["BV"].get(elemento)
    if bv is None or df[elemento].count() == 0:
        return np.nan
    return (df[elemento].mean() / bv - 1) * 100


# ------------------------------------------------------------- resultados
class ResultadosQAQC:
    """Agrupa todo lo calculado a partir de la BD unida."""

    def __init__(self, bd_total):
        self.bd_total = bd_total
        self.blancos = {k: procesar_blanco(bd_total, c) for k, c in cfg.BLANCOS.items()}
        self.duplicados = {k: procesar_duplicado(bd_total, c)
                           for k, c in cfg.DUPLICADOS.items()}
        self.estandares = {k: procesar_estandar(bd_total, c)
                           for k, c in cfg.ESTANDARES.items()}
        self.reportes = generar_reportes(self)

    def conteo_controles(self):
        return self.bd_total["CONTROLES"].value_counts()


def reporte_blanco(df):
    filas = []
    for e in cfg.ELEMENTOS:
        col = f"{e}_ERR"
        filas.append({"Analito": cfg.ANALITO[e],
                      "N° Muestras": int(df[col].count()),
                      "Fallos": int(df[col].sum()),
                      "%Fallos": tasa(df, col)})
    rep = pd.DataFrame(filas)
    rep["Clase"] = _clase_umbral(rep["%Fallos"], cfg.UMBRAL_FALLOS_BLANCO)
    return _redondear(rep)


def reporte_duplicados(duplicados):
    rep = pd.DataFrame({"Analito": [cfg.ANALITO[e] for e in cfg.ELEMENTOS]})
    for clave, df in duplicados.items():
        cols = [f"{e}_ERR" for e in cfg.ELEMENTOS]
        rep[f"N° Muestras {clave}"] = [int(df[c].count()) for c in cols]
        rep[f"%Fallos {clave}"] = [tasa(df, c) for c in cols]
        rep[f"Clase {clave}"] = _clase_umbral(rep[f"%Fallos {clave}"],
                                              cfg.UMBRAL_FALLOS_DUPLICADO)
    return _redondear(rep)


def reporte_estandares(estandares, por_analito=False):
    filas = []
    for clave, conf in cfg.ESTANDARES.items():
        df = estandares[clave]
        for e in cfg.ELEMENTOS_STD:
            mh, de, bv = conf["MH"][e], conf["DE"][e], conf["BV"].get(e)
            sesgo = sesgo_estandar(df, conf, e)
            filas.append({
                "Nombre\nSTD": conf["codigo"],
                "Analito": cfg.ANALITO[e],
                "N° Muestras": int(df[e].count()),
                "Media\nActual\n(MA)": df[e].mean(),
                "Media\nHistorica\n(MH)": mh,
                "Media\nCertificada\n(BV)": np.nan if bv is None else bv,
                "Lim. de\nControl\nMH-2STD": mh - 2 * de,
                "Lim. de\nControl\nMH+2STD": mh + 2 * de,
                "Sesgo(%)": sesgo,
                "Clase": clasificar_sesgo(sesgo, bv is not None),
                "# Fallidos": int(df[f"{e}_OUT"].sum()),
                "% Fallidos": tasa(df, f"{e}_OUT"),
                "_orden": cfg.ELEMENTOS_STD.index(e),
            })
    rep = pd.DataFrame(filas)
    if por_analito:
        rep = rep.sort_values("_orden", kind="stable")
        columnas = list(rep.columns)
        columnas[0], columnas[1] = columnas[1], columnas[0]
        rep = rep[columnas]
    return _redondear(rep.drop(columns="_orden").reset_index(drop=True))


def lista_errores(df, sufijo, elementos, columna_id="SAMPLE_ID"):
    """Una fila por muestra y elemento fallado: [id, LAB_REPORT, OBS]."""
    partes = []
    for e in elementos:
        fallo = df[df[f"{e}_{sufijo}"] == 1]
        partes.append(pd.DataFrame({columna_id: fallo[columna_id],
                                    "LAB_REPORT": fallo["LAB_REPORT"],
                                    "OBS": cfg.ETIQUETA[e]}))
    if not partes:
        return pd.DataFrame(columns=[columna_id, "LAB_REPORT", "OBS"])
    errores = pd.concat(partes, ignore_index=True)
    return errores.sort_values(columna_id, kind="stable").reset_index(drop=True)


def generar_reportes(res):
    rep = {
        "Duplicado": reporte_duplicados(res.duplicados),
        "Estandares": reporte_estandares(res.estandares),
        "Estandares por Analito": reporte_estandares(res.estandares, por_analito=True),
        "Blanco Fino": reporte_blanco(res.blancos["BF"]),
        "Blanco Grueso": reporte_blanco(res.blancos["BG"]),
    }
    errores = {
        "ErrorRechazoGrueso": lista_errores(res.duplicados["RG"], "ERR",
                                            cfg.ELEMENTOS, "SAMPLE_ORIG"),
        "ErrorDuplicadoPulpa": lista_errores(res.duplicados["DP"], "ERR",
                                             cfg.ELEMENTOS, "SAMPLE_ORIG"),
        "ErrorDuplicadoCampo": lista_errores(res.duplicados["DC"], "ERR",
                                             cfg.ELEMENTOS, "SAMPLE_ORIG"),
        "ErrorBlancoGrueso": lista_errores(res.blancos["BG"], "ERR", cfg.ELEMENTOS),
        "ErrorBlancoFino": lista_errores(res.blancos["BF"], "ERR", cfg.ELEMENTOS),
        "ErrorCHNB": lista_errores(res.estandares["B"], "OUT", cfg.ELEMENTOS_STD),
        "ErrorCHNM": lista_errores(res.estandares["M"], "OUT", cfg.ELEMENTOS_STD),
        "ErrorCHNA": lista_errores(res.estandares["A"], "OUT", cfg.ELEMENTOS_STD),
    }
    total = [e.rename(columns={"SAMPLE_ORIG": "SAMPLE_ID"})
             .assign(**{"Ubicacion Error": nombre})
             for nombre, e in errores.items()]
    rep.update(errores)
    rep["TotalError"] = pd.concat(total, ignore_index=True)
    return rep


def exportar_excel(hojas, ruta):
    with pd.ExcelWriter(ruta, engine="xlsxwriter") as writer:
        for nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre[:31], index=False)


# ----------------------------------------------------- control externo
def leer_reporte_externo(ruta):
    bd = pd.read_excel(ruta, header=9)
    for origen, destino in cfg.COMBINAR_EXT.items():
        if origen in bd.columns:
            bd[destino] = (bd[origen].combine_first(bd[destino])
                           if destino in bd.columns else bd[origen])
            bd = bd.drop(columns=[origen])
    bd = bd.drop(columns=[c for c in cfg.DESCARTAR_EXT if c in bd.columns])
    bd.columns = bd.iloc[0]
    return bd.iloc[4:]


def procesar_externo(bd_cert, bd_total, vmin, vmax):
    bd = bd_cert.rename(columns=cfg.RENOMBRAR_EXT).reset_index(drop=True)
    columnas_ct = [v[0] for v in cfg.EXTERNO.values()]
    _requiere(bd, ["SAMPLE_ID", *columnas_ct], "el reporte del laboratorio externo")
    bd = bd.loc[:, ~bd.columns.duplicated()]
    bd = _ids_enteros(bd, "el reporte del laboratorio externo")
    bd = bd[(bd["SAMPLE_ID"] >= vmin) & (bd["SAMPLE_ID"] <= vmax)]
    for col in columnas_ct + (["CuCi_ct"] if "CuCi_ct" in bd else []):
        bd[col] = a_numero(bd[col], cfg.REEMPLAZO_LD_EXT)
    internas = [c for c in cfg.EXTERNO if c in bd_total.columns]
    rma = bd.merge(bd_total[internas + ["SAMPLE_ID"]].drop_duplicates("SAMPLE_ID"),
                   on="SAMPLE_ID", how="left")
    if rma.empty:
        raise ValueError("Ninguna muestra del laboratorio externo esta en el rango seleccionado.")
    return rma.sort_values("SAMPLE_ID").reset_index(drop=True)


def pares_externo(bd_rma, col_sgs):
    """Pares validos SGS (X) vs Certimin (Y) de un analito."""
    col_ct = cfg.EXTERNO[col_sgs][0]
    if col_sgs not in bd_rma.columns:
        return pd.DataFrame(columns=["SAMPLE_ID", "X", "Y", "PROM", "RD", "ER"])
    p = (bd_rma[["SAMPLE_ID", col_sgs, col_ct]].dropna()
         .rename(columns={col_sgs: "X", col_ct: "Y"}).reset_index(drop=True))
    suma = (p["X"] + p["Y"]).replace(0, np.nan)
    p["PROM"] = suma / 2
    p["RD"] = 200 * (p["X"] - p["Y"]) / suma
    p["ER"] = 2 * (p["Y"] - p["X"]) / suma
    return p


def sesgo_rma(x, y):
    """Sesgo = (1 - pendiente RMA) * 100, pendiente = DE(Certimin) / DE(SGS)."""
    if len(x) < 2 or np.std(x) == 0:
        return np.nan
    return (1 - np.std(y) / np.std(x)) * 100


def recta_rma(x, y):
    """Pendiente e intercepto de la regresion de eje mayor reducido (RMA)."""
    if len(x) < 2 or np.std(x) == 0:
        return np.nan, np.nan
    signo = np.sign(np.corrcoef(x, y)[0, 1]) if np.std(y) > 0 else 1.0
    pendiente = signo * np.std(y) / np.std(x)
    return pendiente, np.mean(y) - pendiente * np.mean(x)


def reporte_externo(bd_rma):
    filas = []
    for col_sgs, (_, etiqueta, unidad) in cfg.EXTERNO.items():
        if col_sgs == "FE":
            continue
        p = pares_externo(bd_rma, col_sgs)
        n = len(p)
        outliers = int((p["ER"].abs() > cfg.UMBRAL_ER_EXT).sum())
        prom_sgs, prom_ct = p["X"].mean(), p["Y"].mean()
        sesgo = sesgo_rma(p["X"], p["Y"])
        pct = outliers * 100 / n if n else np.nan
        dentro = (abs(prom_ct - prom_sgs) <= cfg.TOLERANCIA_MEDIA_EXT * abs(prom_sgs)
                  if n else False)
        filas.append({
            "Analito": f"{etiqueta} {'gr/tn' if unidad == 'g/t' else unidad}",
            "RMA\n# Pares": n,
            "RMA\nSesgo (%)": sesgo,
            "RMA\nClase": clasificar_sesgo(sesgo),
            "Promedio\nSGS-MCP": prom_sgs,
            "Promedio\nCertimin": prom_ct,
            "Certimin\ndentro 5%\nSGS": ("ACEPTABLE" if dentro else
                                        "INACEPTABLE" if n else "Sin datos"),
            "Outlier (ER>0.1)\n#": outliers,
            "Outlier (ER>0.1)\nPorcentaje %": pct,
            "Outlier (ER>0.1)\nClase": _clase_umbral(pd.Series([pct]),
                                                     cfg.UMBRAL_OUTLIER_EXT)[0],
        })
    rep = pd.DataFrame(filas)
    clases = rep[["RMA\nClase", "Certimin\ndentro 5%\nSGS", "Outlier (ER>0.1)\nClase"]]
    aceptables = (clases == "ACEPTABLE").sum(axis=1)
    inaceptables = (clases == "INACEPTABLE").sum(axis=1)
    rep["Clase\nGeneral"] = np.where((inaceptables < 2) & (aceptables >= 2),
                                     "ACEPTABLE", "INACEPTABLE")
    return _redondear(rep)


# --------------------------------------------------- ratio de insercion
def ratio_insercion(bd_total, n_externo, anio, mes):
    conteo = bd_total["CONTROLES"].value_counts()

    def n(codigo):
        return int(conteo.get(codigo, 0))

    n_std = sum(n(c["control"]) for c in cfg.ESTANDARES.values())
    tipos = [("Dup de Campo", n(cfg.CTRL_DUP_CAMPO)),
             ("Rechazo Grueso", n(cfg.CTRL_RECHAZO)),
             ("Dup de Pulpa", n(cfg.CTRL_PULPA)),
             ("STD", n_std),
             ("Control Externo", int(n_externo)),
             ("Blanco Fino", n(cfg.CTRL_BLANCO_FINO)),
             ("Blanco Grueso", n(cfg.CTRL_BLANCO_GRUESO))]
    tipos.append(("Total Controles", sum(v for _, v in tipos)))
    total = int(bd_total["SAMPLE_ID"].count())

    filas = [["Total Muestras", total, " "], ["Tipo de Control", "N° Muestras", "RI (%)"]]
    for nombre, valor in tipos:
        filas.append([nombre, valor, round(valor * 100 / total, 2) if total else 0.0])
    return pd.DataFrame(filas, columns=["Año/Mes", str(anio), str(mes)])
