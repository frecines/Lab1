# -*- coding: utf-8 -*-
"""
Graficos QAQC.

Cada funcion dibuja sobre una matplotlib.figure.Figure ya creada (la ventana
de graficos le asigna un canvas Qt antes de dibujar, para que funcionen los
tooltips de mplcursors). Antes habia un bloque copiado por cada elemento;
ahora hay una funcion por tipo de grafico.

@author: frecines
"""
import numpy as np
import pandas as pd
import mplcursors

import qaqc_config as cfg
from qaqc_calc import curva_hiperbolica, recta_rma, sesgo_estandar
from qaqc_estandares import categoria_cv, es_referencial

CAJA_ERROR = dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")


# ----------------------------------------------------------------- helpers
def _tooltip(fig, artista, texto):
    """Tooltip al pasar el mouse. `texto(i)` recibe el indice del punto.
    Las figuras para el informe Word (fig.estatica = True) no los necesitan."""
    if getattr(fig, "estatica", False):
        return
    cursor = mplcursors.cursor(artista, hover=True)

    @cursor.connect("add")
    def _al_agregar(sel):
        sel.annotation.set_text(texto(int(sel.index)))
        sel.annotation.get_bbox_patch().set_alpha(0.85)

    # Mantener una referencia para que el cursor no sea recolectado
    if not hasattr(fig, "_cursores"):
        fig._cursores = []
    fig._cursores.append(cursor)


def _sin_datos(fig, titulo):
    ax = fig.add_subplot(111)
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.text(0.5, 0.5, "Sin datos para el rango / filtros seleccionados",
            ha="center", va="center", fontsize=13, color="gray",
            transform=ax.transAxes)
    ax.set_axis_off()


def _anotar(ax, texto, xy, color="red"):
    ax.annotate(texto, xy=xy, xycoords="axes fraction", fontsize=15,
                color=color, bbox=CAJA_ERROR)


def _fmt(valor, decimales=3):
    return "-" if pd.isna(valor) else f"{valor:.{decimales}f}"


def _marcar_meses(ax, d):
    """Linea vertical al inicio de cada mes y su nombre en el centro."""
    if "MES" not in d or d["MES"].isna().all():
        return
    meses = d["MES"].astype(str)
    bloque = meses.ne(meses.shift()).cumsum()
    ymin = ax.get_ylim()[0]
    for _, grupo in d.groupby(bloque, sort=False):
        ax.axvline(grupo["SAMPLE_ID"].iloc[0], color="gray", linestyle="--")
        ax.annotate(str(grupo["MES"].iloc[0]), xy=(grupo["SAMPLE_ID"].mean(), ymin),
                    xytext=(0, -10), textcoords="offset points",
                    ha="center", va="top", fontsize=10)


def _serie_temporal(fig, ax, d, col, etiqueta, **estilo):
    ax.plot(d["SAMPLE_ID"], d[col], "-", color=estilo.get("color", "gray"), zorder=2)
    puntos = ax.scatter(d["SAMPLE_ID"], d[col], label=etiqueta, zorder=3,
                        color=estilo.get("relleno", "yellow"),
                        edgecolor=estilo.get("color", "gray"))
    _tooltip(fig, puntos, lambda i: f"SAMPLE_ID: {d['SAMPLE_ID'].iloc[i]}\n"
                                    f"{etiqueta}: {_fmt(d[col].iloc[i], 4)}")


# ----------------------------------------------------------------- blancos
def grafico_blanco(fig, df, elemento, titulo, tasa_error, meses=False):
    d = df.dropna(subset=[elemento])
    if d.empty:
        return _sin_datos(fig, titulo)
    ax = fig.add_subplot(111)
    _serie_temporal(fig, ax, d, elemento, cfg.ETIQUETA[elemento])
    ax.axhline(d[f"{elemento}_MAV"].iloc[0], color="red", label="Max Accepted Value")
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.set_ylabel("Values")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=2)
    _anotar(ax, f"Error: {_fmt(tasa_error, 2)}%", (0.75, -0.25))
    if meses:
        _marcar_meses(ax, d)
    ax.set_xticklabels([])
    fig.tight_layout()


def grafico_precedente(fig, df, elemento, titulo):
    prec = f"{elemento}_P"
    d = df.dropna(subset=[elemento, prec])
    if d.empty:
        return _sin_datos(fig, titulo)
    et = cfg.ETIQUETA[elemento]
    ax = fig.add_subplot(111)
    puntos = ax.scatter(d[prec], d[elemento], color="blue", label=elemento)
    ax.axhline(d[f"{elemento}_MAV"].iloc[0], color="red", label="Max Accepted Value")
    if len(d) >= 2 and d[prec].nunique() > 1:
        pendiente, intercepto = np.polyfit(d[prec], d[elemento], 1)
        x = np.array([d[prec].min(), d[prec].max()])
        ax.plot(x, intercepto + pendiente * x, color="green", linestyle="--",
                label=f"Linear {et}")
    _tooltip(fig, puntos, lambda i: f"SAMPLE_ID: {d['SAMPLE_ID'].iloc[i]}\n"
                                    f"{et} Prec: {_fmt(d[prec].iloc[i])}\n"
                                    f"{et}: {_fmt(d[elemento].iloc[i])}")
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.set_ylabel("Ley Muestra Blanca")
    ax.set_xlabel("Ley Muestra Precedente")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=3)
    fig.tight_layout()


# -------------------------------------------------------------- duplicados
def grafico_duplicado(fig, df, elemento, conf, tasa_error):
    titulo = f"{conf['titulo']} {elemento}"
    d = df.dropna(subset=[f"{elemento}_ERR"])
    if d.empty:
        return _sin_datos(fig, titulo)
    et = cfg.ETIQUETA[elemento]
    ax = fig.add_subplot(111)
    for valor, color, marcador, nombre in [(0, "orange", "o", elemento),
                                           (1, "red", "x", f"{elemento} Error")]:
        sub = d[d[f"{elemento}_ERR"] == valor]
        if sub.empty:
            continue
        puntos = ax.scatter(sub[f"{elemento}_MIN"], sub[f"{elemento}_MAX"],
                            color=color, marker=marcador, s=20, label=nombre)
        # Cada grupo usa su propio sub-DataFrame (antes el tooltip mostraba
        # el SAMPLE_ID de otra fila).
        _tooltip(fig, puntos, lambda i, s=sub: (
            f"SAMPLE_ID: {s['SAMPLE_ORIG'].iloc[i]}\n"
            f"{et} MAX: {_fmt(s[f'{elemento}_MAX'].iloc[i])}\n"
            f"{et} MIN: {_fmt(s[f'{elemento}_MIN'].iloc[i])}"))
    xmax = d[f"{elemento}_MAX"].max()
    ax.plot([0, xmax], [0, xmax], color="black", label="Y=X")
    ax.plot(*curva_hiperbolica(conf, elemento, xmax), color="red",
            label=f"Fun. Hiper {et}")
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.set_xlabel("Valor Minimo")
    ax.set_ylabel("Valor Maximo")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3)
    _anotar(ax, f"Error: {_fmt(tasa_error, 2)}%", (0.7, 0.1))
    ax.grid(color="gray", linestyle="-", linewidth=0.5)
    fig.tight_layout()


def grafico_qq(fig, x, y, titulo, xlabel, ylabel, nombre_x, nombre_y,
               etiqueta="Q-Q Plot"):
    x, y = pd.Series(x).dropna(), pd.Series(y).dropna()
    if x.empty or y.empty:
        return _sin_datos(fig, titulo)
    percentiles = np.linspace(0, 100, 101)
    qx, qy = np.percentile(x, percentiles), np.percentile(y, percentiles)
    ax = fig.add_subplot(111)
    ax.plot(qx, qx, color="red", label="Y=X")
    puntos = ax.scatter(qx, qy, color="blue", marker="x", label=etiqueta)
    _tooltip(fig, puntos, lambda i: f"P{i}\n{nombre_x}: {_fmt(qx[i])}\n"
                                    f"{nombre_y}: {_fmt(qy[i])}")
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3)
    ax.grid(color="gray", linestyle="-", linewidth=0.5)
    fig.tight_layout()


def grafico_rd(fig, df, col_prom, col_rd, titulo, etiqueta, id_col="SAMPLE_ID"):
    """RD-Plot: diferencia relativa vs promedio del par + media movil del RD."""
    d = df.dropna(subset=[col_prom, col_rd]).sort_values(col_prom)
    if d.empty:
        return _sin_datos(fig, titulo)
    ax = fig.add_subplot(111)
    puntos = ax.scatter(d[col_prom], d[col_rd], color="blue", label=f"RD Plot {etiqueta}")
    movil = d[col_rd].rolling(cfg.VENTANA_MEDIA_MOVIL_RD, min_periods=1, center=True).mean()
    ax.plot(d[col_prom], movil, color="red", label="Media Movil")
    ax.axhline(0, color="black", linewidth=0.8)
    _tooltip(fig, puntos, lambda i: f"SAMPLE_ID: {d[id_col].iloc[i]}\n"
                                    f"{etiqueta} prom: {_fmt(d[col_prom].iloc[i])}\n"
                                    f"RD {etiqueta}: {_fmt(d[col_rd].iloc[i], 2)}")
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.set_xlabel("Promedio")
    ax.set_ylabel("% RD")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=2)
    ax.grid(color="gray", linestyle="-", linewidth=0.5)
    fig.tight_layout()


# -------------------------------------------------------------- estandares
def grafico_estandar(fig, df, elemento, conf, meses=False):
    et = cfg.ETIQUETA[elemento]
    mh, de, bv = conf["MH"][elemento], conf["DE"][elemento], conf["BV"].get(elemento)
    referencial = bv is not None and es_referencial(conf, elemento)
    titulo = f"{conf['titulo']} {et}" + (" (valor referencial)" if referencial else "")
    d = df.dropna(subset=[elemento])
    if d.empty:
        return _sin_datos(fig, titulo)
    ax = fig.add_subplot(111)
    _serie_temporal(fig, ax, d, elemento, et, color="blue", relleno="blue")
    lineas = [(mh, "orange", f"Media Historica {et}", "-")]
    if bv is not None:
        lineas.append((bv, "black", f"BV {et} (referencial)" if referencial else f"BV {et}",
                       "--" if referencial else "-"))
    lineas += [(mh + 2 * de, "brown", "MH+2DE", "-"), (mh - 2 * de, "brown", "MH-2DE", "-"),
               (mh - 3 * de, "red", "MH-3DE", "-"), (mh + 3 * de, "red", "MH+3DE", "-")]
    for valor, color, nombre, estilo in lineas:
        ax.axhline(valor, color=color, linestyle=estilo, label=nombre)
    ax.set_title(titulo, fontsize=14, pad=20)
    ax.set_ylabel("Values")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3)
    ax.grid(color="gray", linestyle="-", linewidth=0.5)
    if referencial:
        cv = conf["CV"][elemento]
        ax.text(0.5, 1.015, f"Valor referencial ({categoria_cv(cv)}): CV del BV {cv:.2f}% > "
                f"{cfg.UMBRAL_CV_ESTANDAR}%, sin evaluacion de sesgo",
                transform=ax.transAxes, ha="center", va="bottom", fontsize=10,
                color="#1F4E9A", style="italic")
    elif bv is not None:
        _anotar(ax, f"Sesgo: {_fmt(sesgo_estandar(d, conf, elemento), 2)}%", (0.8, -0.25))
    if meses:
        _marcar_meses(ax, d)
    ax.set_xticklabels([])
    fig.tight_layout()


# --------------------------------------------------------- control externo
def grafico_rma(fig, pares, etiqueta, unidad, sesgo):
    titulo = f"External Controls ({cfg.LAB_INTERNO} vs {cfg.LAB_EXTERNO})"
    if pares.empty:
        return _sin_datos(fig, f"{titulo} - {etiqueta}")
    x, y = pares["X"], pares["Y"]
    ax = fig.add_subplot(111)
    puntos = ax.scatter(x, y, color="#339966", edgecolor="black", s=30, marker="^",
                        label="Check\nSamples")
    pendiente, intercepto = recta_rma(x, y)
    xr = np.array([x.min(), x.max()])
    if np.isfinite(pendiente):
        ax.plot(xr, intercepto + pendiente * xr, color="red", label="RMA\nTrend")
    ax.plot([0, x.max()], [0, x.max()], color="#5089BC", label="Y=X")
    _tooltip(fig, puntos, lambda i: f"SAMPLE_ID: {pares['SAMPLE_ID'].iloc[i]}\n"
                                    f"{etiqueta} SGS: {_fmt(x.iloc[i])}\n"
                                    f"{etiqueta} Cert: {_fmt(y.iloc[i])}")
    ax.set_title(titulo, fontsize=14, pad=20,
                 bbox=dict(facecolor="#FFFF99", edgecolor="black", boxstyle="round,pad=0.5"))
    ax.set_xlabel(f"{cfg.LAB_INTERNO} {etiqueta} {unidad}")
    ax.set_ylabel(f"{cfg.LAB_EXTERNO} {etiqueta} {unidad}")
    ax.legend(loc="center left", bbox_to_anchor=(1, 0.5), edgecolor="black")
    sin_borde = dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
    ax.annotate("Sesgo", xy=(1.03, 0.39), xycoords="axes fraction", fontsize=15,
                fontweight="bold", bbox=sin_borde)
    ax.annotate(f"{_fmt(sesgo, 2)}%", xy=(1.03, 0.35), xycoords="axes fraction",
                fontsize=15, color="red", fontweight="bold", bbox=sin_borde)
    ax.grid(color="gray", linestyle="-", linewidth=0.5)
    fig.tight_layout()


def grafico_cuss(fig, bd_rma, incluir_cuci):
    titulo = "CuCi vs CUSS" if incluir_cuci else \
        f"CUSS {cfg.LAB_EXTERNO} vs SGS Toromocho"
    if "CUOX" not in bd_rma:
        return _sin_datos(fig, titulo)
    d = bd_rma.dropna(subset=["CUOX", "CuSS_ct"]).sort_values("SAMPLE_ID")
    if d.empty:
        return _sin_datos(fig, titulo)
    series = [("CuSS_ct", "orange", "CuSS% Certimin", "CuSS_Cert")]
    if incluir_cuci and "CuCi_ct" in d:
        series.append(("CuCi_ct", "red", "CuCi% Certimin", "CuCi_Cert"))
    series.append(("CUOX", "blue", "CuSS% SGS Toromocho", "CuSS_Toro"))
    ax = fig.add_subplot(111)
    for col, color, nombre, corto in series:
        ax.plot(d["SAMPLE_ID"], d[col], color=color, zorder=2)
        puntos = ax.scatter(d["SAMPLE_ID"], d[col], color=color, label=nombre, zorder=3)
        _tooltip(fig, puntos, lambda i, c=col, n=corto: (
            f"SAMPLE_ID: {d['SAMPLE_ID'].iloc[i]}\n{n}: {_fmt(d[c].iloc[i])}"))
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.07), shadow=True, ncol=3)
    ax.set_title(titulo)
    ax.set_xlabel("Muestras")
    ax.set_ylabel("%Cu Valores")
    ax.grid(True)
    fig.tight_layout()
