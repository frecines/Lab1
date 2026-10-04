# -*- coding: utf-8 -*-
"""
Parametros del modulo QAQC.

Todos los valores numericos (limites de deteccion, medias historicas,
valores certificados, desviaciones, pendientes y factores) estaban repetidos
dentro de qaqc.py. Aqui se definen una sola vez: para actualizar un estandar
o un limite basta con editar este archivo.

@author: frecines
"""
from datetime import datetime

# Fecha a partir de la cual la aplicacion deja de abrir (control de licencia
# del autor). Poner en None para desactivar la verificacion.
FECHA_LIMITE = datetime(2026, 12, 1)

# Elementos analizados en la BD de leyes (columna -> etiqueta y unidad).
# La BD original usa 'ARS' para el arsenico; se renombra a 'AS' al procesar.
ELEMENTOS = ["CU", "AG", "MO", "AS", "ZN", "FE"]
ETIQUETA = {"CU": "Cu", "AG": "Ag", "MO": "Mo", "AS": "As", "ZN": "Zn", "FE": "Fe"}
ANALITO = {"CU": "Cu %", "AG": "Ag gr/tn", "MO": "Mo %", "AS": "As %",
           "ZN": "Zn %", "FE": "Fe %"}
COLUMNA_BD = {"CU": "CU", "AG": "AG", "MO": "MO", "AS": "ARS", "ZN": "ZN", "FE": "FE"}

# Valores reportados como texto por el laboratorio interno -> valor numerico.
# Cualquier otro "<x" se reemplaza por x/2 y cualquier otro ">x" por x.
REEMPLAZO_LD = {"<0.001": 0.0005, "<0.3": 0.15, "<0.0015": 0.00075,
                "<0.002": 0.001, ">500": 501}

# Valores reportados como texto por el laboratorio externo (Certimin).
REEMPLAZO_LD_EXT = {"<0.001": 0.0005, "<0.5": 0.25, "<0.0005": 0.00025,
                    "<0.003": 0.0015, "<0.002": 0.001, ">2.5": 2.6,
                    ">2.500": 2.6, ">300": 301, ">1": 1.1, ">10": 10.1}

# Codigos de la columna CONTROLES del reporte de muestras QAQC
CTRL_BLANCO_FINO = "BLK FINO"
CTRL_BLANCO_GRUESO = "BLANCO"
CTRL_DUP_CAMPO = "CAMPO ORIGIN"
CTRL_RECHAZO = "RECHA"
CTRL_PULPA = "PULPA"

# ---------------------------------------------------------------- Blancos
# Limite de deteccion por elemento para cada tipo de blanco
LD_BLANCO_FINO = {"CU": 0.001, "AG": 0.3, "MO": 0.0015, "AS": 0.002,
                  "ZN": 0.0016667, "FE": 0.495}
LD_BLANCO_GRUESO = {"CU": 0.001, "AG": 0.3, "MO": 0.0015, "AS": 0.002,
                    "ZN": 0.001, "FE": 0.127}

BLANCOS = {
    "BF": dict(nombre="Blanco Fino", control=CTRL_BLANCO_FINO,
               desfase_prec=1, factor=3, ld=LD_BLANCO_FINO),
    "BG": dict(nombre="Blanco Grueso", control=CTRL_BLANCO_GRUESO,
               desfase_prec=2, factor=5, ld=LD_BLANCO_GRUESO),
}
# % de fallos a partir del cual el blanco es INACEPTABLE
UMBRAL_FALLOS_BLANCO = 5

# ------------------------------------------------------------- Duplicados
# Limite de deteccion practico usado en la funcion hiperbolica
LD_DUPLICADO = {"CU": 0.001, "AG": 0.3, "MO": 0.0015, "AS": 0.002,
                "ZN": 0.001, "FE": 0.001}

# desfase_par: posicion de la muestra pareja respecto a la fila de control.
#   DC: la fila CAMPO ORIGIN es la original y el duplicado es SAMPLE_ID + 1.
#   RG / DP: la fila de control es el duplicado y la original es SAMPLE_ID - 1.
DUPLICADOS = {
    "DC": dict(nombre="Duplicado de Campo", titulo="DUPLICADO DE CAMPO",
               control=CTRL_DUP_CAMPO, desfase_par=+1, factor=10, pendiente=1.35),
    "RG": dict(nombre="Rechazo Grueso", titulo="RECHAZO GRUESO",
               control=CTRL_RECHAZO, desfase_par=-1, factor=5, pendiente=1.22),
    "DP": dict(nombre="Duplicado de Pulpa", titulo="DUPLICADO PULPA",
               control=CTRL_PULPA, desfase_par=-1, factor=3, pendiente=1.11),
}
UMBRAL_FALLOS_DUPLICADO = 10
# Ventana de la media movil del RD-Plot (n de pares ordenados por promedio)
VENTANA_MEDIA_MOVIL_RD = 5

# ------------------------------------------------------------- Estandares
ELEMENTOS_STD = ["CU", "AG", "MO", "AS", "ZN"]
# Los parametros de cada estandar (MH, BV, DE) ya no estan aqui: se editan
# desde la aplicacion y se guardan en estandares.json (ver qaqc_estandares.py).

# Clasificacion del sesgo (%): |sesgo| <= 5 aceptable, <= 10 cuestionable
SESGO_ACEPTABLE = 5
SESGO_CUESTIONABLE = 10

# ---------------------------------------------- Control externo (Certimin)
# columna en BD interna -> (columna Certimin, etiqueta, unidad)
EXTERNO = {
    "CU": ("Cu_ct", "Cu", "%"),
    "AG": ("Ag_ct", "Ag", "g/t"),
    "MO": ("Mo_ct", "Mo", "%"),
    "ARS": ("As_ct", "As", "%"),
    "ZN": ("Zn_ct", "Zn", "%"),
    "CUOX": ("CuSS_ct", "CuSS", "%"),
}
RENOMBRAR_EXT = {"Cu": "Cu_ct", "Ag": "Ag_ct", "Mo": "Mo_ct", "As": "As_ct",
                 "Zn": "Zn_ct", "CuSolH+": "CuSS_ct", "CuScitrico": "CuCi_ct",
                 "SAMPLE": "SAMPLE_ID"}
# Columnas del reporte de Certimin que se combinan (origen -> destino)
COMBINAR_EXT = {"G0095MCR1": "G0172MCR1", "G0040MCR1": "G0125MCR1"}
DESCARTAR_EXT = ["G0465MCR1", "G0781MCR1", "G0732MCR1"]
UMBRAL_ER_EXT = 0.1          # |error relativo| > 0.1 se considera outlier
UMBRAL_OUTLIER_EXT = 10      # % de outliers a partir del cual es INACEPTABLE
TOLERANCIA_MEDIA_EXT = 0.05  # medias de Certimin dentro del 5% de SGS
LAB_INTERNO = "CHINALCO SGS"
LAB_EXTERNO = "CERTIMIN"
