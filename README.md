# Módulo QAQC

Aplicación de escritorio (PyQt5) que genera gráficos y reportes QAQC (blancos,
duplicados, estándares, control externo y ratio de inserción) a partir de la
BD de leyes y del reporte de muestras QAQC.

## Uso

```bash
pip install -r requirements.txt
python qaqc.py
```

1. **Importación**: seleccionar la BD de Leyes (hoja `General`) y el Reporte de
   Muestras QAQC (hoja `LEYES SGS`), revisar el rango de `SAMPLE_ID` y los filtros
   de Malla / Reporte y presionar **Procesar Datos**. Con eso se llenan los cuadros de
   todas las pestañas. **Exportar Reporte Excel** genera `Reporte_QAQC.xlsx`.
2. En cada pestaña, cada botón abre **una ventana con una pestaña por elemento**
   (zoom, desplazamiento y guardado desde la barra de herramientas; al pasar el mouse
   sobre un punto se muestra su `SAMPLE_ID`).
3. Los PNG (300 dpi) y los Excel se guardan en la **carpeta de salida**
   (por defecto `Resultados_QAQC`, junto a la BD de Leyes). Guardar los PNG es opcional.

## Estructura

| Archivo | Contenido |
|---|---|
| `qaqc.py` | Interfaz (punto de entrada) |
| `qaqcbs.ui` | Diseño de la ventana (Qt Designer) |
| `qaqc_calc.py` | Lectura, unión, cálculos y reportes (sin Qt, se puede usar desde scripts) |
| `qaqc_plots.py` | Una función por tipo de gráfico |
| `qaqc_config.py` | **Parámetros**: límites de detección, medias históricas, valores certificados, desviaciones, pendientes, umbrales de clasificación y fecha límite |
| `tests/` | Pruebas con datos sintéticos (`python -m pytest tests`) |

Para actualizar un estándar o un límite basta con editar `qaqc_config.py`.

## Cambios respecto a la versión anterior

### Errores corregidos

- **Lista de errores de Duplicado de Campo**: se generaba con los datos de
  *Duplicado de Pulpa* (`BD_PULPAF`) en vez de `BD_OrigF`, tanto en la tabla como en
  el Excel.
- **QQ-Plot RMA de Mo**: graficaba los cuantiles de **Ag**.
- **Gráficos RMA borraban datos**: `RMAPlot` hacía `self.BD_RMA = self.BD_RMA.dropna(...)`
  elemento por elemento, por lo que después de verlo los QQ/RD, el reporte y el ratio de
  inserción trabajaban con menos muestras.
- **Línea "RMA Trend"**: era una regresión por mínimos cuadrados (OLS); ahora es la
  recta RMA (eje mayor reducido), consistente con el sesgo informado
  `(1 − DE(Certimin)/DE(SGS))·100`, que además se calcula sólo con pares completos.
- **Tooltips de Duplicados** (DC/RG/DP): mostraban el `SAMPLE_ID` de otra fila (se usaba el
  índice del subconjunto "error / sin error" sobre la tabla completa).
- **RD-Plot**: los puntos se dibujaban en `x = media móvil del promedio` (desplazados de su
  posición real) y la "media móvil" era una línea que unía todos los puntos. Ahora: puntos
  en (promedio, RD) y línea = media móvil del RD (`VENTANA_MEDIA_MOVIL_RD`).
- **Signo del RD** de Rechazo Grueso era (duplicado − original) mientras DC y DP usaban
  (original − duplicado); ahora todos usan (original − duplicado).
- **Clasificación de sesgo**: un sesgo de exactamente ±5 % (o ±10 %) quedaba como
  INACEPTABLE. Además se usaba el valor centinela `99` para "no certificado", que se
  confundía con un sesgo real de 99 %.
- **Muestras sin par / sin ley**: se contaban como "aprobadas" en el % de fallos
  (y en duplicados, si faltaba una de las dos leyes, MIN = MAX = la que existía). Ahora el %
  se calcula sólo con pares válidos y `N° Muestras` es la cantidad de pares válidos.
- **Importación RMA**: si existía la columna `G0040MCR1` se intentaba borrar `G0095MCR1`
  (ya borrada) → error. También fallaba si no existía la columna destino.
- **Filtros de Malla / Reporte**: comparaban texto contra números (p. ej. `"100"` vs `100`)
  y el filtro se ignoraba sin avisar. Ahora se eligen de una lista con los valores del archivo.
- **Valores con `<` o `>` no listados** (p. ej. `<0.01`) quedaban como texto y rompían
  los cálculos; ahora `<x` → x/2 y `>x` → x.
- **El programa no arrancaba fuera de Spyder**: faltaba `app.exec_()` y el `.ui` se buscaba
  en la carpeta actual en vez de junto al script.
- **Botón "Histórico"** estaba habilitado antes de procesar los datos (error al presionarlo).
- **Ratio de inserción** sólo se habilitaba después de exportar el reporte RMA, y con
  Año y Mes iguales (o vacíos) se perdía una columna del cuadro. Ahora se habilita al
  procesar (el control externo cuenta 0 si no se procesó) y valida Año/Mes.
- Regresión de la muestra precedente usaba `model.params[1]` (posicional, obsoleto en
  pandas 2 y eliminado en pandas 3) y fallaba con valores vacíos.
- `QQ-Plot` con algún valor vacío devolvía todos los cuantiles `NaN`.
- Se eliminaron los `print` y los CSV de depuración (`RMA2.csv`, `RMA3.csv`,
  `RMAentrada.csv`, `RMAsalida.csv`) que se escribían en cada ejecución.

### Optimización

- De ~5.800 líneas en un solo archivo a ~1.430 en cuatro módulos: cada gráfico
  estaba copiado 6 veces (una por elemento) y cada grupo de gráficos 3 veces
  (DC/RG/DP, CHNB/CHNM/CHNA). Ahora hay una función por tipo de gráfico.
- Las constantes (medias, BV, DE, límites) estaban duplicadas en `NewCol` y `REPORT`;
  ahora están una sola vez en `qaqc_config.py`.
- Se eliminó la dependencia de `statsmodels`.

### Interfaz

- Ventana redimensionable (layouts en vez de posiciones fijas). Antes la tabla
  "Reporte por Estándar" estaba fuera del área visible.
- Flujo en 3 pasos: *Archivos → Rango y filtros → Procesar*. Los botones
  "Selección Rango", "Unión de BD", "Generación de Columnas" y "Generar Reporte" se
  reemplazaron por **Procesar Datos** + **Exportar Reporte Excel**.
- Filtros de Malla y Reporte como listas desplegables; rango con validación numérica.
- Registro de actividad con hora y errores en rojo; cursor de espera durante los procesos;
  los errores se muestran en un mensaje sin cerrar la aplicación.
- Gráficos agrupados en una ventana con pestañas, con barra de herramientas de matplotlib.
- Carpeta de salida configurable y opción para no guardar los PNG (más rápido).
- Ambos reportes de estándares (por analito y por estándar) en sub-pestañas.

### Nota

`FECHA_LIMITE` en `qaqc_config.py` (1 de diciembre de 2026) mantiene el bloqueo por fecha
de la versión original; poner `None` para desactivarlo.
