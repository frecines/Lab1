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

## Parámetros de estándares (sin modificar el código)

Los valores de cada estándar (BV, MH y DE por elemento) ya no están en el código: se
guardan en **`estandares.json`**, junto al `.exe` (si esa carpeta no permite escribir,
en `%APPDATA%\ModuloQAQC`). Se administran desde **Herramientas → Parámetros de
estándares** (o el botón en la pestaña *Estándares*):

1. La ventana abre en **solo lectura**. *Desbloquear edición* pide la contraseña
   (inicial: `qaqc`; **cambiarla en el primer uso** con *Cambiar contraseña*).
2. **Importar desde Excel**: acepta el mismo formato de `Estandares_2026.xlsx`
   (Estándar | Elemento | Valor Esperado BV: Media | Valor Histórico: Media, Std. Dev.).
   Muestra qué valores cambian antes de aplicarlos. Columnas opcionales: `Código` y
   `Descripción`.
3. También se puede editar a mano, **agregar** un estándar nuevo (p. ej. `STD-19`, tal
   como aparece en la columna CONTROLES) o **eliminar** uno.
4. **Guardar** pide confirmación, valida que MH y DE estén completos (BV vacío =
   "No Certificado") y registra cada cambio en **`historial_estandares.csv`**
   (fecha, usuario de Windows, estándar, elemento, campo, valor anterior y nuevo).
   Si ya había datos procesados, se recalculan automáticamente.
5. *Exportar a Excel* guarda los valores vigentes en el mismo formato (respaldo).

La pestaña *Estándares* muestra solo los estándares que tienen muestras en el rango
procesado; los controles `STD-xx` sin parámetros se avisan en el registro de actividad.

> La contraseña evita cambios accidentales, pero no es una protección fuerte: quien
> tenga acceso a la carpeta puede editar `estandares.json`. Para control real, dejar
> la carpeta del programa en una ubicación de red con permisos de escritura solo para
> el responsable de QAQC; los demás usuarios podrán usar la aplicación pero no guardar.

## Informe mensual en Word

Después de **Procesar Datos** (y, si corresponde, procesar el control externo en la
pestaña RMA), el botón **Generar Informe Word...** crea
`Informe_de_Control_de_Calidad_<Mes><Año>.docx` en la carpeta de salida, con la misma
estructura del informe de enero 2025:

- Memo (A, De, Fecha, Asunto, Anexos). "A" y "De" se recuerdan para el mes siguiente.
- Tabla 1 (lotes / ratio de inserción), Tabla 2 (duplicados), Tablas 3-4 (blancos),
  Tabla 5 (estándares), Tabla 6 (control externo), límites prácticos de detección y
  MRCs utilizados (con su CV): tablas de Word reales, con los mismos colores de la app.
- Los textos con números se llenan solos: cantidades, tasas de inserción, rango de
  tasas de error ("entre 0% para Mo y 12.74% para Zn"), elementos inaceptables,
  clasificación de sesgos y RMA, y un borrador de conclusiones.
- **Lo que requiere interpretación queda resaltado en amarillo** `[Interpretación]`
  para completarlo en Word (y luego quitar el resaltado).
- Análisis del acumulado (gráficos históricos por mes de los estándares) y
  **Anexo A** con todos los gráficos (opcionales). Unos 70-90 gráficos tardan ~15 s.

## Estándares con CV > 5%

Cada estándar tiene, por elemento, el **CV (%) del valor esperado BV** (columna nueva en
*Parámetros de estándares*; se importa del Excel de estándares, columna "CV (%)").
Los elementos con **CV > 5%** (Wood, 2021) **no se grafican**: ni en la ventana del
estándar, ni en el histórico, ni en el Anexo del informe. El registro de actividad indica
cuáles se omitieron. Con los valores de `Estandares_2026.xlsx` se omite **Ag** en
STD-16, STD-17 y STD-18. El umbral está en `UMBRAL_CV_ESTANDAR` (`qaqc_config.py`).
Los estándares STD-13/14/15 no tienen CV cargado, por lo que se grafican todos sus
elementos hasta que se ingrese su CV.

## Prueba en Spyder (antes de compilar)

1. **Instalar dependencias** (una sola vez). Abrir *Anaconda Prompt*, activar el entorno
   que usa Spyder y, en la carpeta de los archivos:
   ```
   cd C:\ruta\a\ModuloQAQC
   pip install -r requirements.txt
   ```
2. Abrir `qaqc.py` en Spyder. Todos los `.py`, `qaqcbs.ui` y `ui_qaqcbs.py` deben estar
   en la misma carpeta.
3. **Ejecutar > Configuración por archivo** (Ctrl+F6) → marcar
   **"Ejecutar en una terminal del sistema externa"** → Aceptar. Las aplicaciones PyQt
   ejecutadas dentro de la consola IPython de Spyder suelen colgarse o cerrar el kernel
   en la segunda ejecución; en una terminal externa se comportan como el `.exe`.
   (Alternativa: ejecutar en la consola y usar *Consolas > Reiniciar kernel* antes de cada
   ejecución.)
4. **F5** para ejecutar. Probar: BD Leyes → Reporte QAQC → Procesar Datos → gráficos de
   cada pestaña → *Parámetros de estándares* → *Generar Informe Word*.
5. Opcional, pruebas automáticas (en Anaconda Prompt): `pip install pytest` y
   `python -m pytest tests`.

Al ejecutar desde Spyder, `estandares.json` y `historial_estandares.csv` se crean en la
carpeta del código; el `.exe` usa los de su propia carpeta (se pueden copiar).

## Ejecutable (PyInstaller)

Desde *Anaconda Prompt* (con el mismo entorno de Spyder), en la carpeta de los archivos,
ejecutar **`build_exe.bat`**: instala dependencias, recompila la interfaz y ejecuta
`pyinstaller qaqc.spec`. Resultado: **`dist\Modulo QAQC.exe`** con el icono `Grid.ico`.

- `qaqc.spec` reemplaza al spec anterior (mismo nombre, icono y modo de un solo archivo).
  Con `UN_SOLO_ARCHIVO = False` se genera una carpeta `dist\Modulo QAQC\` que abre más
  rápido (no descomprime ~100 MB en cada inicio); en ese caso se distribuye la carpeta.
- UPX está desactivado: comprimir las DLL de Qt con UPX suele dañarlas y provoca falsos
  positivos de antivirus.
- Si se modifica `qaqcbs.ui` en Qt Designer, `build_exe.bat` regenera `ui_qaqcbs.py`
  (o manualmente: `python -m PyQt5.uic.pyuic qaqcbs.ui -o ui_qaqcbs.py`).
- Los errores técnicos (el .exe no tiene consola) se guardan en
  `%USERPROFILE%\QAQC_errores.log`.

### Optimizaciones para el ejecutable

| | Antes | Ahora |
|---|---|---|
| Lectura del Reporte QAQC (19 MB, 83.000 filas) | 12,8 s, ventana "No responde" | ~3 s (lector `calamine`, sólo 6 columnas), en segundo plano con barra de progreso |
| Importaciones al iniciar | 1,3 s | 0,3 s (matplotlib se carga al abrir el primer gráfico; interfaz precompilada) |
| `.ui` | leído en tiempo de ejecución desde la carpeta actual | compilado a `ui_qaqcbs.py` (no depende de archivos externos) |
| Errores sin consola | `traceback.print_exc()` con `sys.stderr = None` hacía fallar el .exe | se escriben en `QAQC_errores.log` |
| Dependencias | incluía `statsmodels` (+ scipy, ~100 MB) | eliminadas; `qaqc.spec` excluye tkinter, scipy, QtWebEngine, QML, etc. |

Compilación y ejecución verificadas en Linux con PyInstaller 6 y los archivos reales
(`BD_20241225.xlsx` y `00- REPORTE DE MUESTRAS QAQC-2024.xlsx`); el .exe de Windows
debe compilarse en Windows.

## Estructura

| Archivo | Contenido |
|---|---|
| `qaqc.py` | Interfaz (punto de entrada) |
| `qaqcbs.ui` | Diseño de la ventana (Qt Designer) |
| `qaqc_calc.py` | Lectura, unión, cálculos y reportes (sin Qt, se puede usar desde scripts) |
| `qaqc_plots.py` | Una función por tipo de gráfico |
| `qaqc_config.py` | Parámetros fijos: límites de detección, pendientes, umbrales de clasificación y fecha límite |
| `qaqc_estandares.py` | Parámetros de estándares (`estandares.json`), importación/exportación Excel e historial |
| `qaqc_dialogo_estandares.py` | Ventana *Parámetros de estándares* |
| `qaqc_informe.py` | Informe mensual en Word (python-docx) |
| `Grid.ico` | Icono del ejecutable |
| `ui_qaqcbs.py` | Interfaz compilada desde `qaqcbs.ui` (generada, no editar) |
| `qaqc.spec`, `build_exe.bat` | Compilación con PyInstaller |
| `tests/` | Pruebas con datos sintéticos (`python -m pytest tests`) |

Los estándares se actualizan desde la aplicación (ver arriba); los límites de detección y
umbrales, en `qaqc_config.py`.

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
  ahora están una sola vez (`qaqc_config.py` y `estandares.json`).
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

### Verificación con los archivos reales

Con `BD_20241225.xlsx` y el Reporte QAQC 2024, los % de fallos de blancos, duplicados y
estándares y los sesgos coinciden **exactamente** con la versión original (0 diferencias).
La única diferencia es la lista de errores de Duplicado de Campo: la versión original
mostraba 1 muestra (613354, que es de Pulpa) en lugar de las 15 muestras que realmente
fallan en Cu.

Los filtros de Malla y Reporte muestran sólo los valores con muestras en la BD de Leyes
cargada (59 mallas y 146 reportes en vez de 785 y 1.150) y permiten buscar escribiendo.

Los controles `STD-10`, `STD-11` y `STD-12` del reporte 2024 no tienen parámetros cargados,
por lo que no se evalúan (el registro de actividad lo avisa); se pueden agregar desde
*Parámetros de estándares*.

### Nota

`FECHA_LIMITE` en `qaqc_config.py` (1 de diciembre de 2026) mantiene el bloqueo por fecha
de la versión original; poner `None` para desactivarlo.
