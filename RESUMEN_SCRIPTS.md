# Resumen de Scripts de Automatización de Reportes

Este documento explica de forma resumida la función de cada uno de los scripts en el directorio `ReportAutomation` y cómo ejecutarlos desde la terminal.

## 1. `download_reports.py`
**Descripción:** Se encarga de conectarse al servidor y descargar los reportes de producción (como el Excel de Novedades) para una fecha y turno en especifico, basándose en la configuración de `config.json`. Si el archivo ya existe en la carpeta `descargas/reporteturno` o en su histórico (`descargas/reporteturno/_procesados`), se salta la descarga.  
**Uso/Llamado:**
```bash
python download_reports.py --fecha YYYY-MM-DD --turno 1
# Ejemplo:
python download_reports.py --fecha 2026-03-11 --turno 2
```

## 2. `batch_download.py`
**Descripción:** Realiza una descarga masiva de reportes invocando iterativamente a los scripts de descarga seleccionados.
**Uso/Llamado:**
```bash
python batch_download.py --tipo [novedades|desperdicios|novedades_impresion|todos] --fecha_inicio YYYY-MM-DD --fecha_fin YYYY-MM-DD
# Ejemplo:
python batch_download.py --tipo todos --fecha_inicio 2026-03-01 --fecha_fin 2026-03-10
# Si omites las fechas, tomará por defecto desde el 2026-03-01 hasta el día actual.
```

## 3. `process_reports.py`
**Descripción:** Procesa un archivo Excel específico que ya ha sido descargado en la carpeta `descargas/reporteturno/` leyendo las distintas hojas (como novedades e histórico) y envía los datos transformados a la base de datos local SQLite (haciendo uso de Prisma).  
**Uso/Llamado:**
```bash
python process_reports.py --fecha YYYY-MM-DD --turno 1
# Ejemplo:
python process_reports.py --fecha 2026-03-11 --turno 1
```

## 4. `process_all.py`
**Descripción:** Busca todos los archivos pendientes de procesamiento en el directorio `descargas/reporteturno/`. Extrae la fecha y turno del nombre del archivo, lo procesa con la lógica de `process_reports.py` y, tras insertarlo a la base de datos, mueve el archivo procesado a la subcarpeta `descargas/reporteturno/_procesados/`. Finalmente recalcula los "deltas" para las fechas que acaban de ser procesadas. Ideal para programar en una tarea automatizada.  
**Uso/Llamado:**
```bash
python process_all.py
```

## 5. `recalc_deltas.py`
**Descripción:** Fuerza el recálculo masivo de las discrepancias (deltas) en la tabla `Historico` por turno a partir de una fecha determinada, actualizando la base de datos SQLite en consecuencia. Es útil si se han modificado reportes anteriores manualmente y las restas de producción/desperdicio entre turnos se han desincronizado.  
**Uso/Llamado:**
```bash
python recalc_deltas.py --desde YYYY-MM-DD
# Ejemplo (por defecto inicia el 2026-02-01 si no se le envían parámetros):
python recalc_deltas.py --desde 2026-03-01
```

## 6. `download_desperdicios.py`
**Descripción:** Se encarga de descargar específicamente los reportes de "Desperdicios" para una fecha y turno seleccionados. Usa lógica idéntica a `download_reports.py` pero está enfocado a otra URL.
**Uso/Llamado:**
```bash
python download_desperdicios.py --fecha YYYY-MM-DD --turno {1,2}
```

## 7. `process_desperdicios.py`
**Descripción:** Levanta todos los excels no procesados en la carpeta `descargas/desperdicios/` y carga las métricas de la hoja 'Desperdicio Liquidados Turno' en la base de datos (tabla `Desperdicio`). Al finalizar, mueve los archivos procesados a `descargas/desperdicios/_procesados`.
**Uso/Llamado:**
```bash
python process_desperdicios.py
```

## 8. `download_novedades.py`
**Descripción:** Realiza la descarga del archivo de "Novedades de Impresión", basándose de manera parecida en fecha y turno. Los archivos se guardan en la carpeta `descargas/novedades/`.
**Uso/Llamado:**
```bash
python download_novedades.py --fecha YYYY-MM-DD --turno {1,2}
```

## 9. `process_novedades_impresion.py`
**Descripción:** Escanea los reportes XLS de la carpeta `descargas/novedades/`, extrayendo las observaciones del turno a nivel máquina y Novedades a nivel "OT". Esta información es convertida internamente desde la hoja Excel a una jerarquía y volcada en la tabla `NovedadImpresionOT` evitando el guardado en caso de existir previamente en la DB. Los excels validados se mueven a `descargas/novedades/_procesados`.
**Uso/Llamado:**
```bash
# Procesa todos los archivos en la carpeta de novedades (por defecto)
python process_novedades_impresion.py

# Alternativamente, puedes forzar un archivo puntual:
python process_novedades_impresion.py --archivo descargas/novedades/Novedades_impresion_Intranet_Turno_01_20260310.xls
```

---
**Nota:** Recuerda siempre tener activo el entorno de conda correspondiente o el ambiente virtual (venv) que contenga las librerías indicadas en tu `requirements.txt` (pandas, requests, prisma, openpyxl, xlrd, etc.) antes de ejecutar estos comandos.
