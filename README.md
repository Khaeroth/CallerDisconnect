# CallerDisconnect

Script en Python para procesar un reporte exportado de **Five9** y generar, dentro del mismo archivo Excel, una hoja de **resumen visual** con:

- Total de llamadas que se desconectaron.
- Llamadas con espera **menor a 30 segundos**.
- Llamadas con espera **mayor a 1 minuto**.
- Gráficas semanales y diarias por día/hora.

El script trabaja sobre un archivo `.xls` de entrada y escribe el resultado en una hoja llamada `Report`.

---

## ¿Cómo funciona?

`CallerDisconnect.py` hace, en orden, lo siguiente:

1. Abre el archivo Excel recibido por argumento de línea de comandos.
2. Busca la hoja de datos llamada **`Sheet0`** (requisito obligatorio).
3. Lee columnas clave por encabezado:
   - `TIMESTAMP`
   - `QUEUE WAIT TIME`
4. Detecta y clasifica llamadas:
   - **Total calls**: todos los timestamps válidos encontrados.
   - **<30 sec waiting**: esperas menores a 30 segundos.
   - **>1 min waiting**: esperas mayores a 1 minuto.
5. Crea/limpia la hoja `Report`.
6. Inserta:
   - Tabla con llamadas procesadas.
   - Gráfica semanal agrupada.
   - Gráficas diarias agrupadas por día de la semana.
7. Guarda el mismo archivo con los cambios.

---

## Requisitos

- **Windows + Microsoft Excel instalado** (el script usa `xlwings`, que controla Excel localmente).
- Python 3.x.
- Dependencias:

```bash
pip install xlwings matplotlib numpy
```

> Nota: `collections`, `os`, `sys` y `tempfile` son módulos estándar de Python.

---

## Uso

### Ejecución directa (Python)

```bash
python CallerDisconnect.py "C:\ruta\al\archivo.xls"
```

Ejemplo:

```bash
python CallerDisconnect.py "C:\Reports\Abandon_Caller Disconnect 250729_144801.xls"
```

Al final verás el mensaje para presionar Enter, y el archivo quedará guardado con la hoja `Report` actualizada.

### Generar ejecutable (.exe)

El proyecto incluye configuración para PyInstaller:

```bash
pyinstaller --onefile --icon CD.ico CallerDisconnect.py
```

También existe `CallerDisconnect.spec` para builds reproducibles.

---

## Formato esperado del archivo de entrada

Para que el procesamiento funcione correctamente:

- Debe existir una hoja llamada exactamente **`Sheet0`**.
- Deben existir los encabezados:
  - `TIMESTAMP` (se busca en rango `A1:D20`)
  - `QUEUE WAIT TIME` (se busca en rango `P10:R25`)
- El script recorre filas hasta detectar 3 celdas vacías consecutivas en cada bloque relevante.

Si no se cumple lo anterior, el script mostrará error y terminará.

---

## Salida generada

En la hoja **`Report`** del mismo archivo:

- Lista de llamadas procesadas.
- Gráfica semanal con 3 series:
  - Total Calls
  - `<30seg+ Waiting Calls`
  - `1 Min> Waiting Calls`
- Gráficas diarias (Lun-Dom) con el mismo enfoque, por hora en formato am/pm.

---

## Estructura del repositorio

- `CallerDisconnect.py`: lógica principal de procesamiento y generación de reportes.
- `CallerDisconnect.spec`: configuración de PyInstaller.
- `CallerDisconnect.pyproj` y `CallerDisconnect.sln`: archivos de proyecto/solución para Visual Studio.

---

## Troubleshooting rápido

- **"There's no file to process."**
  - Ejecuta el script pasando la ruta del archivo como argumento.

- **"Please check the name of the sheet... must be 'Sheet0'."**
  - Renombra la hoja de datos a `Sheet0`.

- **No se generan datos esperados**
  - Verifica que los encabezados `TIMESTAMP` y `QUEUE WAIT TIME` estén presentes en los rangos indicados.
