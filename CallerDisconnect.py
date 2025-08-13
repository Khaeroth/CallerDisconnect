# -*- coding: utf-8 -*-
# pyinstaller --onefile --icon CD.ico CallerDisconnect.py
from calendar import c
from math import e
import os
import sys
from numpy import long
import xlwings as xw
from xlwings.constants import LineStyle
from collections import Counter
import matplotlib.pyplot as plt
import tempfile
import numpy as np # Necesitas importar numpy para la manipulación de arrays en las barras agrupadas

try:
    archivo = sys.argv[1]       # Para abrir un archivo que se arraste al programa/acceso directo.
    #archivo = "C:\\Users\\Hugo\\source\\repos\\CallerDisconect\\Abandon_Caller Dissonnect 250729_144801.xls"
    print("Processing data...")
    #print(archivo)
    #input("Presione Enter para continuar")

except IndexError:
    print("ERROR: There's no file to process.")
    input("Press Enter to continue.")
    sys.exit()

wb = xw.Book(archivo)
try:
    hoja = wb.sheets["Sheet0"]

except:
    print("ERROR: Please check the name of the sheet with the data. The name of the sheet must be 'Sheet0'.")
    input("Press Enter to continue.")
    sys.exit()


# Verifica si la hoja 'Reporte' ya existe, sino, la crea.
if 'Report' in [s.name for s in wb.sheets]:
    hoja_graficas = wb.sheets['Report']
else:
    hoja_graficas = wb.sheets.add(name='Report')

reporte = wb.sheets["Report"]

# Vacía la hoja de reporte para evitar duplicados.
reporte.clear()
for imagen in reporte.pictures:
    imagen.delete()


# FUNCIONES ######################################################################################################################################

# Esta función convierte el número que arroja i.column de BuscaPalabras() y lo traduce a una letra del abecedario. Queda pendiente para cuando es AA, AB y así.
def Numero_A_Letra(valor):

    if valor > 26:
        valor = valor - 25
    unicode = chr(65 + valor)
    return unicode

# Esta función busca una cadena de texto entre las celdas que se le especifique, y devuelve el valor de la columna y la fila. Cabe destacar que tiene i.column-1 porque
# normalmente arroja el valor a partir de cero. Un ejemplo del resultado de esta función es: ('D', 17)
def BuscaPalabras(palabra, rango):
    
    for i in hoja.range(rango):
        if i.value == palabra:
            columna = Numero_A_Letra(i.column-1)
            return columna, i.row

def GraficaDiariaAgrupada(source_total, source_long_queue, fila, columna, titulo_principal):
    # Paso 1: Definir colores por día
    colores_por_dia = {
        'Sun': 'orange',
        'Mon': 'red',
        'Tue': 'green',
        'Wed': 'purple',
        'Thu': 'blue',
        'Fri': 'pink',
        'Sat': 'brown'
    }

    # Recorremos los días de la semana para generar una gráfica por cada día
    for dia in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']:
        lista_total = source_total.get(dia, [])
        lista_long_queue = source_long_queue.get(dia, [])
        lista_short_queue = source_long_queue.get(dia, [])

        if not lista_total and not lista_long_queue and not lista_short_queue:
            continue # Si no hay datos para el día en ninguna de las fuentes, saltar al siguiente día

        # Contar ocurrencias por hora para llamadas totales
        # Por ejemplo, queda así: Counter({'Sat, 26, 12': 4, 'Sat, 26, 09': 1, 'Sat, 26, 13': 1, 'Sat, 26, 20': 1})
        conteo_total = Counter(lista_total)
                
        # Contar ocurrencias por hora para llamadas en espera, lo mismo que arriba.
        conteo_long_queue = Counter(lista_long_queue)

        # Contar ocurrencias por hora para llamadas en espera, lo mismo que arriba.
        conteo_short_queue = Counter(lista_short_queue)

        # Obtener todas las horas únicas presentes en ambos conjuntos de datos para este día:
        # El operador | entre dos conjuntos significa unión. La unión de dos conjuntos incluye todos los elementos únicos de ambos conjuntos, es decir, se eliminan los duplicados.
        # Por ejemplo: {'Sat, 26, 12', 'Sat, 26, 09', 'Sat, 26, 13'} | {'Sat, 26, 12', 'Sat, 26, 09', 'Sat, 26, 20''} → {'Sat, 26, 12', 'Sat, 26, 09', 'Sat, 26, 13', 'Sat, 26, 20'}
        # list(...) convierte el conjunto resultante de la unión en una lista, ya que la unión con el operador | devuelve un conjunto y queremos una lista ordenada.
        # sorted() ordena los elementos de la lista de menor a mayor. En este caso, ordenará las horas alfabéticamente (como cadenas de texto).
        horas_unicas = sorted(list(set(conteo_total.keys()) | set(conteo_long_queue.keys() | set(conteo_short_queue.keys()))))
        #print(horas_unicas)

        # Preparar los valores para las barras. Esta es una lista que contiene los valores de llamadas para todas las horas en horas_unicas, usando .get() para manejar posibles claves faltantes.
        # conteo_total es un diccionario que mapea horas a la cantidad de llamadas totales.
        # hora es una variable que recorre todas las horas de horas_unicas
        # .get(hora, 0) intenta obtener el valor asociado a la hora en el diccionario conteo_total. Si la hora no se encuentra en el diccionario, devuelve 0 (especificado como segundo argumento de .get()).
        # Es decir, si la hora está presente en conteo_total, se obtiene el valor correspondiente (el número de llamadas). Si no está, se asigna 0
        valores_total = [conteo_total.get(hora, 0) for hora in horas_unicas]
        valores_long_queue = [conteo_long_queue.get(hora, 0) for hora in horas_unicas]
        valores_short_queue = [conteo_short_queue.get(hora, 0) for hora in horas_unicas]
        # Para el ciclo del lunes, print(valores_totales, valores_long_queue) da: [6, 5, 1, 8, 24, 7, 3, 12, 4, 2, 1] [1, 0, 0, 0, 5, 1, 0, 2, 1, 0, 0]
        # Para el ciclo del martes, print(valores_totales, valores_long_queue) da: [1, 1, 5, 5, 4, 5, 2, 2, 4, 6, 3, 1, 2] [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        # Para el ciclo del miércoles, print(valores_totales, valores_long_queue) da: [3, 1, 8, 3, 7, 9, 4, 5, 3, 1, 2] [0, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0]
        # Para el ciclo del jueves, print(valores_totales, valores_long_queue) da: [1, 6, 5, 2, 4, 6, 3, 5, 5, 4, 3, 2] [0, 3, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1]
        # Para el ciclo del viernes, print(valores_totales, valores_long_queue) da: [5, 2, 4, 3, 13, 6, 6, 8, 4] [1, 0, 0, 0, 1, 0, 0, 0, 1]
        # Para el ciclo del sábado, print(valores_totales, valores_long_queue) da: [1, 4, 1, 1] [0, 0, 0, 0]
        # Para el ciclo del domingo, print(valores_totales, valores_long_queue) da: [2] [1]

        # Lo anterior es lo mismo que:
        '''
        # Crear una lista para los valores de llamadas totales por hora
        valores_total = []
        for hora in horas_unicas:
            # Obtener el número de llamadas totales para esta hora, o 0 si no hay datos
            valor = conteo_total.get(hora, 0)
            valores_total.append(valor)

        # Crear una lista para los valores de llamadas en espera por hora
        valores_long_queue = []
        for hora in horas_unicas:
            # Obtener el número de llamadas en espera para esta hora, o 0 si no hay datos
            valor = conteo_long_queue.get(hora, 0)
            valores_long_queue.append(valor)
        '''
        
        # Calcular el total de llamadas para este día
        total_llamadas_dia = sum(valores_total)
        total_queue_dia = sum(valores_long_queue)
        total_short_queue_dia = sum(valores_short_queue)

        # Convertir a formato AM/PM para las etiquetas del eje X
        etiquetas_am_pm = []
        for etiqueta_completa in horas_unicas:
            # El formato es 'Dia, Numero, Hora' (ej. 'Mon, 10, 08')
            partes = etiqueta_completa.split(',')   # Para separar por las comas
            numero = partes[1].strip()              # Número de la fecha
            hora_24 = int(partes[2].strip())        # Hora como tal

            if hora_24 == 0:
                hora_12 = '12am'
            elif 1 <= hora_24 < 12:
                hora_12 = f'{hora_24}am'
            elif hora_24 == 12:
                hora_12 = '12pm'
            else:
                hora_12 = f'{hora_24 - 12}pm'
            etiquetas_am_pm.append(hora_12)

        # Crear gráfico de barras agrupadas
        plt.figure(figsize=(10, 4)) # Ajusta el tamaño de la gráfica (x, y) = (ancho, alto)
        bar_width = 0.25 # Ancho de cada barra
        # np.arange() es una función de NumPy que genera una secuencia de números enteros desde 0 hasta 'len(etiquetas_am_pm) - 1'. En este caso, genera una lista de posiciones para las barras de llamadas totales.
        r1 = np.arange(len(etiquetas_am_pm)) # Posiciones para las barras de llamadas totales        
        r2 = [x + bar_width for x in r1] # Posiciones para las barras de llamadas en espera
        r3 = [x + 2*bar_width for x in r1] # Posiciones para las barras de llamadas en espera
        #print(r1, r2)

        plt.bar(r1, valores_total, color=colores_por_dia.get(dia, 'gray'), width=bar_width, label=f'Total Calls ({total_llamadas_dia})')
        plt.bar(r2, valores_long_queue, color='darkgray', width=bar_width, label=f'<30seg Waiting Calls ({total_short_queue_dia})')
        plt.bar(r3, valores_long_queue, color='lightgray', width=bar_width, label=f'1 Min> Waiting Calls ({total_queue_dia})')

        # Mostrar valor encima de cada barra
        for i, val in enumerate(valores_total):
            plt.text(r1[i], val + 0.2, str(val), ha='center', va='bottom', fontsize=8)
        for i, val in enumerate(valores_short_queue):
            plt.text(r2[i], val + 0.2, str(val), ha='center', va='bottom', fontsize=8)
        for i, val in enumerate(valores_long_queue):
            plt.text(r3[i], val + 0.2, str(val), ha='center', va='bottom', fontsize=8)

        # Añadir líneas de referencia horizontales
        plt.grid(axis='y', linestyle='--', alpha=0.6)

        # Ponerle margen en la parte de arriba para que los números no se salgan del recuadro
        max_val = max(max(valores_total) if valores_total else 0, max(valores_long_queue) if valores_long_queue else 0)
        plt.ylim(0, max_val * 1.5) # Aumenta el tope del eje Y en 50%
        plt.tight_layout(pad=3.0) # Aumenta separación entre subcomponentes del gráfico

        # Títulos y etiquetas
        plt.xlabel('Time')
        plt.ylabel('Call Count')
        #plt.title(f'{dia} {numero} - {titulo_principal}: {total_llamadas_dia}')
        plt.title(f'{dia} {numero} - {titulo_principal}')
        plt.xticks([r + bar_width/2 for r in r1], etiquetas_am_pm, rotation=90) # Centra las etiquetas X entre las barras
        plt.legend() # Muestra la leyenda

        # Guardar imagen temporal
        img_path = os.path.join(tempfile.gettempdir(), f'{titulo_principal}_{dia}_diaria.png')
        plt.savefig(img_path)
        plt.close()

        # Insertar imagen en Excel
        celda_inicial = f'{columna}{fila}'
        reporte.pictures.add(img_path, name=f'{titulo_principal}_{dia}_diaria', update=True,
                             left=reporte.range(celda_inicial).left,
                             top=reporte.range(celda_inicial).top)

        fila += 25 # Avanza hacia abajo para la siguiente imagen (ajusta este valor si es necesario)

def GraficaSemanalAgrupada(source_total, source_long_queue, source_short_queue, fila, columna, nombre):

    colores_por_dia = {
            'Sun': 'orange',
            'Mon': 'red',
            'Tue': 'green',
            'Wed': 'purple',
            'Thu': 'blue',
            'Fri': 'pink',
            'Sat': 'brown'
        }
    etiquetas_semana = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    
    valores_total_semana = []
    valores_long_queue_semana = []
    valores_short_queue_semana = []

    # Calcular los totales por día para ambas fuentes
    for dia in etiquetas_semana:
        total_dia_total = sum(Counter(source_total.get(dia, [])).values())
        total_dia_long_queue = sum(Counter(source_long_queue.get(dia, [])).values())
        total_dia_short_queue = sum(Counter(source_short_queue.get(dia, [])).values())
        
        valores_total_semana.append(total_dia_total)
        valores_long_queue_semana.append(total_dia_long_queue)
        valores_short_queue_semana.append(total_dia_short_queue)

    # Calcular el total de llamadas de toda la semana
    total_llamadas_semanal = sum(valores_total_semana)
    total_queue_semanal = sum(valores_long_queue_semana)
    total_short_queue_semanal = sum(valores_short_queue_semana)

    # Crear gráfico de barras agrupadas
    plt.figure(figsize=(10, 4)) # Ajusta tamaño según lo necesario
    bar_width = 0.25 # Ancho de cada barra
    r1 = np.arange(len(etiquetas_semana)) # Posiciones para las barras de llamadas totales
    r2 = [x + bar_width for x in r1] # Posiciones para las barras de llamadas en espera que duraron <30seg 
    r3 = [x + 2*bar_width for x in r1] # Posiciones para las barras de llamadas en espera que duraron 1min>

    plt.bar(r1, valores_total_semana, color=['red', 'green', 'purple', 'blue', 'pink', 'brown','orange'], width=bar_width, label=f'Total Calls ({total_llamadas_semanal})')
    plt.bar(r2, valores_short_queue_semana, color='darkgrey', width=bar_width, label=f'<30seg+ Waiting Calls ({total_short_queue_semanal})')
    plt.bar(r3, valores_long_queue_semana, color='lightgray', width=bar_width, label=f'1 Min> Waiting Calls ({total_queue_semanal})')
    
    # Mostrar valor encima de cada barra
    for i, val in enumerate(valores_total_semana):
        plt.text(r1[i], val + 0.2, str(val), ha='center', va='bottom', fontsize=8)
    for i, val in enumerate(valores_short_queue_semana):
        plt.text(r2[i], val + 0.2, str(val), ha='center', va='bottom', fontsize=8)
    for i, val in enumerate(valores_long_queue_semana):
        plt.text(r3[i], val + 0.2, str(val), ha='center', va='bottom', fontsize=8)

    # Añadir líneas de referencia horizontales
    plt.grid(axis='y', linestyle='--', alpha=0.6)

    # Ponerle margen en la parte de arriba para que los números no se salgan del recuadro
    max_val = max(max(valores_total_semana) if valores_total_semana else 0, max(valores_long_queue_semana) if valores_long_queue_semana else 0)
    plt.ylim(0, max_val * 1.5) # Aumenta el tope del eje Y en 50%
    plt.tight_layout(pad=3.0) # Aumenta separación entre subcomponentes del gráfico

    # Títulos y etiquetas
    plt.xlabel('Days of the Week')
    plt.ylabel('Total Call Count')
    #plt.title(f'{nombre} - Week Overview - Total: {total_llamadas_semanal}')
    plt.title(f'{nombre} - Week Overview')
    plt.xticks([r + bar_width/2 for r in r1], etiquetas_semana, rotation=45) # Centra las etiquetas X entre las barras
    plt.legend() # Muestra la leyenda

    # Guardar imagen temporal
    img_path = os.path.join(tempfile.gettempdir(), f'{nombre}_semanal_agrupada.png')
    plt.savefig(img_path)
    plt.close()

    # Insertar imagen en Excel
    celda_inicial = f'{columna}{fila}'
    reporte.pictures.add(img_path, name=f'{nombre}_semanal_agrupada', update=True,
                         left=reporte.range(celda_inicial).left,
                         top=reporte.range(celda_inicial).top)

# OBTENCIÓN DE DATOS #############################################################################################################################
timestamp = BuscaPalabras("TIMESTAMP", "A1:D20")
llamadas = []
contador = 0
for cell in hoja.range(f"{timestamp[0]}{timestamp[1]+1}:{timestamp[0]}400"):
    if contador == 3:
        break # Usa break en lugar de continue para salir del bucle completamente
    if cell.value is None:
        contador += 1
    else:
        llamadas.append(cell.value)
        contador = 0

print(f"{len(llamadas)} calls processed.")

dias = []
for item in llamadas:
    if len(item) == 24:
        dias.append(f"{item[0:6]}, {item[-8:-6]}")
    elif len(item) == 25:
        dias.append(f"{item[0:7]}, {item[-8:-6]}")

queue_wait_time = BuscaPalabras("QUEUE WAIT TIME", "P10:R25")
contador = 0
long_queue = []
short_queue = []
for cell in hoja.range(f"{queue_wait_time[0]}{queue_wait_time[1]+1}:{queue_wait_time[0]}400"):
    if contador == 3:
        break
    if cell.value is None:
        contador += 1
    elif cell.value != None:
        contador = 0
        if int(cell.value[4]) > 0:                      #Si la llamada duró más de 1 minuto, procesarla.
            celda = hoja.range(f"{timestamp[0]}{cell.row}")
            if len(celda.value) == 24:
                long_queue.append(f"{celda.value[0:6]}, {celda.value[-8:-6]}")
            elif len(celda.value) == 25:
                long_queue.append(f"{celda.value[0:7]}, {celda.value[-8:-6]}")
        elif int(cell.value[6]) < 3:                    #Si la llamada duró menos de 30 segundos, procesarla.          
            celda = hoja.range(f"{timestamp[0]}{cell.row}")
            if len(celda.value) == 24:
                short_queue.append(f"{celda.value[0:6]}, {celda.value[-8:-6]}")
            elif len(celda.value) == 25:
                short_queue.append(f"{celda.value[0:7]}, {celda.value[-8:-6]}")

# TABLA CON DATOS ################################################################################################################################

cursor = 5

titulo = f"Total calls: {len(llamadas)}"
reporte.range(f"A{cursor-2}").value = titulo

for i in llamadas:
    celda = reporte.range(f"A{cursor}")
    celda.value = i
    for i in range(7, 11):
        celda.api.Borders(i).LineStyle = LineStyle.xlContinuous
    cursor += 1

reporte.range("A:A").autofit()

# SETUP DATOS PARA GRAFICAS ######################################################################################################################

llamadas_por_dia = {'Mon': [], 'Tue': [], 'Wed': [], 'Thu': [], 'Fri': [], 'Sat': [],'Sun': []}
long_queue_por_dia = {'Mon': [], 'Tue': [], 'Wed': [], 'Thu': [], 'Fri': [], 'Sat': [],'Sun': []}
short_queue_por_dia = {'Mon': [], 'Tue': [], 'Wed': [], 'Thu': [], 'Fri': [], 'Sat': [],'Sun': []}

for item in dias:
    partes = item.split(',')
    dia_semana = partes[0].strip()
    llamadas_por_dia[dia_semana].append(item)

for item in long_queue:
    partes = item.split(',')
    dia_semana = partes[0].strip()
    long_queue_por_dia[dia_semana].append(item)

for item in short_queue:
    partes = item.split(',')
    dia_semana = partes[0].strip()
    short_queue_por_dia[dia_semana].append(item)


### Llamar a las nuevas funciones de gráfica agrupada ###

# Ubicación inicial para pegar imágenes
fila_actual = 3

# Gráfica Semanal Agrupada
GraficaSemanalAgrupada(source_total=llamadas_por_dia, source_long_queue=long_queue_por_dia, source_short_queue=short_queue_por_dia,
                        fila=fila_actual, columna="D", nombre="Callers Disconnected")

fila_actual += 25 # Ajusta para el siguiente gráfico

# Gráfica Diaria Agrupada
GraficaDiariaAgrupada(source_total=llamadas_por_dia, source_long_queue=long_queue_por_dia, 
                        fila=fila_actual, columna="D", titulo_principal="Total Callers Disconnected")

# GUARDAR Y CERRAR EXCEL #########################################################################################################################

wb.save()
input("Press Enter to continue.")
#wb.close()