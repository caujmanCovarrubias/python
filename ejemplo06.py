# Dadas las fechas de las ventas realizadas por un vendedor
# en la empresa, distribuya en columnas de acuerdo al trimestre
# sea 1ro o 2do, identificar que fecha en la menor del primer
# trimestre y cual es la mayor del segundo trimestre

import datetime
import openpyxl

wb = openpyxl.load_workbook(filename='ejemplos.xlsx')
hoja = wb['practica3']
tabla = hoja['O6':'Q25']

min1 = tabla[0][0].value
max2 = tabla[0][0].value
for fila in tabla:
    fecha = fila[0].value
    if fecha.month in {1, 2, 3}:
        if fecha < min1:
            min1 = fecha
        fila[1].value = fecha.strftime('%d/%m/%Y')
    if fecha.month in {4, 5, 6}:
        if fecha > max2:
            max2 = fecha
        fila[2].value = fecha.strftime('%d/%m/%Y')
hoja['P26'].value = min1.strftime('%d/%m/%Y')
hoja['Q26'].value = max2.strftime('%d/%m/%Y')

wb.save('ejemplos.xlsx')
