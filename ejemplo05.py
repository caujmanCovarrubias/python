# Dados los sueldos de una serie de empleados de una empresa,
# determinar cuántos ganan más de Bs. 6000, entre Bs. 3000 y
# Bs. 6000 y menos de Bs. 3000. mas el promedio obtenido en
# cada rango. Debe distribuir en columnas por cada rango.

import openpyxl

wb = openpyxl.load_workbook(filename='ejemplos.xlsx')
hoja = wb['practica3']
tabla = hoja['H6':'K25']

ran03 = 0
ran06 = 0
ran10 = 0
ran03tot = 0
ran06tot = 0
ran10tot = 0
for fila in tabla:
    sue = fila[0].value
    if sue <= 3000:
        ran03 += 1
        ran03tot += sue
        fila[1].value = sue
    elif sue <= 6000:
        ran06 += 1
        ran06tot += sue
        fila[2].value = sue
    elif sue <= 10000:
        ran10 += 1
        ran10tot += sue
        fila[3].value = sue
hoja['I26'].value = ran03tot // ran03
hoja['J26'].value = ran06tot // ran06
hoja['K26'].value = ran10tot // ran10

wb.save('ejemplos.xlsx')
