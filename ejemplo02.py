# Determinar el monto que se debe pagar por el préstamo de un video,
# el precio es de 4 Bs. por el préstamo. A partir del número de días
# prestados se debe determinar la multa si corresponde, si es mayor
# a 2 días se le cobra una multa de 2 Bs. por día y si es mayor a 5
# se le cobra 1 Bs. por día.

import openpyxl

wb = openpyxl.load_workbook(filename='ejemplos.xlsx')
hoja = wb['practica2']
tabla = hoja['F6':'G25']

for fila in tabla:
    prestamo = 4
    dias = int(fila[0].value)
    if dias > 4:
        prestamo += 1 * (dias - 2)
    elif dias > 2:
        prestamo += 2 * (dias - 2)
    fila[1].value = prestamo

wb.save('ejemplos.xlsx')
