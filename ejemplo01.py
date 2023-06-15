# Para un numero indicar si tiene 1, 2, 3 o más dígitos.
# (Considere solo los números positivos)

import openpyxl

wb = openpyxl.load_workbook(filename='ejemplos.xlsx')
hoja = wb['practica2']
tabla = hoja['B6':'C25']

for fila in tabla:
    numero = int(fila[0].value)
    if numero > 999:
        digitos = "4 o mas"
    elif numero > 99:
        digitos = "3 digitos"
    elif numero > 9:
        digitos = "2 digitos"
    elif numero > -1:
        digitos = "1 digito"
    fila[1].value = digitos

wb.save('ejemplos.xlsx')
