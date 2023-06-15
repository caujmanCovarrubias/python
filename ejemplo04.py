# Se tiene registrados los sexos y los estados civiles de varias
# personas, se pide con esta informacion calcular del total de
# hombres que porcentaje son solteros, casados o divorciados
# debe colocar en una columna los estados civiles de los hombres

import openpyxl

wb = openpyxl.load_workbook(filename='ejemplos.xlsx')
hoja = wb['practica3']
tabla = hoja['B6':'D25']

sol = 0
cas = 0
div = 0
for fila in tabla:
    sex = fila[0].value
    estciv = fila[1].value
    if sex == "M":
        if estciv == 'S':
            sol += 1
        elif estciv == 'C':
            cas += 1
        elif estciv == 'D':
            div += 1
        fila[2].value = estciv
tot = sol + cas + div
hoja['D26'].value = sol / tot
hoja['D26'].number_format = '0%'
hoja['D27'].value = cas / tot
hoja['D27'].number_format = '0%'
hoja['D28'].value = div / tot
hoja['D28'].number_format = '0%'

wb.save('ejemplos.xlsx')
