# para cada cliente determine el tiempo de registro en meses a la fecha
# para cada cliente determine el saldo en minutos considerando si es del
# sexo masculino y su saldo es menor a 80 bolivianos, pagando por minuto
# 1.5 bolivianos
# para cada cliente determine si es del sexo femenino y es casado con una
# cuenta prepago o pospago

# determine la cantidad de clientes registrados por sexo
# determine el porcentaje de clientes por tipo de plan
# distribuya en columnas los clientes por saldo en tres rangos hasta 40,
# hasta 80 hasta 120 bolivianos indicando los totales en cada rango

import datetime
import openpyxl

wb = openpyxl.load_workbook(filename='tigo.xlsx')
hoja = wb['examen']
tabla = hoja['A5':'V104']

for fila in tabla:
    diferencia = datetime.datetime.now()-fila[8].value
    fila[11].value = diferencia.days // 30

for fila in tabla:
    if fila[5].value == 'M' and fila[9].value < 80:
        fila[12].value = fila[9].value // 1.5

for fila in tabla:
    if fila[5].value == 'F' and fila[6].value == 'C' and (fila[10].value == "prepago" or fila[10].value == "pospago"):
        fila[13].value == 'cumple'

hom = 0
muj = 0
for fila in tabla:
    if fila[5].value == 'M':
        hom += 1
    elif fila[5].value == "F":
        muj += 1
    fila[14].value = fila[5].value
hoja['O105'].value = hom
hoja['O106'].value = muj

pre = 0
pos = 0
cor = 0
for fila in tabla:
    if fila[10].value == 'prepago':
        pre += 1
    elif fila[10].value == 'pospago':
        pos += 1
    elif fila[10].value == 'corporativo':
        cor += 1
    fila[15].value = fila[10].value
tot = pre + pos + cor
hoja['P105'].value = pre / tot
hoja['P105'].number_format = '0%'
hoja['P106'].value = pos / tot
hoja['P106'].number_format = '0%'
hoja['P107'].value = cor / tot
hoja['P107'].number_format = '0%'

sum40 = 0
sum80 = 0
sum120 = 0
for fila in tabla:
    if fila[9].value <= 40:
        sum40 += int(fila[9].value)
        fila[16].value = fila[9].value
    elif fila[9].value <= 80:
        sum80 += int(fila[9].value)
        fila[17].value = fila[9].value
    elif fila[9].value <= 120:
        sum120 += int(fila[9].value)
        fila[18].value = fila[9].value
hoja['Q105'].value = sum40
hoja['R105'].value = sum80
hoja['S105'].value = sum120


wb.save('tigo.xlsx')
