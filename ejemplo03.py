# Se procede a registrar el monto de una compra de productos, para lo
# que se debe determinar el número de boletos para el sorteo de premios
# que se le va a asignar de acuerdo al monto de la compra. Si la compra
# está entre 1 y 50 Bs. le corresponde un boleto, si la compra está
# entre 50 y 100 Bs. le corresponde tres boletos y si la compra fuera
# mayor a 100 Bs. le corresponde 5 boletos directamente. Determinar el
# número de boletos que se deben asignar.

import openpyxl

wb = openpyxl.load_workbook(filename='ejemplos.xlsx')
hoja = wb['practica2']
tabla = hoja['j6':'k25']

for fila in tabla:
    monto = int(fila[0].value)
    if monto > 100:
        boletos = 5
    elif monto > 50:
        boletos = 3
    else:
        boletos = 1
    fila[1].value = boletos

wb.save('ejemplos.xlsx')
