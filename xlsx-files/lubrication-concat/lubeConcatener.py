
# by: andresrokp
# repo: https://github.com/andresrokp/files-and-python/blob/main/xlsx-files/lubrication-concat/lubeConcatener.py
# goal: migrate an spreadsheet data from one order to another
# lic: an open crap for anyone to use :)
# gracias

import openpyxl
import os
import math

print('\nHello concatener')
print('-\n---\n-----\n-------\n-----\n---\n-')
print('Hello concatener\n')

# nombramientos
dataFileName = 'MNN-PLN-003 REV001 PLAN DE MANTENIMIENTO PREVENTIVO DE LUBRICACIÓN.xlsx'
inputSheetName = 'PREVENTIVO_LUBRICACIÓN'
outSheetName = 'AMxEQ'
outFileName = 'out-lubeDataFile.xlsx'

# declaración de filas
inputSheetBeginRow = 14
inputSheetEndRow = 207
outSheetBeginRow = 4

# traer data al workspace
dataFile = openpyxl.load_workbook(dataFileName)
originSheet = dataFile[inputSheetName]
outSheet = dataFile[outSheetName]

# iteración sobre filas / row iterator
for row in range(inputSheetBeginRow,inputSheetEndRow + 1):
    # getting values from origin file
    sistema = str(originSheet[f'C{row}'].value)
    equipo = str(originSheet[f'E{row}'].value)
    posicion = str(originSheet[f'F{row}'].value)
    lubricante = str(originSheet[f'G{row}'].value)
    cantPorUnid = float(originSheet[f'H{row}'].value)
    cantTotal = float(originSheet[f'I{row}'].value)
    tarea = str(originSheet[f'J{row}'].value)
    freqSemanas = int(originSheet[f'K{row}'].value)
    tiempoMinu = int(originSheet[f'N{row}'].value)
    estadoMaq = str(originSheet[f'R{row}'].value)
    tipoLube = str(originSheet[f'S{row}'].value)
    print(f'{row} :: {sistema}\t{equipo}\t{lubricante}\t{cantPorUnid}\t{cantTotal}\t{tarea}\t{freqSemanas}\t{tiempoMinu}\t{estadoMaq}\t{tipoLube}')
    
    # contruyendo valores de salida / assemblying output values
    unidadLube = 'gr' if tipoLube == 'GRASA' else 'lt'
    numPuntos = int(cantTotal/cantPorUnid)
    tipoHta = 'Engrasadora' if tipoLube == 'GRASA' else 'Oil safe (Recipiente)'
    descripcion = f"""{tarea}
    Lubricante: {lubricante}
    Cantidad: {cantPorUnid} {unidadLube} (Por punto)
    # Puntos: {numPuntos}
    Herramienta: {tipoHta}
    """
    nombreEquipo = f'{equipo} : {posicion} ({sistema}) (BOPP)'
    reqOper = 'En Operación' if estadoMaq == 'FUNCIONANDO' else 'Parado por Mantenimiento'
    freqDias = freqSemanas * 7
    hardcodedList = ['Media','Lubricacion','Sistemática','Flotante','Tiempo']
    holgura = 0;
    if freqDias <= 60:
        holgura = math.ceil(0.2*freqDias)
    elif freqDias <= 120:
        holgura = math.ceil(0.15*freqDias)
    else:
        holgura = math.ceil(0.10*freqDias)
    # writing at outSheet
    outRow = row - 10;
    outSheet[f'B{outRow}'] = tarea
    outSheet[f'C{outRow}'] = descripcion
    outSheet[f'E{outRow}'] = nombreEquipo
    outSheet[f'H{outRow}'] = hardcodedList[0]
    outSheet[f'I{outRow}'] = hardcodedList[1]
    outSheet[f'J{outRow}'] = hardcodedList[2]
    outSheet[f'K{outRow}'] = hardcodedList[3]
    outSheet[f'L{outRow}'] = hardcodedList[4]
    outSheet[f'M{outRow}'] = reqOper
    outSheet[f'N{outRow}'] = freqDias
    outSheet[f'O{outRow}'] = holgura
    outSheet[f'R{outRow}'] = tiempoMinu
    print('-')

# deleting and creating out file
if os.path.exists(outFileName) : os.remove(outFileName)
dataFile.save(filename=outFileName)
