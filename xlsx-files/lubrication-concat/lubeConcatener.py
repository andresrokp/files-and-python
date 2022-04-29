
import openpyxl

print('Hello concatener')
print('-\n---\n-----\n-------\n-----\n---\n-')
print('Hello concatener\n')

dataFileName = 'lubeDataFile.xlsx'
inputSheetName = 'PREVENTIVO_LUBRICACIÓN'
outSheetName = 'AMxEQ'

inputSheetBeginRow = 14
inputSheetEndRow = 24
outSheetBeginRow = 4

dataFile = openpyxl.load_workbook(dataFileName)
originSheet = dataFile[inputSheetName]
outSheet = dataFile[outSheetName]

for row in range(inputSheetBeginRow,inputSheetEndRow + 1):
    
    # getting values from origin file
    sistema = str(originSheet[f'C{row}'].value)
    equipo = str(originSheet[f'E{row}'].value)
    posicion = str(originSheet[f'F{row}'].value)
    lubricante = str(originSheet[f'G{row}'].value)
    cantPorUnid = str(float(originSheet[f'H{row}'].value))
    cantTotal = str(float(originSheet[f'I{row}'].value))
    cantTotalFactor = type(cantTotal) # cantTotal.split('*')[1]
    tarea = str(originSheet[f'J{row}'].value)
    freqSemanas = str(int(originSheet[f'K{row}'].value))
    tiempoMinu = str(int(originSheet[f'N{row}'].value))
    estadoMaq = str(originSheet[f'R{row}'].value)
    tipoLube = str(originSheet[f'S{row}'].value)
    
    print(f'{row} :: {sistema}\t{equipo}\t{lubricante}\t{cantPorUnid}\t{cantTotal}\t{tarea}\t{freqSemanas}\t{tiempoMinu}\t{estadoMaq}\t{tipoLube}')
    
    
    unidadLube = 'gr' if tipoLube == 'GRASA' else 'lt'
    numPuntos = float(cantTotal)/float(cantPorUnid)
    tipoHta = 'Engrasadora' if tipoLube == 'GRASA' else 'Oil safe (Recipiente)'
    
    descripcion = f"""{tarea}
    Lubricante: {lubricante}
    Cantidad: {cantPorUnid} {unidadLube} (Por punto)
    # Puntos: {numPuntos}
    Herramienta: {tipoHta}
    """

    nombreEquipo = f'({sistema}) ({equipo}) - ({posicion})'
    reqOper = 'En Operación' if estadoMaq == 'FUNCIONANDO' else 'Parado por Mantenimiento'

    # writing at outSheet
    outRow = row - 10;
    outSheet[f'B{outRow}'] = tarea
    outSheet[f'C{outRow}'] = descripcion
    outSheet[f'E{outRow}'] = nombreEquipo

    print('-')

dataFile.save(filename='out-'+dataFileName)

    # out file creatinon and loading
    # outFileName = f'OT {otNumber}.xlsx'
    # outFile = openpyxl.load_workbook('OT 0000.xlsx')
    # outSheet = outFile[outSheetName]

    # output workbook writing
    # outSheet['R6'] = otSolNum
    # outSheet['G7'] = "IVÁN DE ALBA" if otEng == 'MECANICA' else 'JORGE MOLINA'
    # if otType == 'PREVENTIVO': outSheet['N14'] = 'x'
    
    # outFile.save(filename=outFileName)
    

