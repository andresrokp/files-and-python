
import openpyxl

print('Hello concatener')
print('-\n---\n-----\n-------\n-----\n---\n-')
print('Hello concatener\n')

dataFileName = 'lubeDataFile.xlsx'
inputSheetName = 'PREVENTIVO_LUBRICACIÓN'
outSheetName = 'AMxEQ'

inputSheetBeginRow = 12
inputSheetEndRow = 22
outSheetBeginRow = 4

dataFile = openpyxl.load_workbook(dataFileName)
originSheet = dataFile[inputSheetName]
outSheet = dataFile[outSheetName]

for row in range(inputSheetBeginRow,inputSheetEndRow + 1):
    
    # getting values from origin file
    sistema = str(originSheet[f'C{row}'].value)
    equipo = str(originSheet[f'E{row}'].value)
    lubricante = str(originSheet[f'G{row}'].value)
    cantPorUnid = str(int(originSheet[f'H{row}'].value))
    cantTotal = str(int(originSheet[f'I{row}'].value))
    cantTotalFactor = type(cantTotal) # cantTotal.split('*')[1]
    tarea = str(originSheet[f'J{row}'].value)
    freqSemanas = str(int(originSheet[f'K{row}'].value))
    tiempoMinu = str(int(originSheet[f'N{row}'].value))
    estadoMaq = str(originSheet[f'R{row}'].value)
    tipoLube = str(originSheet[f'S{row}'].value)
    
    print(f'{row} :: {sistema}\t{equipo}\t{lubricante}\t{cantPorUnid}\t{cantTotal}\t{tarea}\t{freqSemanas}\t{tiempoMinu}\t{estadoMaq}\t{tipoLube}')
    
    reqOper = 'En Operación' if estadoMaq == 'FUNCIONANDO' else 'Parado por Mantenimiento'
    unidad = 'gr' if tipoLube == 'GRASA' else 'lt'
    # numPuntos = cantTotal/cantPorUnid
    tipoHta = 'Engrasadora' if tipoLube == 'GRASA' else 'Oil safe (Recipiente)'

    # writing at outSheet
    outRow = row - 8;
    outSheet[f'B{outRow}'] = tarea
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
    

