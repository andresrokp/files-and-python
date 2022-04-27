
import openpyxl

print('Hello concatener')
print('-\n---\n-----\n-------\n-----\n---\n-')
print('Hello concatener\n')

dataFileName = 'lubeDataFile.xlsx'
inputSheetName = 'PREVENTIVO_LUBRICACIÓN'
outSheetName = 'AMxEQ'

inputSheetBeginRow = 12
inputSheetEndRow = 14
outSheetBeginRow = 4

originFile = openpyxl.load_workbook(dataFileName)
originSheet = originFile[inputSheetName]

for row in range(inputSheetBeginRow,inputSheetEndRow + 1):
    
    # getting values from origin file
    sistema = str(originSheet[f'C{row}'].value)
    equipo = str(originSheet[f'E{row}'].value)
    cantPorUnid = str(int(originSheet[f'H{row}'].value))
    freqSemanas = str(int(originSheet[f'K{row}'].value))
    
    
    print(f'{sistema}\t{equipo}\t{cantPorUnid}\t{freqSemanas}')
    print('-')

    # out file creatinon and loading
    # outFileName = f'OT {otNumber}.xlsx'
    # outFile = openpyxl.load_workbook('OT 0000.xlsx')
    # outSheet = outFile[outSheetName]

    # output workbook writing
    # outSheet['R6'] = otSolNum
    # outSheet['G7'] = "IVÁN DE ALBA" if otEng == 'MECANICA' else 'JORGE MOLINA'
    # if otType == 'PREVENTIVO': outSheet['N14'] = 'x'
    
    # outFile.save(filename=outFileName)
    

