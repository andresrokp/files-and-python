import openpyxl

print('-\n---\n-----\n-------\n-----\n---\n-')

originFileName = 'del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx'
originFile = openpyxl.load_workbook(originFileName)
originSheet = originFile['Solicitudes']
outSheetName = 'MNN-REG-014'

for row in range(2024,2083 + 1):
    
    # getting values from origin file
    otSolNum = str(originSheet[f'A{row}'].value)
    date = originSheet[f'B{row}'].value
    otDate = f'{date.day}/{date.month}/{date.year}'
    otEng = str(originSheet[f'I{row}'].value)
    otType = str(originSheet[f'J{row}'].value)
    otDetail = str(originSheet[f'E{row}'].value)
    otNumber = str(int(originSheet[f'L{row}'].value))
    otPlace = str(originSheet[f'G{row}'].value)
    otDevice = str(originSheet[f'H{row}'].value)

    print(f'{otDate}\n{otEng}\n{otType}\n{otDetail}\n{otNumber}\n{otPlace}\n{otDevice}')
    print('-')

    # out file creatinon and loading
    outFileName = f'OT {otNumber}.xlsx'
    outFile = openpyxl.load_workbook('OT 0000.xlsx')
    outSheet = outFile[outSheetName]

    # output workbook writing
    outSheet['G6'] = otNumber
    outSheet['R6'] = otSolNum
    outSheet['G7'] = "IVÁN DE ALBA" if otEng == 'MECANICA' else 'JORGE MOLINA'
    outSheet['H10'] = otDate
    outSheet['T10'] = otDate
    if otType == 'PREVENTIVO': outSheet['N14'] = 'x'
    if otType == 'CORRECTIVO': outSheet['AB14'] = 'x'
    if otType == 'MEJORA': outSheet['AI14'] = 'x'
    outSheet['F16'] = otPlace
    outSheet['F17'] = otDevice
    outSheet['A19'] = otDetail
    
    outFile.save(filename=outFileName)

