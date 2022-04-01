import openpyxl

print('-\n---\n-----\n-------\n-----\n---\n-')

originFileName = 'del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx'
originFile = openpyxl.load_workbook(originFileName)
originSheet = originFile['Solicitudes']
outSheetName = 'MNN-REG-014'

for row in range(2118,2132 + 1):
    
    # getting values from origin file
    otDate = str(originSheet[f'B{row}'].value).split(' ')[0]
    otEng = str(originSheet[f'C{row}'].value)
    otType = str(originSheet[f'D{row}'].value)
    otDetail = str(originSheet[f'E{row}'].value)
    otNumber = str(int(originSheet[f'J{row}'].value))
    otPlace = str(originSheet[f'K{row}'].value)
    otDevice = str(originSheet[f'L{row}'].value)

    print(f'{otDate}\n{otEng}\n{otType}\n{otDetail}\n{otNumber}\n{otPlace}\n{otDevice}')
    print('-')

    # out file creatinon and loading
    outFileName = f'OT {otNumber}.xlsx'
    outFile = openpyxl.load_workbook('OT 0000.xlsx')
    outSheet = outFile[outSheetName]

    # output workbook writing
    outSheet['G6'] = otNumber
    outSheet['G7'] = "IVÁN DE ALBA" if otEng == 'MEC' else 'JORGE MOLINA'
    outSheet['H10'] = otDate
    outSheet['T10'] = otDate
    if otType == 'PREV': outSheet['N14'] = 'x'
    if otType == 'CORR': outSheet['AB14'] = 'x'
    if otType == 'MEJ': outSheet['AI14'] = 'x'
    outSheet['F16'] = otPlace
    outSheet['F17'] = otDevice
    outSheet['A19'] = otDetail
    
    outFile.save(filename=outFileName)

