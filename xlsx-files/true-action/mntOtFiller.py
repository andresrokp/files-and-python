import shutil
import openpyxl

print('-\n---\n-----\n-------\n-----\n---\n-')

originFileName = 'del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx'
originFile = openpyxl.load_workbook(originFileName)
originSheet = originFile['Solicitudes']
outSheetName = 'MNN-REG-014'

for row in range(2118,2120 + 1):
    
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
    # shutil.copyfile('OT 0000.xlsx',outFileName)
    outFile = openpyxl.load_workbook('OT 0000.xlsx')
    outSheet = outFile[outSheetName]
    outSheet['G6'] = otNumber
    outFile.save(filename=outFileName)

