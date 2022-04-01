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

# originFile.save(filename=outFileName)
