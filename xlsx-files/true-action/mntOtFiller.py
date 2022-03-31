import shutil
import openpyxl

print('-\n---\n-----\n-------\n-----\n---\n-')

originFileName = 'del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx'
originFile = openpyxl.load_workbook('del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx')
originSheet = originFile['Solicitudes']

for row in range(2025,2033 + 1):
    otNumber = str(int(originSheet[f'J{row}'].value))
    otDate = str(originSheet[f'B{row}'].value).split(' ')[0]
    otDetail = str(originSheet[f'E{row}'].value)
    otType = str(originSheet[f'D{row}'].value)
    originSheet[f'C{row}'] = 'hello .xlsx'
    print(f'{otNumber}\n{otDate}\n{otDetail}\n{otType}')
    print('-')

originFile.save(filename=originFileName)
