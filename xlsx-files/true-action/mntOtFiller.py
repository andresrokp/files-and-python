import shutil
import openpyxl

print('-\n---\n-----\n-------\n-----\n---\n-')

originFileName = 'del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx'
originFile = openpyxl.load_workbook('del drive - MNN-REG-046 Solicitudes de Mantenimiento.xlsx')
originSheet = originFile['Solicitudes']
otNumber = str(int(originSheet['J2025'].value))
otDate = str(originSheet['B2025'].value).split(' ')[0]
otDetail = str(originSheet['E2025'].value)
otType = str(originSheet['D2025'].value)
print(f'{otNumber}\n{otDate}\n{otDetail}\n{otType}')

