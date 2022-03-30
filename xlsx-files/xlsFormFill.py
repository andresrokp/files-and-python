import openpyxl
import shutil

print('-\n---\n-----\n-------\n-----\n---\n-')

originFileName = 'fromThis.xlsx'
originFile = openpyxl.load_workbook(originFileName)
allSheetNames = originFile.sheetnames

aSheetName = 'Sheet1'
aSheet = originFile[aSheetName]
aCell = aSheet['E5']
cellContent = aCell.value

print('\n---------------------------\n')
print(originFile)
print('\n---------------------------\n')
print(allSheetNames)
print('\n---------------------------\n')
print(aSheet)
print('\n---------------------------\n')
print(aCell)
print('\n---------------------------\n')
print(cellContent)


