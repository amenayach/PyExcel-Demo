import xlrd
from xlutils.copy import copy

# read
originalWorkbook = xlrd.open_workbook('b.xlsx')

originalSheet0 = originalWorkbook.sheet_by_index(0)

print('Value at (0,0): {0}'.format(originalSheet0.cell(0,0).value))

# write only support xls extension
wb = copy(originalWorkbook)

sheet = wb.get_sheet(0)

sheet.write(0,0, 29834)

wb.save('b.xls')