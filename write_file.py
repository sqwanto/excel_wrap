import os
import openpyxl
from openpyxl import Workbook

destwb = "C:\\Users\\Christopher.Tully\\Documents\\Excel Wrap Scripts\\Test"

wb = Workbook()
ws = wb.active
ws.title = "QA Scores"

test_list = ['Mark', '1', 'Nick', '2']
i = 0

while (i < 4):
    for rowNum in range (1, 3):
        for columnNum in range (1, 3):
            ws_cell = ws.cell(row = rowNum, column = columnNum)
            ws_cell.value = test_list[i]
            i += 1

os.chdir(destwb)
wb.save("Wrap Up.xlsx")
