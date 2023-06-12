import openpyxl
import json

wb = openpyxl.load_workbook('test.xlsx')
name_worksheet = input("Wprowadź nazwę arkusza który przeliczyć: ")
ws = wb[name_worksheet]
count_shoes_in_pl = {}
for i in range(1, ws.max_row+1):
    if isinstance(ws[f'C{i}'].value, float):
        if count_shoes_in_pl.get(ws[f'A{i}'].value):
            count_shoes_in_pl[ws[f'A{i}'].value] += int(ws[f'C{i}'].value)
        else:
            count_shoes_in_pl[ws[f'A{i}'].value] = int(ws[f'C{i}'].value)

with open("amount.json", "w") as f:
    json.dump(count_shoes_in_pl, f, indent=4)

