import openpyxl
import math

with open("data.csv", "r" , encoding="utf8") as f:
    next(f)
    next(f)
    row1 = next(f)
    exdate  = row1.rstrip().split(";")
    row2 = next(f)
    data  = row2.rstrip().split(";")

index = 0

for elem in data:
    try:
        value = float(elem)
        break
    except ValueError:
        index = index + 1

del exdate[0:index]
del data[0:index]

date = []

for elem in exdate:
    a = elem.replace('"','')
    date.append(a)

wb = openpyxl.Workbook()
ws=wb.active
ws['A1'] = 56
ws['B2'] = 44
sum = ws['B2'].value + ws['A1'].value
print(sum)

wb.close()