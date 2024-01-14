import openpyxl
import math
from funkcijas import *
import numpy as np

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

ws['A1'] = "Slidosas vidējas metode"
VidSlid(ws,date,data,3,3)
ws[Xlsx_Letter(6+len(data)//2) + "1"] = "Vidēja kvadratiska klūda"
KvadKlud(ws,date,data, 3,3,3+len(data)//2,3+len(data))

#len(data) + 8

ws['A' + str(len(data) + 8)] = "Eksponenciālas noguldīsānas metode"
EksponNoglud(ws,date,data,3,len(data) + 10)
ws[Xlsx_Letter(6+9) + str(len(data) + 8)] = "Vidēja kvadratiska klūda"
KvadKlud(ws,date,data, 3,len(data) + 10,3+9,len(data) + 10+len(data))

#2*len(data) + 15

ws['A' + str(2*len(data) + 15)] = "Tendences metode"
TendencesMetode(ws,date,data,3,2*len(data) + 17)
ws[Xlsx_Letter(6+11) + str(2*len(data) + 15)] = "Vidēja kvadratiska klūda"
KvadKlud(ws,date,data, 4,2*len(data) + 17,3+11,2*len(data) + 17+len(data))

#3*len(data) + 22

number = 4
first_sezon = 1

ws['A' + str(3*len(data) + 22)] = "Sezonālais modelis"
SezonMet(ws,date,data,number,first_sezon,3,3*len(data) + 24)
ws[Xlsx_Letter(6+5) + str(3*len(data) + 22)] = "Vidēja kvadratiska klūda"
KvadKlud(ws,date,data, 7,3*len(data) + 24,8,3*len(data) + 24+len(data))



wb.save("Test.xlsx")
wb.close()