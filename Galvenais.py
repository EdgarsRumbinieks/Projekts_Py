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



#wb = openpyxl.Workbook()
#ws=wb.active
#VidSlid(ws,date,data,3,1)

#wb.save("Test.xlsx")
#wb.close()