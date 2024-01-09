import math
import openpyxl
import numpy as np

def Xlsx_Letter(number):
    letter = {1:'A',2:'B',3:'C',4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J',11:'K',12:'L',13:'M',14:'N',15:'O',16:'P',17:'Q',18:'R',19:'S',20:'T',21:'U',22:'V',23:'W',24:'X',25:'Y',26:'Z'}
    a = ''
    jnumber = number

    if number>702:
        number = number - 702
        c = number/676 + 1
        if number%676 == 0:
            c = math.floor(c)
            c = c - 1
        else:
            c = math.floor(c)
        a = a + letter[c]
        if c>1:
            number = number - 676*(c-1)
    if jnumber>26:
        b = number/26
        if number%26 == 0:
            b = math.floor(b)
            b = b - 1
        else:
            b = math.floor(b)
        if jnumber>702:
            b = b + 1
            if b == 27:
                b = 1
                a = a + letter[b]
                b = 26
            else:
                a = a + letter[b]
                b = b - 1
        else:
            a = a + letter[b]
        number = number - 26*b
    if number == 0:
        number = 26
    a = a + letter[number]
    return a


def DataDisplay(workbook,date,data,x,y):
    a = 0
    for elem in date:
        workbook[Xlsx_Letter(x) + str(y + a)] = elem
        a = a + 1
    a = 0
    for elem in data:
        workbook[Xlsx_Letter(x + 1) + str(y + a)] = elem
        a = a + 1
    return workbook


def SlidPapild(workbook,x,y,number):
    sum = 0
    for i in range(1,number+1):
        sum = sum + float(workbook[Xlsx_Letter(x) + str(y + i - 1)].value)
    a = sum/number
    return a


def VidSlid(workbook,date,data,x,y):
    DataDisplay(workbook,date,data,x-2,y)
    slid_skaits = len(data)/2
    slid_skaits = math.floor(slid_skaits)
    for j in range (1, slid_skaits+ 1):
        for i in range (1, len(data) - j + 2):
            workbook[Xlsx_Letter(x + j - 1) + str(y + i + j - 1)] = SlidPapild(workbook,x-1,y+i-1,j)
    return workbook


def TendencesMetode(workbook,date,data,x,y):
    DataDisplay(workbook,date,data,x-2,y)
    floatdata = []
    x_ass = []
    workbook[Xlsx_Letter(x) + str(y-1)] = "t"
    for i in range(1,len(data)+1):
        floatdata.append(float(data[i-1]))
        x_ass.append(i)
        workbook[Xlsx_Letter(x) + str(y+i-1)] = i
    workbook[Xlsx_Letter(x) + str(y+i)] = i+1

    xAss = np.array(x_ass)
    yAss = np.array(floatdata)
    x_ass.append(i+1)

    z = np.polyfit(xAss, yAss, 2)

    for elem in x_ass:
        workbook[Xlsx_Letter(x+1) + str(y+elem-1)] = z[0]*pow(elem,2)+z[1]*elem+z[2]

    return workbook





wb = openpyxl.Workbook()
ws=wb.active
date = [1,2,22,31,23,12,3,123,1,23,2]
data = [14672.7, 19109.2, 21328.5, 18644.3, 14368.0, 19676.6, 19674.0, 21986.7, 16754.7, 20894.9, 24650.0, 19302.5, 18189.1, 19610.1, 23948.8, 19199.8, 16038.9, 18053.3]
TendencesMetode(ws,date,data,3,2)

#wb.save("Test.xlsx")
wb.close()