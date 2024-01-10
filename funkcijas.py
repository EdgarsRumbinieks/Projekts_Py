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
    workbook[Xlsx_Letter(x + 1) + str(y - 1)]= "y"
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
    sum = 0
    for i in range (1,10):
        workbook[Xlsx_Letter(x+i) + str(y-1)] = "poly(" + str(i) + ")"
        z = np.polyfit(xAss, yAss, i)
        for elem in x_ass:
            for j in range(0,i+1):
                sum = sum + z[j]*pow(elem,i-j)
            workbook[Xlsx_Letter(x+i) + str(y+elem-1)] = sum
            sum = 0
        z = np.polyfit(np.log(xAss), yAss, 1)
        workbook[Xlsx_Letter(x+i+1) + str(y-1)] = "log"
        for elem in x_ass:
            workbook[Xlsx_Letter(x+i+1) + str(y+elem-1)] = z[0]*np.log(elem) + z[1]
    return workbook


def EksponNoglud(workbook,date,data,x,y):
    DataDisplay(workbook,date,data,x-2,y)
    for i in range(1,len(data)):
        for j in range(1,10):
            workbook[Xlsx_Letter(x+j-1) + str(y-1)] = str(j/10)
            workbook[Xlsx_Letter(x+j-1) + str(y+i+1)] = float(workbook[Xlsx_Letter(x-1) + str(y+i)].value)*(j/10)+(1-j/10)*float(workbook[Xlsx_Letter(x-1) + str(y+i-1)].value)
    return workbook


def SezonMet(workbook,date,data,number,first_sezon,x,y):
    DataDisplay(workbook,date,data,x-2,y)
    workbook[Xlsx_Letter(x) + str(y-1)] = "linear"
    floatdata = []
    x_ass = []
    for i in range(1,len(data)+1):
        floatdata.append(float(data[i-1]))
        x_ass.append(i)
    xAss = np.array(x_ass)
    yAss = np.array(floatdata)
    x_ass.append(i+1)
    z = np.polyfit(xAss, yAss, 1)
    for elem in x_ass:
            workbook[Xlsx_Letter(x) + str(y+elem-1)] = z[0]*elem + z[1]
    workbook[Xlsx_Letter(x+1) + str(y-1)] = "sezona"
    a = first_sezon-1
    for i in range(1,len(data)+2):
        a = a+1
        if a<=number:
            workbook[Xlsx_Letter(x+1) + str(y+i-1)] = a
        else:
            a = 1
            workbook[Xlsx_Letter(x+1) + str(y+i-1)] = a
    workbook[Xlsx_Letter(x+2) + str(y-1)] = "i"
    for i in range(1,len(data)+1):
        workbook[Xlsx_Letter(x+2) + str(y+i-1)] = float(workbook[Xlsx_Letter(x-1) + str(y+i-1)].value)/float(workbook[Xlsx_Letter(x) + str(y+i-1)].value)
    workbook[Xlsx_Letter(x+3) + str(y-1)] = "i sez"
    sum = 0
    a = 0
    for i in range(1,len(data)+2):
        curent_sez = int(workbook[Xlsx_Letter(x+1) + str(y+i-1)].value)
        for j in range(1,len(data)+1):
            if int(workbook[Xlsx_Letter(x+1) + str(y+j-1)].value) == curent_sez:
                sum = sum + float(workbook[Xlsx_Letter(x+2) + str(y+j-1)].value)
                a = a+1
        workbook[Xlsx_Letter(x+3) + str(y+i-1)] = sum/a
        sum = 0
        a = 0
    workbook[Xlsx_Letter(x+4) + str(y-1)] = "y~"
    for i in range(1,len(data)+2):
        workbook[Xlsx_Letter(x+4) + str(y+i-1)] = float(workbook[Xlsx_Letter(x) + str(y+i-1)].value)*float(workbook[Xlsx_Letter(x+3) + str(y+i-1)].value)
    return workbook




wb = openpyxl.Workbook()
ws=wb.active
date = [1,2,22,31,23,12,3,123,1,23,2]
data = [14672.7, 19109.2, 21328.5, 18644.3, 14368.0, 19676.6, 19674.0, 21986.7, 16754.7, 20894.9, 24650.0, 19302.5, 18189.1, 19610.1, 23948.8, 19199.8, 16038.9, 18053.3]
SezonMet(ws,date,data,4,1,3,2)

#wb.save("Test.xlsx")
wb.close()