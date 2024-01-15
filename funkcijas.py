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
    slid_data = []
    for j in range (1, slid_skaits+ 1):
        workbook[Xlsx_Letter(x+j-1) + str(y-1)] = "s = " + str(j)
        for i in range (1, len(data) - j + 2):
            workbook[Xlsx_Letter(x + j - 1) + str(y + i + j - 1)] = SlidPapild(workbook,x-1,y+i-1,j)
            if i == len(data) - j + 1:
                slid_data.append(SlidPapild(workbook,x-1,y+i-1,j))
    return workbook, slid_data


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
    trend_data = []
    for i in range (1,10):
        workbook[Xlsx_Letter(x+i) + str(y-1)] = "poly(" + str(i) + ")"
        z = np.polyfit(xAss, yAss, i)
        for elem in x_ass:
            for j in range(0,i+1):
                sum = sum + z[j]*pow(elem,i-j)
            workbook[Xlsx_Letter(x+i) + str(y+elem-1)] = sum
            if elem == x_ass[-1]:
                trend_data.append(sum)
            sum = 0
        z = np.polyfit(np.log(xAss), yAss, 1)
        workbook[Xlsx_Letter(x+i+1) + str(y-1)] = "log"
        for elem in x_ass:
            workbook[Xlsx_Letter(x+i+1) + str(y+elem-1)] = z[0]*np.log(elem) + z[1]
    trend_data.append(z[0]*np.log(x_ass[-1]) + z[1])
    return workbook, trend_data


def EksponNoglud(workbook,date,data,x,y):
    DataDisplay(workbook,date,data,x-2,y)
    ekspon_data = []
    for i in range(1,len(data)):
        for j in range(1,10):
            workbook[Xlsx_Letter(x+j-1) + str(y-1)] = str(j/10)
            workbook[Xlsx_Letter(x+j-1) + str(y+i+1)] = float(workbook[Xlsx_Letter(x-1) + str(y+i)].value)*(j/10)+(1-j/10)*float(workbook[Xlsx_Letter(x-1) + str(y+i-1)].value)
            if i == len(data)-1:
                ekspon_data.append(float(workbook[Xlsx_Letter(x+j-1) + str(y+i+1)].value))
    return workbook, ekspon_data


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
    sezon_data = float(workbook[Xlsx_Letter(x+4) + str(y+i-1)].value)
    return workbook, sezon_data


def KvadKlud(workbook,date,data,x_sak,y_sak,x_beig,y_beig):
    DataDisplay(workbook,date,data,x_beig+3,y_sak)
    kvad_klud_data = []
    for i in range(1,x_beig-x_sak+1):
        workbook[Xlsx_Letter(x_beig+4+i) + str(y_sak-1)] = "Q" + str(workbook[Xlsx_Letter(x_sak+i-1) + str(y_sak-1)].value)
    j = 0
    for i in range(1,x_beig-x_sak+1):
        for elem in data:
            j = j + 1
            if workbook[Xlsx_Letter(x_sak+i-1) + str(y_sak+j-1)].value != None:
                workbook[Xlsx_Letter(x_beig+4+i) + str(y_sak+j-1)] = pow((float(workbook[Xlsx_Letter(x_sak+i-1) + str(y_sak+j-1)].value)-float(elem)),2)
        j = 0
    workbook[Xlsx_Letter(x_beig+3) + str(y_beig+1)] = "AVERENGE"
    sum = 0
    a = 0
    for i in range(1,x_beig-x_sak+1):
        for l in range(1,y_beig-y_sak+1):
            if workbook[Xlsx_Letter(x_beig+4+i) + str(y_sak+l-1)].value != None:
                sum = sum + float(workbook[Xlsx_Letter(x_beig+4+i) + str(y_sak+l-1)].value)
                a = a+1
        workbook[Xlsx_Letter(x_beig+4+i) + str(y_beig+1)]= sum/a
        kvad_klud_data.append(sum/a)
        sum = 0
        a = 0
    return workbook, kvad_klud_data


def sezon_input(text1,text2,expect1,expect2,expect3):
    print(text1)
    inp = input()
    if inp == expect1 or inp == expect2 or inp == expect3:
        return inp
    else:
        print(text2)
        inp = sezon_input(text1,text2,expect1,expect2,expect3)
        return inp


def first_sezon_input(text1,text2,expect):
    print(text1)
    inp = input()
    if int(inp)<= int(expect) and int(inp)>0:
        return inp
    else:
        print(text2 + str(expect))
        inp = first_sezon_input(text1,text2,expect)
        return inp