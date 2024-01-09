import math
import openpyxl

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
