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
    return(a)

#print(Xlsx_Letter(1352))
wb = openpyxl.Workbook()
ws=wb.active

for i in range (1,10000):
    index = Xlsx_Letter(i)
    ws[index + '1'] = index

#wb.save("Test.xlsx")
wb.close()