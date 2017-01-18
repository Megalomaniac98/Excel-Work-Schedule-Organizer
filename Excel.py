import openpyxl as Excel

def sumHours(b):
    if b[0] == '':
        return
    else:

        length = len(b)
        sum = int(b[0])+int(b[2])
        if int(b[4]) == 3:
            sum += 0.5
        print(sum)
        return sum

def Excel_read():
    wb = Excel.load_workbook('JANUARY DRAFT.xlsx')
    ws = wb.active
    for i in range(1,101):
        for j in range(1,101):
            a = ws.cell(row = i, column = j)
            if a.value != None:
                if isinstance(a.value,str) :
                    for k in range(0,len(a.value)):
                        if a.value[k] == 'A':
                            if a.value[k+1] == 'r' and a.value[k+2] == 'i':
                                length = len(a.value)
                                if a.value[k+4:length] == '':
                                    print('All Day')
                                else:
                                    list = a.value[k+4:length]
                                    print(list)
                                    print(sumHours(list))
                                    print(a)

def Main():
    Excel_read()

Main()
