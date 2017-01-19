import openpyxl as Excel

def sumHours(a):
        for k in range(0,len(a)):
            if a[k] == '-':
                first_time = int(a[k-1])
                second_time = int(a[k+1])
            if a[k] == ':':
                semicolon = int(a[k+1])
        if second_time < 7:
            second_time = second_time + 12
        length = len(a)
        sums = first_time - second_time
        if semicolon == 3:
            sums += 0.5
        print('Number of hours worked was: '+str(sums))
        return sums

def Excel_read():
    wb = Excel.load_workbook('JANUARY DRAFT.xlsx')
    ws = wb.active
    hours_worked = 0
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
                                    print('')
                                    print('9:30-6:30 at cell ' + str(a))
                                    hours_worked += 8
                                else:
                                    hours = a.value[k+4:length]
                                    print('')
                                    print(str(hours) + ' at cell ' + str(a))
                                    hours_worked+= sumHours(hours)
    print('')
    print('Monthly Hours ' + str(hours_worked))
    print('Weekly Hours ' + str(hours_worked/4))

def Main():
    print('---------------------------------------------')
    Excel_read()

Main()
