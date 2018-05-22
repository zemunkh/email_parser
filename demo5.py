import csv
import xlsxwriter
import outlook
from datetime import datetime, date, time

row = 0
num = 1

# Acquiring Time Value for the text name 
date = datetime.now().isoformat("-").split(".")[0].replace(":","-")

# Opening a file as excel format
workbook = xlsxwriter.Workbook('XeroxLog_' + date + '.xlsx')
worksheet = workbook.add_worksheet()


mail = outlook.Outlook()
mail.login("mail@mail.com","password")
mail.inbox()

## Creating Buffer file for email
buff = open("buffer.txt", 'w')
buff.write('%s' % mail.read())
buff.close()

while True:
        
    f = open("buffer.txt", 'r')

    while True:
        try:
            text = f.readline()
            if text != '0':
                if 'X-Xerox-Source-Name:' in text:
                    print text
                    field = text.split(":")
                    sourceName = field[1]
                else:
                    sourceName = '\0'
                if 'X-Xerox-DeviceName:' in text:
                    print text
                    field = text.split(":")
                    deviceName = field[1]
                else:
                    deviceName = '\0'
                if 'Delivery-date:' in text:
                    print text
                    field = text.split(",")
                    time = field[1]
                else:
                    time = '\0'
                if 'System Location:' in text:
                    print text
                    field = text.split(":")
                    sysLoc = field[1]
                else:
                    sysLoc = 'NaN'
                if 'IP address:' in text:
                    print text
                    field = text.split(":")
                    ipAddr = field[1]
                else:
                    ipAddr = '\0'
                if 'System Model:' in text:
                    print text
                    field = text.split(":")
                    sysMod = field[1]
                else:
                    sysMod = '\0'
                if 'System Serial Number:' in text:
                    print text
                    field = text.split(":")
                    serNum = field[1]
                    break
                else:
                    serNum = '\0'
            else:
                print("Parameters are not found! \n")   
        except ValueError:
            print("Tanii hussen medeelel baihgui baina. Dahin oroldono uu...")
                
    print('Now creating Excel file...')       
    # Assigning Parsed value into matrix
    dic = [num, sourceName, deviceName,
           time, ipAddr, sysMod,
           serNum, sysLoc ]

    if num == 1:
        # Defining Headers of parameters
        header = ['Number', 'Source-Name', 'DeviceName', 'Delivery-date',
                  'IP address', 'System Model', 'System Serial Number', 'System Location']

        # Printing headers with bold text
        cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})

        col = 0
        j=1
        #worksheet.set_column(0, 0, 3)
        for j, t in enumerate(header):
            worksheet.set_column(row + 1, col + j, 25)
            worksheet.write(row, col + j, t, cell_format)
            j +=1

        col = 0
        k = 0
        for key in range(len(dic)):
            value = dic[key]
            #worksheet.set_column(row + 1, col + k, 25)
            worksheet.write(row + 1, col + k, value)
            k +=1
            
    else:     
        # Assigning Parsed data and declaring loop index

        col = 0
        k = 0
        for key in range(len(dic)):
            value = dic[key]
            #worksheet.set_column(row + 1, col + k, 25)
            worksheet.write(row + 1, col + k, value)
            k +=1
        break
            
    # Adding counter values
    num += 1
    row += 1
    sourceName = '\0'
    deviceName = '\0'
    time  = '\0' 
    ipAddr = '\0'
    sysMod = '\0'
    serNum = '\0'

workbook.close()
print("File saved successfully...") 
