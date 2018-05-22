from Tkinter import *
import xlsxwriter
import time
import os
import outlook
from datetime import datetime, date
import schedule
import threading
# creates the main window object, defines its name, and default size
main = Tk()
main.title('Xerox Log Application')
main.geometry('225x150')
 
def run_schedule():
    schedule_clock = '13:40'
    print ("set clock is " + schedule_clock)
    schedule.every().day.at(schedule_clock).do(email_parser)
    print ("Xerox Email Filter Program will start at " + schedule_clock)
    print ("Now is waiting!")
    schedule.run_pending()
    print ("Program will be executed at " + schedule_clock)
    print (datetime.now())
    

def start_cmd(*event):
    # Able to be called from a key binding or a button click because of the '*event'
    print 'Schedule Time is set ' 
    ## Running Parallel processing unit
    main.destroy()
    while True:
        run_schedule()
        time.sleep(60)

    # If I wanted I could also pass the username and password I got above to another 
    # function from here.
def stop(*event):
    print ("Program is stopped ", datetime.now().isoformat("-").split(".")[0].replace(":","-"))
    second.destroy()

def email_parser():
    mail = outlook.Outlook()
    mail.login("email","pass")
    mail.inbox()
    mail.select("Inbox")

    # Acquiring Time Value for the text name 
    date = datetime.now().isoformat("-").split(".")[0].replace(":","-")

    # Variable declations
    row = 0
    num = 1
    num_list = 0

    # Main parameters declaration
    param_flag = 0
    xerox_flag = 0
    sourceName = '\0'
    deviceName = '\0'
    delivery_date  = '\0' 
    ipAddr = '\0'
    sysMod = '\0'
    serNum = '\0'
    sysLoc = '\0'
    totImp = '\0'
    #print mail.getIdswithWord(5, "Subject")
    print mail.readIdsToday()

    list_id = mail.readIdsToday()
    ## Assigning the number of iteration in excel file
    num_list = len(list_id) 
    if list_id[0] == '':
        print "Emails are already synced..."
    else:
        # Opening a file as excel format
        workbook = xlsxwriter.Workbook('XeroxLog_' + date + '.xlsx')
        worksheet = workbook.add_worksheet()
        for i in range (len(list_id)):
            print list_id[i]
            print ("Checking Email ID with " + list_id[i] + " ..." )
            buff = open("buffer.txt", 'w')
            ## Creating Buffer file for email
            mail_buffer = mail.getEmail(list_id[i])
            buff.write('%s' % mail_buffer)
            buff.close()
            
            ## Now parse and save to excel file 
            f = open("buffer.txt", 'r')
            
            while True:
                try:
                    text = f.readline()
                    if text != '0':
                        if 'X-Xerox-Source-Name:' in text:
                            print text
                            field = text.split(":")
                            sourceName = field[1]
                            param_flag = 1
                            xerox_flag = 1

                        if 'X-Xerox-DeviceName:' in text:
                            print text
                            field = text.split(":")
                            deviceName = field[1]
                            param_flag = 1

                        if 'Delivery-date:' in text:
                            print text
                            field = text.split(",")
                            delivery_date = field[1]
                            if xerox_flag == 1:
                                param_flag = 1
                        if 'System Location:' in text:
                            print text
                            field = text.split(":")
                            sysLoc = field[1]
                            param_flag = 1

                        if 'IP address:' in text:
                            print text
                            field = text.split(":")
                            ipAddr = field[1]
                            param_flag = 1

                        if 'System Model:' in text:
                            print text
                            field = text.split(":")
                            sysMod = field[1]
                            param_flag = 1
                        
                        if 'System Serial Number:' in text:
                            print text
                            field = text.split(":")
                            serNum = field[1]
                            param_flag = 1
                            
                        if 'Total Impressions:' in text:
                            print text
                            field = text.split(":")
                            totImp = field[1]
                            param_flag = 1
                            break
                        else:
                            ## Check it is the last line of text
                            if text == '':
                                print "There are no Xerox Notification E-mails."
                                break
                    else:
                        param_flag = 0
                        print("Parameters are not found! \n")
                        break
                except ValueError:
                    print("Tanii hussen medeelel baihgui baina. Dahin oroldono uu...")
            f.close()
            try:
                os.remove('buffer.txt')
            except OSError:
                pass
            
            print('Now writing on Excel file...')       
            # Assigning Parsed value into matrix
            # Checking parameter is available on the email with current ID
            if param_flag == 1:
                dic = [num, sourceName, deviceName, totImp,
                   delivery_date, ipAddr, sysMod,
                   serNum, sysLoc ]

                if num == 1:
                # Defining Headers of parameters
                    header = ['Number', 'Source-Name', 'DeviceName', 'Total Impressions', 'Delivery-date',
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

                    num = 2
                else:     
                    # Assigning Parsed data and declaring loop index
                    if num_list < num:
                        #break
                        print "All new mails are successfully saved to Excel file."
                    else:
                        col = 0
                        k = 0
                        for key in range(len(dic)):
                            value = dic[key]
                            #worksheet.set_column(row + 1, col + k, 25)
                            worksheet.write(row + 1, col + k, value)
                            k +=1
                        #break
                        num += 1
                # Adding counter values
                
                #Reset variable for the next operation
                row += 1
                sourceName = '\0'
                deviceName = '\0'
                delivery_date  = '\0' 
                ipAddr = '\0'
                sysMod = '\0'
                serNum = '\0'
                sysLoc = '\0'
                totImp = '\0'
                param_flag = 0
                xerox_flag = 0

            #Delay for server side
            time.sleep(.200)
            
        workbook.close()
        print("Excel File is saved successfully...") 
#################################################################################
#################################################################################
    
# defines a grid 50 x 50 cells in the main window
rows = 0
while rows < 10:
    main.rowconfigure(rows, weight=1)
    main.columnconfigure(rows, weight=1)
    rows += 1
 
lbl = Label(main, text="XEROX", fg = 'Navy', font=("Helvetica", 20,"bold italic"), bg='sky blue')
lbl.grid(row=0, column=5) 

# adds start button and defines its properties
start_btn = Button(main, text='Start', command=start_cmd)
start_btn.bind('<Return>', start_cmd)
start_btn.grid(row=5, column=5, sticky='NESW')


main.mainloop()
############# Starting Next Window ###############
schedule_clock = '15:00'


second = Tk()
second.title('Xerox App: Status Box')
second.geometry('225x150')

rows = 0
while rows < 10:
    second.rowconfigure(rows, weight=1)
    second.columnconfigure(rows, weight=1)
    rows += 1

# adds start button and defines its properties
stop_btn = Button(second, text='Stop', command=stop)
stop_btn.bind('<Return>', stop)
stop_btn.grid(row=5, column=5, sticky='NESW')


lbl = Label(second, text="Xerox App", fg = 'red', font=("Helvetica", 18, "italic bold"), bg='sky blue')
lbl.grid(row=2, column=5)
lbl2 = Label(second, text="Program started successfully" , font=("Helvetica", 8, "italic "))
lbl2.grid(row=3, column=5)

time_lbl = Label(second, text="Scheduled Time: " + schedule_clock, fg = 'dodger blue', font=("Helvetica", 10, "bold") )
time_lbl.grid(row=4, column=5) 


second.mainloop()
