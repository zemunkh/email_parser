from Tkinter import *
import xlsxwriter
import time
import os
import outlook
from datetime import datetime, date
# creates the main window object, defines its name, and default size
main = Tk()
main.title('Authentication Box')
main.geometry('225x150')
 
def clear_widget(event):
 
    # will clear out any entry boxes defined below when the user shifts
    # focus to the widgets defined below
    if username_box == main.focus_get() and username_box.get() == 'E-mail':
        username_box.delete(0, END)
    elif password_box == password_box.focus_get() and password_box.get() == '     ':
        password_box.delete(0, END)
 
def repopulate_defaults(event):
 
    # will repopulate the default text previously inside the entry boxes defined below if
    # the user does not put anything in while focused and changes focus to another widget
    if username_box != main.focus_get() and username_box.get() == '':
        username_box.insert(0, 'E-mail')
    elif password_box != main.focus_get() and password_box.get() == '':
        password_box.insert(0, '     ')

def login(*event):
    # Able to be called from a key binding or a button click because of the '*event'
    print 'E-mail: ' + username_box.get()
    print 'Password: ' + password_box.get()
    email_id = username_box.get()
    password = password_box.get()

    buff = email_parser(email_id, password)
    print ("Number of Xerox Email = ", buff[0]) 
    print ("Notification = ", buff[1]) 
    if buff[1] == 1:
        print "Program should be stopped"
        status_lbl.configure(text = "Wrong Email or Password.")
        #main.destroy()
        #second.destroy()
    else:
        print "Successfully Signed In..."
        main.destroy()
    #main.destroy()
    
def stop(*event):
    print ("Program is stopped ")
    second.destroy()

def email_parser(email_id, password):
    # Main parameters declaration
    param_flag = 0
    xerox_flag = 0
    xerox_counter = 0
    sourceName = '\0'
    deviceName = '\0'
    delivery_date  = '\0' 
    ipAddr = '\0'
    sysMod = '\0'
    serNum = '\0'
    sysLoc = '\0'
    totImp = '\0'
    
    mail = outlook.Outlook()
    noft = mail.login(email_id, password)
    if noft == 0:

        mail.inbox()
        mail.select("Inbox")

        # Acquiring Time Value for the text name 
        date = datetime.now().isoformat("-").split(".")[0].replace(":","-")

        # Variable declations
        row = 0
        num = 1
        num_list = 0

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
                                xerox_counter += 1

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
                            print("Email is empty \n")
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
    return [xerox_counter, noft]
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

# adds username entry widget and defines its properties
username_box = Entry(main)
username_box.insert(0, 'E-mail')
username_box.bind("<FocusIn>", clear_widget)
username_box.bind('<FocusOut>', repopulate_defaults)
username_box.grid(row=2, column=5, sticky='NS')
 
 
# adds password entry widget and defines its properties
password_box = Entry(main, show='*')
password_box.insert(0, '     ')
password_box.bind("<FocusIn>", clear_widget)
password_box.bind('<FocusOut>', repopulate_defaults)
password_box.bind('<Return>', login)
password_box.grid(row=3, column=5, sticky='NS')

status_lbl = Label(main, text="Hello")
status_lbl.grid(row=4, column=5) 
 
# adds login button and defines its properties
start_btn = Button(main, text='Login', command=login)
start_btn.bind('<Return>', login)
start_btn.grid(row=6, column=5, sticky='NESW')
 
main.mainloop()

########################|^^^^|##################################

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

time_lbl = Label(second, text="Program is done", fg = 'dodger blue', font=("Helvetica", 10, "bold") )
time_lbl.grid(row=4, column=5) 

second.mainloop()

