from openpyxl import Workbook,load_workbook
from datetime import datetime
from tkinter import *
from tkinter import filedialog
from tkinter import scrolledtext
from parsedatetime import Calendar
from time import sleep
import threading
import os
import re
import subprocess
import requests


class WorkOrder():
    order_number = None
    phone_number = None
    customer_name = None
    delivery_day = None
    start_time = None
    end_time = None
    address = None
    message = None

    def __str__(self):
        return self.order_number

message =  """Artificial Grass Delivery Confirmation- Your order, {}, has been dispatched and will be delivered {} between {} - {} at {}.To prepare for your delivery please make sure nothing is blocking the delivery location selected. You will receive another text notification 30 minutes prior to arrival. If there is a gate or entry approval, please provide and confirm."""

class MainWindow(Frame):
    def __init__(self, master = None):
        Frame.__init__(self, master)

        #Initial Window Set Up
        self.master = master
        master.title("Delivery Message Generator")
        self.master.resizable(0,0)      #Setting window to not be resizable

        #Class Variables
        self.wb = None      #Defining Workbook var. Keeping as None for now
        self.ws = None      #Defining Worksheet var. Keeping as None for now
        self.path = None    #Var to hold path to text file. Set to be same directory as the spreadsheet

        self.order_list = []

        #Widget Definitions
        self.lbl_file = Label(master = self.master, text = "Selected File:")

        self.ent_file = Entry(
            master = self.master,
            state = "disabled",
            width = 75,
            justify = LEFT,
            disabledbackground = 'white',       #Overriding the default disabled background
            disabledforeground = 'black'        #Overriding the default disabled text color
        )

        self.btn_file = Button(
            master = self.master,
            text = "Choose File...",
            command = self.choose_file,
            width = 20
        )
        self.btn_run = Button(
            master = self.master,
            text = "Run",
            command = lambda: threading.Thread(target = self.process).start(),
            width = 20
        )

        #Widget Deployment
        self.lbl_file.grid(column = 1, row = 0)
        self.ent_file.grid(column = 0, row = 1, columnspan = 3, pady = 10)
        self.btn_file.grid(column = 0, row = 2, padx = 50, pady = 10)
        self.btn_run.grid(column = 2, row = 2, padx = 50, pady = 10)

    #Class Functions

    #Used to Edit ent_file
    def edit_entry(self,text):  
        self.ent_file.config(state = 'normal')      #Enabling the entry so I can be written
        self.ent_file.delete(0, END)                #Deleting all text currently in the entry
        self.ent_file.insert(0,text)                #Writing the new text to the entry

    #Create Check Messages Window
    def create_check_window(self):
        self.check_window = CheckWindow(self)

    #Used to write the order and its message to the text document
    def write_to_text(self, file, order):
        msg = """Customer Name: {}
Customer Number: {}
{}
================================================================================
"""
        file.write(msg.format(
            order.customer_name, 
            order.phone_number, 
            order.message
            )
        )

    #Function to handle spreadsheet opening
    def choose_file(self):  
        filetypes = (("Excel Spreadsheet", "*.xlsx"),)      #Setting the file types that are displayed in the file dialogue
        self.path = filedialog.askopenfilename(
            title = "Open File...",
            initialdir = os.path.expanduser('~'),           #Default directory is the Users "Home" Directory
            filetypes = filetypes
        )
        try:
            self.wb = load_workbook(filename= self.path)
            self.ws = self.wb.active
        except:
            self.edit_entry("Make sure Excel file isn't open")
            return
        self.edit_entry(os.path.split(self.path)[1])        #Writing just the file name to the entry

    #Main Processing function
    def process(self):
        if self.path == None:                               #Check to ensure a file has been selected
            self.edit_entry("Please Select an Excel File")
            return

        self.edit_entry("Processing... Please Wait")
    
        txt_path = os.path.split(self.path)[0] + "/" + datetime.now().strftime("%m-%d-%Y") + ".txt" #Creating path to text in same directory as spreadsheet
        f = open(txt_path, "a")

        
        row = 1             #Set Initial Row number
        blank_rows = 0      #Var to count number of blank rows
        while row <= 10000:
            value = str(self.ws[("B" + str(row))].value)
            if value.isnumeric() == True:
                order = WorkOrder()
                blank_rows = 0      #Reset blank row count as order was found
                order.order_number = value
                order.phone_number = re.sub('\D','',str(self.ws[("I" + str(row))].value))
                order.customer_name = str(self.ws[("D" + str(row))].value)
                self.get_time(row, order)
                order.address = str(self.ws[("F" + str(row))].value)
                order.message = message.format(
                    order.order_number,
                    order.delivery_day, 
                    order.start_time,
                    order.end_time,
                    order.address,
                )
                self.order_list.append(order)
                self.write_to_text(f, order)
                
            else:
                if value == "None":
                    blank_rows = blank_rows + 1
                else:
                    blank_rows = 0      #Reset blank row count as value of some kind was found
            
            if blank_rows > 4:          #if we have 4 or more blank rows in a row
                break                   #End the loop
            else:
                row = row + 1           #otherwise move to the next row

        self.edit_entry("Done!")
        f.close()
        subprocess.Popen(["notepad.exe", txt_path])     #Open the written text file so the user can send messages
        self.create_check_window()

    def get_time(self, row, order): #Gets time from spreadsheet and calculates range. Formats time with AM/PM.
        value = str(self.ws[("H" + str(row))].value)
        cal = Calendar()      #Create calendar object so we can parse time
        time = cal.parse(value)             #Convert human readable time to timedate object
        start_time = time[0].tm_hour        #Set start_time to hour found before
        if start_time <= 3:     #Early Morning deliveries are all set to 4am
            start_time = 4
        order.start_time = start_time
        order.end_time = start_time + 2

        if order.start_time > 12:
            order.start_time = str(order.start_time - 12)
            order.start_time = order.start_time + ":00 PM"
        elif order.start_time == 12:
            order.start_time = str(order.start_time) + ":00 PM"
        else:
            order.start_time = str(order.start_time) + ":00 AM"

        if order.end_time > 12:
            order.end_time = str(order.end_time - 12)
            order.end_time = order.end_time + ":00 PM"
        elif order.end_time == 12:
            order.end_time = str(order.end_time) + ":00 PM"
        else:
            order.end_time = str(order.end_time) + ":00 AM"

        if datetime.today().weekday() == 4:     #Check if sending day is Friday. If it is change the message to say Monday delivery day
            order.delivery_day = "Monday"
        else:
            order.delivery_day = "tomorrow"

class CheckWindow(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.check_window = Toplevel()  #Set the window as a toplevel window
        self.check_window.title("Check Messages")
        self.check_window.resizable(0,0)
        self.check_window.grab_set()    #grab focus

        self.order = None
        self.order_list = StringVar(value = parent.order_list)

        #Widget Definitions
        self.frm_order_list = Frame(
            self.check_window,
            relief = "ridge",
            borderwidth = 4
        )
        self.lbl_order_list = Label(
            self.frm_order_list,
            text = "Work Orders",
            font = ("Ariel", 18)
        )
        self.lbl_order_count = Label(
            self.frm_order_list,
            text = "{} Orders Found".format(len(parent.order_list))
        )
        self.scrollbar = Scrollbar(self.frm_order_list)
        self.lbox_orders = Listbox(
            self.frm_order_list,
            selectmode = SINGLE,
            height = 22,
            width = 22,
            listvariable = self.order_list,
            yscrollcommand = self.scrollbar.set
        )
        self.lbox_orders.config(exportselection = False)         #Prevent listbox selection from un-highlighting on focus change
        self.scrollbar.config(command = self.lbox_orders.yview)  #bind scrollbar to listbox
        
        self.frm_order_info = Frame(
            self.check_window,
            relief = "ridge",
            borderwidth = 4
        )
        self.lbl_order_info = Label(
            self.frm_order_info,
            text = "Order Information",
            font = ("Ariel", 18)
        )
        self.frm_info = Frame(self.frm_order_info)

        self.lbl_order_number = Label(
            self.frm_info,
            text = "Order Number"
        )
        self.lbl_order_phone = Label(
            self.frm_info,
            text = "Phone Number"
        )
        self.ent_order_number = Entry(
            self.frm_info,
            state = "disabled",
            width = 25,
            justify = LEFT,
            disabledbackground = 'white',       #Overriding the default disabled background
            disabledforeground = 'black'        #Overriding the default disabled text color
        )
        self.ent_order_phone = Entry(
            self.frm_info,
            state = "disabled",
            width = 25,
            justify = LEFT,
            disabledbackground = 'white',       #Overriding the default disabled background
            disabledforeground = 'black'        #Overriding the default disabled text color
        )
        self.lbl_message = Label(
            self.frm_info,
            text = "Message"
        )
        self.txt_message = Text(
            self.frm_info,
            width = 70,
            height = 20,
            wrap = WORD
        )
        self.btn_update = Button(
            self.frm_info,
            text = "Update Message",
            command = self.update_message
        )
        self.btn_confirm = Button(
            self.frm_info,
            text = "Send Messages",
            command = lambda: threading.Thread(target = self.send_messages, args = (parent,)).start()
        )

        self.frm_log = Frame(
            self.check_window,
            relief = "ridge",
            borderwidth = 4
        )
        self.log = scrolledtext.ScrolledText(
            self.frm_log,
            width = 70,
            height = 20,
            wrap = WORD,
            state = DISABLED,
        )
        self.btn_finish = Button(
            self.frm_log,
            text = "Finish",
            state = DISABLED,
            command = parent.quit
        )
        
        """
        TODO
        - Change toplevel frames geometry to grid rather then pack
        - Move Send button to the bottom right corner of the window outside the rest of the frames
            - Maybe add an image to the send button. Something like an envelope
        """
        #Place Widgets
        self.frm_order_list.pack(pady = 5, padx = 5, side = LEFT)
        self.lbl_order_list.pack(pady = 5, padx = 5)
        self.lbl_order_count.pack(padx = 5, fill = X)
        self.lbox_orders.pack(side = LEFT, pady = 5, padx = 5, fill = X)
        self.scrollbar.pack(side = RIGHT, fill = Y)
        self.frm_order_info.pack(pady = 5, padx = 5, side = LEFT)
        self.lbl_order_info.pack(pady = 5, padx = 5)
        self.frm_info.pack(pady = 5, padx = 5)
        self.frm_info.pack_propagate(0)
        self.lbl_order_number.grid(column = 0, row = 0)
        self.lbl_order_phone.grid(column = 1, row = 0)
        self.ent_order_number.grid(column = 0, row = 1)
        self.ent_order_phone.grid(column = 1, row = 1)
        self.lbl_message.grid(column = 0, row = 2, columnspan = 2)
        self.txt_message.grid(column = 0, row = 3, columnspan = 2)
        self.btn_update.grid(column = 1, row = 4, padx = 5, pady = 5, sticky = E)
        self.btn_confirm.grid(column = 1, row = 5, padx = 5, pady = 5, sticky = E)
        
        self.lbox_orders.selection_anchor(0)
        self.lbox_orders.bind("<Double-1>", lambda event: self.get_order_info(parent))
        self.lbox_orders.bind("<Map>", lambda event: self.get_order_info(parent))

    #Class Methods
    
    #Function to edit entry widgets
    def edit_entry(self,entry, text):  
        entry.config(state = 'normal')      #Enabling the entry so I can be written
        entry.delete(0, END)                #Deleting all text currently in the entry
        entry.insert(0,text)                #Writing the new text to the entry
        entry.config(state = 'disabled')    #Disabling the entry again
    
    #Function to edit the message text box
    def edit_text(self, content, text_box):
        text_box.delete(1.0, "end")
        text_box.insert(1.0, content)

    def write_log(self,text):
        self.log.config(state = NORMAL)
        self.log.insert("end", text)
        self.log.see(END)
        self.log.config(state = DISABLED)
    
    def get_order_info(self, parent):
        selection = self.lbox_orders.get(ANCHOR)    #Get current selection from listbox
        
        for i in parent.order_list:
            if i.order_number == selection:     #Search order list for order number
                self.order = i

        self.edit_entry(self.ent_order_number, self.order.order_number)
        self.edit_entry(self.ent_order_phone, self.order.phone_number)
        self.edit_text(self.order.message, self.txt_message)

    def update_message(self):
        if self.order == None:
            return
        self.order.message = self.txt_message.get(1.0, 'end-1c')

    def send_messages(self,parent):
        #Create the log box
        self.frm_log.place(relx = 0.5, rely = 0.5, anchor = CENTER)
        self.log.pack(pady = 10, padx = 10)
        self.btn_finish.pack()

        #Disable other buttons and message text box
        self.btn_confirm.config(state = DISABLED)
        self.btn_update.config(state = DISABLED)
        self.txt_message.config(state = DISABLED)

        self.check_window.protocol("WM_DELETE_WINDOW", disable_event)       #Disable X button on window

        hasError = False

        self.write_log("Getting Ready to Send Messages\n")
        self.write_log("***PLEASE DO NOT CLOSE THE PROGRAM***\n")
        self.write_log("Processing {} Orders\n\n".format(len(parent.order_list)))
        sleep(5)

        for order in parent.order_list:
            self.write_log("Sending Work Order {}\n".format(order.order_number))
            payload = {
                "destination" : order.phone_number,
                "source" : "1234567890",
                "clientMessageId" : order.order_number,
                "text" : order.message,
            }
            response = requests.post("https://eoinm2i8j9t7pkn.m.pipedream.net", payload)
            if response.status_code == 200:
                self.write_log("Message Sent to 8x8 Successfully\n\n")
            else:
                self.write_log("***ERROR Code {}***\n".format(response.status_code))
                self.write_log("Message Failed to Send\n")
                hasError = True
            self.write_log("====================\n\n")
        self.write_log("Sending Complete\n")
        if hasError:
            self.write_log("*Warning* At Least Message had an Error During Sending")
        
        self.btn_finish.config(state = NORMAL)

#Function to disable on-click events
def disable_event():
    pass

if __name__ == "__main__":
    root = Tk()
    main_app = MainWindow(root)
    root.mainloop()
