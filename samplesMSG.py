from openpyxl import Workbook,load_workbook
from datetime import datetime
from tkinter import *
from tkinter import filedialog
from parsedatetime import Calendar
import os
import re
import subprocess


class WorkOrder():
    order_number = None
    phone_number = None
    customer_name = None
    delivery_day = None
    start_time = None
    end_time = None
    address = None
    message = None

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
        self.order = WorkOrder()

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
            command = self.process,
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

    #Used to write the order and its message to the text document
    def write_to_text(self, file):
        msg = """Customer Name: {}
Customer Number: {}
{}
================================================================================"""
        file.write(msg.format(
            self.order.customer_name, 
            self.order.phone_number, 
            self.order.message
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
                blank_rows = 0      #Reset blank row count as order was found
                self.order.order_number = value
                self.order.phone_number = re.sub('\D','',str(self.ws[("I" + str(row))].value))
                self.order.customer_name = str(self.ws[("D" + str(row))].value)
                self.get_time(row)
                self.order.address = str(self.ws[("F" + str(row))].value)
                self.order.message = message.format(
                    self.order.order_number,
                    self.order.delivery_day, 
                    self.order.start_time,
                    self.order.end_time,
                    self.order.address,
                )
                self.write_to_text(f)
            else:
                if value == "None":
                    blank_rows = blank_rows + 1
                else:
                    blank_rows = 0      #Reset blank row count as value of some kind was found
            
            print(blank_rows)
            if blank_rows > 4:          #if we have 4 or more blank rows in a row
                break                   #End the loop
            else:
                row = row + 1           #otherwise move to the next row

        self.edit_entry("Done!")
        f.close()
        subprocess.Popen(["notepad.exe", txt_path])     #Open the written text file so the user can send messages

    def get_time(self, row): #Gets time from spreadsheet and calculates range. Formats time with AM/PM.
        value = str(self.ws[("H" + str(row))].value)
        cal = Calendar()      #Create calendar object so we can parse time
        time = cal.parse(value)             #Convert human readable time to timedate object
        start_time = time[0].tm_hour        #Set start_time to hour found before
        if start_time <= 3:     #Early Morning deliveries are all set to 4am
            start_time = 4
        self.order.start_time = start_time
        self.order.end_time = start_time + 2

        if self.order.start_time > 12:
            self.order.start_time = str(self.order.start_time - 12)
            self.order.start_time = self.order.start_time + ":00 PM"
        elif self.order.start_time == 12:
            self.order.start_time = str(self.order.start_time) + ":00 PM"
        else:
            self.order.start_time = str(self.order.start_time) + ":00 AM"

        if self.order.end_time > 12:
            self.order.end_time = str(self.order.end_time - 12)
            self.order.end_time = self.order.end_time + ":00 PM"
        elif self.order.end_time == 12:
            self.order.end_time = str(self.order.end_time) + ":00 PM"
        else:
            self.order.end_time = str(self.order.end_time) + ":00 AM"

        if datetime.today().weekday() == 4:     #Check if sending day is Friday. If it is change the message to say Monday delivery day
            self.order.delivery_day = "Monday"
        else:
            self.order.delivery_day = "tomorrow"

if __name__ == "__main__":
    root = Tk()
    main_app = MainWindow(root)
    root.mainloop()
