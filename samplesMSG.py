from openpyxl import Workbook,load_workbook
from datetime import datetime
from tkinter import *
from tkinter import filedialog
import os
import re
import time

class Delivery():
    order_number = None
    phone_number = None
    customer_name = None
    time_start = None
    time_end = None
    address = None
    message = None

class MainWindow(Tk):
    def __init__(self, master=None):
        #Window set up
        self.master = master
        master.title("Delivery Message Generator")
        self.master.resizable(0, 0)

        #Class Variables
        self.wb = None
        self.ws = None
        self.path = None
        self.msg = """Artificial Grass Delivery Confirmation- Your {} order has been dispatched and will be delivered tomorrow between {} - {} at {}.To prepare for your delivery please make sure nothing is blocking the delivery location selected. You will receive another text notification 30 minutes prior to arrival. If there is a gate or entry approval, please provide and confirm
        """
        self.current = Delivery()

        #Widget Definitions 
        self.lbl_file = Label(
            master = self.master,
            text = "Selected File:"
        )

        self.ent_file = Entry(
            master = self.master,
            state = "disabled",
            width = 75,
            justify = LEFT,
        )
        self.ent_file.config(
            disabledbackground = 'white',
            disabledforeground = 'black'
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
    def choose_file(self):
        filetypes = (("Excel Spreadsheet", "*.xlsx"),)
        self.path = filedialog.askopenfilename(
            title = "Open File...",
            initialdir = os.path.expanduser('~'),
            filetypes = filetypes
        )
        self.wb = load_workbook(filename= self.path)
        self.ws = self.wb.active
        self.edit_entry(os.path.split(self.path)[1])

    def process(self):
        if self.path == None:
            self.edit_entry("Please Select an Excel File")
            return

        txt_path = os.path.split(self.path)[0] + "/" + datetime.now().strftime("%m-%d-%Y") + ".txt"
        f = open(txt_path, "a")

        i = 1
        while i <= 1000:
            value = str(self.ws[("B" + str(i))].value)
            if value.isnumeric() == True:
                current.order_number = value
                current.phone_number = re.sub('\D','',str(self.ws[("I" + str(i))].value))
                current.customer_name = str(self.ws[("D" + str(i))].value)
                self.get_time(i)
                current.address = str(self.ws[("F" + str(i))].value)
                current.message = msg.format(
                    current.order_number, 
                    current.time_start,
                    current.time_end,
                    current.address,
                )
                self.write_to_text(f)
            i = i + 1

        time.sleep(3)
        self.edit_entry("Done")
        f.close()

    def write_to_text(self, file):
        msg = """Customer Name: {}
Customer Number: {}
{}
================================================================================

"""
        file.write(msg.format(
            current.customer_name, 
            current.phone_number, 
            current.message
            )
        )

    def edit_entry(self,text):
        self.ent_file.config(state = 'normal')
        self.ent_file.delete(0, END)
        self.ent_file.insert(0,text)
        

    def get_time(self, row): #Gets time from spreadsheet and calculates range. Formats time with AM/PM.
        start_time = datetime.strptime(str(self.ws[("H" + str(row))].value), '%m/%d/%Y %H:%M %p')
        start_time = int(start_time.hour) + 1
        current.time_start = start_time
        current.time_end = start_time + 2

        if current.time_start > 12:
            current.time_start = str(current.time_end - 12)
            current.time_start = current.time_start + ":00 PM"
        elif current.time_start == 12:
            current.time_start = str(current.time_start) + ":00 PM"
        else:
            current.time_start = str(current.time_start) + ":00 AM"

        if current.time_end > 12:
            current.time_end = str(current.time_end - 12)
            current.time_end = current.time_end + ":00 PM"
        elif current.time_end == 12:
            current.time_end = str(current.time_end) + ":00 PM"
        else:
            current.time_end = str(current.time_end) + ":00 AM"

if __name__ == "__main__":
    root = Tk()
    main_app = MainWindow(root)
    root.mainloop()