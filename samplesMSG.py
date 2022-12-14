from openpyxl import Workbook,load_workbook
from datetime import datetime
from tkinter import *
from tkinter import filedialog
import os
import re
import time
import subprocess
import threading

class Delivery():
    order_number = None
    phone_number = None
    customer_name = None
    day = None
    time_start = None
    time_end = None
    address = None
    message = None

class MainWindow(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        #Window set up
        self.master = master
        master.title("Delivery Message Generator")
        self.master.resizable(0, 0)

        #Class Variables
        self.wb = None
        self.ws = None
        self.path = None
        self.msg = """Artificial Grass Delivery Confirmation- Your {} order has been dispatched and will be delivered {} between {} - {} at {}.To prepare for your delivery please make sure nothing is blocking the delivery location selected. You will receive another text notification 30 minutes prior to arrival. If there is a gate or entry approval, please provide and confirm
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
            command = lambda: threading.Thread(target = self.process).start(),
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
        try:
            self.wb = load_workbook(filename= self.path)
            self.ws = self.wb.active
        except:
            self.edit_entry("Make sure Excel file isn't open")
            return
        self.edit_entry(os.path.split(self.path)[1])

    def process(self):
        if self.path == None:
            self.edit_entry("Please Select an Excel File")
            return

        self.edit_entry("Processing... Please Wait")

        txt_path = os.path.split(self.path)[0] + "/" + datetime.now().strftime("%m-%d-%Y") + ".txt"
        f = open(txt_path, "a")

        i = 1
        while i <= 10000:
            value = str(self.ws[("B" + str(i))].value)
            if value.isnumeric() == True:
                self.current.order_number = value
                self.current.phone_number = re.sub('\D','',str(self.ws[("I" + str(i))].value))
                self.current.customer_name = str(self.ws[("D" + str(i))].value)
                self.get_time(i)
                self.current.address = str(self.ws[("F" + str(i))].value)
                self.current.message = self.msg.format(
                    self.current.order_number,
                    self.current.day, 
                    self.current.time_start,
                    self.current.time_end,
                    self.current.address,
                )
                print(self.current.order_number)
                self.write_to_text(f)
            i = i + 1

        time.sleep(3)
        self.edit_entry("Done!")
        f.close()
        subprocess.Popen(["notepad.exe", txt_path])


    def write_to_text(self, file):
        msg = """Customer Name: {}
Customer Number: {}
{}
================================================================================

"""
        file.write(msg.format(
            self.current.customer_name, 
            self.current.phone_number, 
            self.current.message
            )
        )

    def edit_entry(self,text):
        self.ent_file.config(state = 'normal')
        self.ent_file.delete(0, END)
        self.ent_file.insert(0,text)
        

    def get_time(self, row): #Gets time from spreadsheet and calculates range. Formats time with AM/PM.
        value = str(self.ws[("H" + str(row))].value)
        try:
            start_time = datetime.strptime(value, '%m/%d/%Y %H:%M %p')
        except:
            start_time = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
        start_time = int(start_time.hour) + 1
        self.current.time_start = start_time
        self.current.time_end = start_time + 2

        if self.current.time_start > 12:
            self.current.time_start = str(self.current.time_end - 12)
            self.current.time_start = self.current.time_start + ":00 PM"
        elif self.current.time_start == 12:
            self.current.time_start = str(self.current.time_start) + ":00 PM"
        else:
            self.current.time_start = str(self.current.time_start) + ":00 AM"

        if self.current.time_end > 12:
            self.current.time_end = str(self.current.time_end - 12)
            self.current.time_end = self.current.time_end + ":00 PM"
        elif self.current.time_end == 12:
            self.current.time_end = str(self.current.time_end) + ":00 PM"
        else:
            self.current.time_end = str(self.current.time_end) + ":00 AM"

        if datetime.today().weekday() == 4:
            self.current.day = "Monday"
        else:
            self.current.day = "tomorrow"

if __name__ == "__main__":
    root = Tk()
    main_app = MainWindow(root)
    root.mainloop()