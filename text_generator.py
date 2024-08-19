import os
import re
import subprocess
import sys
import warnings
import math
from time import sleep
from tkinter import filedialog as fd
from datetime import datetime

import pandas
from parsedatetime import Calendar

# Text color escape codes for Windows terminal
text_colors = {
    "green": '\033[92m',
    "red": '\033[91m',
    "endc": '\033[0m'
}

# Message Template to include in the final output
MESSAGE = """Artificial Grass Delivery Confirmation- Your order, {}, has been dispatched and will be delivered {} between {} - {} at {}.
To prepare for your delivery please make sure nothing is blocking the delivery location selected. 
You will receive another text notification 30 minutes prior to arrival. If there is a gate or entry approval, please provide and confirm."""


def main():
    has_error = False
    header_row = None
    columns = dict(
        wo_number=0,
        address=0,
        phone_number=0,
        start_time=0,
        customer_name=0,
        so_number=0
    )

    platform = None
    term_clear = None
    txt_editor = None

    print(
        f"{text_colors['green']}Turf Distributors ETA Text Message Generator")
    print("Written by Jordan Del Pilar")
    print("jordan.delpilar@turfdistributors.com")
    print("Version 3.0 The Pandas Update")
    print(f"{text_colors['endc']}")
    print("=" * 80)
    print("\n")

    match sys.platform:

        case "win32":
            platform = "windows"
            term_clear = "cls"
            txt_editor = "notepad.exe"

        case "darwin":
            platform = "macos"
            term_clear = "clear"
            txt_editor = "open -a TextEdit"

        case _:
            print(f"{text_colors['red']}\n!!! Warning !!!{text_colors['endc']}")
            print("This script must be run on either MacOS or Windows")
            input("Press Enter to exit")

    # Loop through file selection until user selects a file
    while True:
        print("Please select route sheet... ")
        sleep(2)
        sheet = fd.askopenfilename(
            title="Select File...",
            initialdir=os.path.expanduser("~"),
            filetypes=(("Excel Spreadsheet", "*.xlsx"),)
        )

        if sheet:
            break

    os.system(term_clear)

    try:
        with warnings.catch_warnings(record=True):
            excel_data = pandas.read_excel(
                sheet, "Route Sheet Report Inc Weight")
            excel_dict = excel_data.to_dict(orient='records')
    except Exception:
        print(
            f"{text_colors['red']}ERROR{text_colors['endc']} -  Could not open Excel file.")
        print("Please make sure the Excel file is not open")
        input("Press Enter to Close")
        sys.exit()

    os.system(term_clear)

    print("Looking for header row")
    i = 0
    header_row = 0
    for row in excel_dict:
        for cell in row.items():
            if cell[1] == "Work Order Number":
                header_row = i
            else:
                continue
        i += 1

    if header_row is not None:
        print("Header Row Found\nParsing Columns")
    else:
        print(f"{text_colors['red']}ERROR{text_colors['endc']} - Unable to find header row.")
        input("Press Enter to Close")

    for cell in excel_dict[header_row].items():
        match cell[1]:

            case "Work Order Number":
                columns["wo_number"] = cell[0]
                continue

            case "Address":
                columns["address"] = cell[0]
                continue

            case "Contact Phone Number":
                columns["phone_number"] = cell[0]
                continue

            case "Scheduled Start":
                columns["start_time"] = cell[0]
                continue

            case "Account: Account Name":
                columns["customer_name"] = cell[0]
                continue

            case "Appointment Number":
                columns["so_number"] = cell[0]
                continue

            case _:
                continue

    for col in columns.values():
        if col == 0:
            print(f"{text_colors['red']}ERROR{text_colors['endc']} - Unable to find all needed columns.")
            input("Press Enter to Close")
        else:
            continue
    print("Header Parsing complete")

    txt_path = os.path.split(sheet)[0] + '/' + \
               datetime.now().strftime("%m-%d-%Y") + ".txt"
    txt = open(txt_path, "a", encoding="utf-8")

    os.system(term_clear)
    print(f"{text_colors['green']}Starting Processing{text_colors['endc']}")

    records_processed = 0
    for record in excel_dict:
        is_work_order = False
        work_order = record[columns["wo_number"]]

        if isinstance(work_order, float):
            if math.isnan(work_order):
                continue
            else:
                print(f"{work_order}\t{math.isnan(work_order)}")
                is_work_order = True
        elif isinstance(work_order, str) and work_order != "nan":
            if work_order.isnumeric():
                is_work_order = True

        if is_work_order:
            print(f"*** Processing Work Order {work_order} ***")

            customer_state = None
            if record[columns["address"]].find("FL") > 0:
                customer_state = "FL"
            elif record[columns["address"]].find("NV") > 0:
                customer_state = "NV"
            else:
                customer_state = "CA"

            phone_number = re.sub(r'\D', '', record[columns["phone_number"]])
            if len(phone_number) > 0:
                if phone_number[0] == "1":
                    phone_number = "+" + phone_number
                else:
                    phone_number = "+1" + phone_number
                if len(phone_number) > 12:
                    phone_number = phone_number[:12]
            else:
                phone_number = "None Provided"

            raw_time = record[columns["start_time"]]
            if isinstance(raw_time, datetime):
                try:
                    raw_time = raw_time.strftime("%m/%d/%Y %H:%M:%S %p")
                except Exception:
                    try:
                        raw_time = raw_time.strftime("%m/%d/%Y %H:%M %p")
                    except Exception:
                        print(
                            f"{text_colors['red']}ERROR{text_colors['endc']} - Failed to convert date to string. Skipping Row")
                        has_error = True
                        continue
            try:
                cal = Calendar()
                time = cal.parse(raw_time)
            except TypeError:
                print(
                    f"{text_colors['red']}ERROR{text_colors['endc']} - Failed to parse date. Skipping Row")
                has_error = True
                continue

            start_time = time[0].tm_hour
            if customer_state == "FL":
                start_time += 3
            if customer_state == "NV":
                start_time += 1

            if start_time <= 3:
                start_time = 4
            end_time = start_time + 2

            if start_time > 12:
                start_time = str(start_time - 12) + ":00 PM"
            elif start_time == 12:
                start_time = str(start_time) + ":00 PM"
            else:
                start_time = str(start_time) + ":00 AM"

            if end_time > 12:
                end_time = str(end_time - 12) + ":00 PM"
            elif end_time == 12:
                end_time = str(end_time) + ":00 PM"
            else:
                end_time = str(end_time) + ":00 AM"

            if datetime.today().weekday() == 4:
                delivery_day = "Monday"
            else:
                delivery_day = "tomorrow"

            customer_msg = MESSAGE.format(
                work_order,
                delivery_day,
                start_time,
                end_time,
                record[columns["address"]]
            )

            print("Writting to File...")
            txt.write(
                f"""Work Order Number: {work_order}
Service Appointment Number: {record[columns["so_number"]]}
Customer Name: {record[columns["customer_name"]]}
Phone Number: {phone_number}
Customer State: {customer_state}


{customer_msg}


================================================================================\n
""")

            records_processed += 1

        sleep(1)

    txt.close()

    os.system(term_clear)

    try:
        subprocess.Popen(f"{txt_editor} {txt_path}", shell=True)
    except FileNotFoundError:
        print(
            f"{text_colors['red']}ERROR{text_colors['endc']} - Failed to open notepad. Please open the file directly")

    print(f"{text_colors['green']}")
    print("=" * 80)
    print("Processing Finished!")
    if has_error:
        print(f"{text_colors['red']}\n!!! Warning !!!{text_colors['endc']}")
        print("There were some errors encounted during processing. Some Orders may not have been processed.")
        print("Please verify which orders were processed and report any errors displayed to Jordan\n")
        print(f"{text_colors['green']}")
        print("=" * 80)
    print(f"{records_processed} Work Orders Processed")
    input(f"{text_colors['endc']}Press Enter to exit")


if __name__ == '__main__':
    main()
