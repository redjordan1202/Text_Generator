# ETA Text Generator
A lightweight script for automatically generating ETA text messages

## About this project
This project was created to automate the creation of ETA text messages. It reads information directly from the route, 
pulls the information needed to generate the message, and then outputs the messages for all orders to a text file.

## Current Version
### Version 3 
### The Pandas Update
New version uses pandas to convert excel data directly to a python dict. This allows for simpler parsing of the data and
faster processing due to only needing one call to read the Excel sheet. 

## How to use

- Download the .exe file from the [releases page](https://github.com/redjordan1202/Text_Generator/releases "Text Generator Releases").
Make sure to get the newest version listed.
- Run the .exe
- You will be prompted to select the route sheet you want to use
- Follow the prompts to start processing
- Once processing is complete the text file will open automatically

## Requirements
### Script requirements
There are no requirements to run the exe version of this script. 
However, if you want to run the source directly you will need the following reqirements
- Python 3.10 or newer (Needed to use Switch/Case)
- pandas (For reading excel sheets and parsing to python dict)
- parsedatetime (For parsing date/time from Excel sheet)

All required packages and versions used are listed in requirements.txt

### Spreadsheet requirements
Spreadsheet must have columns in the following order.
1. Service Resource: Name
2. Work Order Number
3. Appointment Number
4. Account: Account Name
5. Description
6. Address
7. \* Can be any data in this column \*
8. Scheduled Start
9. Contact: Phone Number

Column 7 can have any data, but can not be missing. 