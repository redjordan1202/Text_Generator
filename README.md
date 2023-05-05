# ETA Text Generator
A lightweight script for automatically generating ETA text messages

## About this project
This project was created to automate the creation of ETA text messages. It reads information directly from the route, 
pulls the information needed to generate the message, and then outputs the messages for all orders to a text file.

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
- openpyxl (For reading xlsx files)
- parsedatetime (For parsing date/time from Excel sheet)

All required packages and versions used are listed in requirements.txt

### Spreadsheet requirements
The spreadsheet used must be in xlsx format. Column A must-have Service Resource: Name somewhere in the first 10 rows.
This is needed to tell the script where the header row is so the script can see where the other column headers are.

Please let me know if any major changes are made to the spreadsheet so that testing can be done and any needed changes 
made to the script.