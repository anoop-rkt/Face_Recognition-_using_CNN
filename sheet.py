import os
import datetime
from openpyxl import load_workbook
from openpyxl import Workbook

st_name = 'Anoop'


# Get current date
today_date = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
def mark_present(st_name):
    names = os.listdir('output/')
    # print(names)

    
    sub = 'Attendance_' + today_date

    # Check if the attendance file for the current date exists
    attendance_file = 'attendance/' + sub + '.xlsx'
    if not os.path.exists(attendance_file):
        workbook = Workbook()
        print("Creating Spreadsheet with Title: " + sub)
        sheet = workbook.active
        sheet.title = sub  # Set sheet title to the current date
        count = 2
        for i in names:
            sheet.cell(row=count, column=1).value = i
            count += 1
        workbook.save(attendance_file)

    # Load workbook using openpyxl
    wb = load_workbook(attendance_file)

    # Get reference to active sheet
    sheet = wb.active

    # Write current timestamp to cell B2
    sheet['B2'] = str(datetime.datetime.now())

    count = 2
    for i in names:
        if i in st_name:
            sheet.cell(row=count, column=2).value = 'P'
        else:
            sheet.cell(row=count, column=2).value = 'A'
        sheet.cell(row=count, column=1).value = i
        count += 1

    wb.save(attendance_file)


mark_present(st_name)
