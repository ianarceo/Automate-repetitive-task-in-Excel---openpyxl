# Compute if the employee rendered overtime.
# Program incomplete.

import openpyxl
from datetime import datetime
from datetime import timedelta

def if_late(time_in):
    pass

wb = openpyxl.load_workbook(r"C:\Users\BG-PC164a\PycharmProjects\pythonProject2\dailyattendance.xlsx")
ws = wb.active

comp_hrs = timedelta(hours=9)
reg_hrs = timedelta(hours=8)
lunch_break = timedelta(hours=1)

for row in ws.iter_rows(min_row=3, max_col=34, max_row=10000, values_only=True):
    if not row[0]:
        break
    if row[9] == "CWW":  # Compressed Work Week 7 - 5
        timestamp = (datetime.strptime(str(row[5]), "%H:%M:%S") - datetime.strptime(str(row[4]), "%H:%M:%S")) \
                    - comp_hrs - lunch_break

    elif row[9] == "REG":  # Regular Working Hours 7 - 4
        timestamp = (datetime.strptime(str(row[5]), "%H:%M:%S") - datetime.strptime(str(row[4]), "%H:%M:%S")) \
                    - reg_hrs - lunch_break

    print(row[1], timestamp)

