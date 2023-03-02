import os
import openpyxl
from ping3 import ping

# Open the Excel file and select the sheet containing the server names
workbook = openpyxl.load_workbook('servers3.xlsx')
sheet = workbook['Sheet1']

# Loop through each row in the sheet, and ping each server in turn
values = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    values.append(row[0])

for i in range(len(values)):
    if values[i] is not None:
        print(values[i])
        print(ping(values[i]))

"""
    server_name = row[0]
    response_time = ping(server_name)

    if response_time is not False or None:
        print(f'{server_name} responded in {response_time} ms')
    else:
        print(f'{server_name} is down')

    print(row)
    """