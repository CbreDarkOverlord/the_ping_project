import openpyxl
from ping3 import ping

# Open the Excel file and select the sheet containing the server names
workbook = openpyxl.load_workbook('example.xlsx')
sheet = workbook['Sheet1']

# Loop through each row in the sheet, and ping each server in turn
values = [row[0] for row in sheet.iter_rows(min_row=1, values_only=True)]

# Loop through each server and ping it once
rows = []
for i, server_name in enumerate(values):
    print(f'{server_name}')
    response_time = ping(server_name)
    status = 'Down' if response_time is None or response_time is False else 'Up'
    rows.append((server_name, status, response_time))
    print(f'{server_name} is {status}')

# Write the results to the worksheet using batch cell writing
sheet.append(['Server Name', 'Status', 'Response Time (ms)'])
for row in rows:
    sheet.append(row)

# save all the hawtness to the spreadsheet and crack a cold one mate, yer job's done
workbook.save("example.xlsx")