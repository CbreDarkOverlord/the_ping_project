import openpyxl
from ping3 import ping

# Open the Excel file and select the sheet containing the server names
workbook = openpyxl.load_workbook('example.xlsx')
sheet = workbook['Sheet1']

# Loop through each row in the sheet, and ping each server in turn
values = []
for row in sheet.iter_rows(min_row=1, values_only=True):
    values.append(row[0])

for i in range(len(values)):
    server_name = values[i]
    print(f'{server_name}')
    if ping(server_name) is None or ping(server_name) is False:
        print(f'{server_name} is down')
        sheet.cell(row=i+1, column=2).value = 'Down'
    else:
        print(f'{server_name} responded in {ping(server_name)} ms')
        sheet.cell(row=i+1, column=2).value = 'Up'

# save all the hawtness to the spreadsheet and crack a cold one mate, yer job's done
workbook.save("example.xlsx")