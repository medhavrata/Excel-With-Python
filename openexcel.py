import openpyxl

# Open the workbook via load_workbook
wb = openpyxl.load_workbook('example.xlsx')

print(type(wb))

print(wb.sheetnames) # it will print all the sheetnames of workbook

sheet = wb['Sheet1'] # to access a worksheet

print(sheet['A1'].value) # get value of a cell

c = sheet['B1']
print(c.value)

sheet.cell(row = 10, column = 10).value = 55 # update a cell value
wb.save('example.xlsx')

print(sheet.title)

activeSheet = wb.active  # to access the active sheet

