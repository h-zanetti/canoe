import openpyxl as xl

wb = xl.Workbook()
ws = wb.active
ws.title = 'Test'

ws['B1'] = 'Hello, world!'

wb.save(filename='test.xlsx')