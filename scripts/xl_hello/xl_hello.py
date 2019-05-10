import openpyxl as xl

wb = xl.Workbook()
ws = wb.active
ws.title = 'Test'

ws['A1'] = 'Hello, world!'

wb.save(filename='test.xlsx')