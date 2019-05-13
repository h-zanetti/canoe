import openpyxl as xl
import datetime as dt

date = dt.date.today().isoformat().replace('-', '') # Create string for the filename 'yyyymmdd'
programer = 'ABC'

wb = xl.Workbook()
ws = wb.active
ws.title = f'{programer}'

wb.save(filename=f'invoice-{programer}-{date}.xlsx')