import openpyxl as xl

df = input('Dataframe path: ')
revenue = input('Revenue path: ')

wb_df = xl.load_workbook(f'{df}.xlsx')
wb_revenue = xl.load_workbook(f'{revenue}.xlsx', data_only=True)

wb_invoice = xl.Wrokbook()
ws_invoice = wb_invoice.active

ws_revenue = wb_revenue.get_sheet_by_name('Revenue Calc 2019')
ws_impressions = wb_revenue.get_sheet_by_name('2019 Impressions')
ws_invoice_num = wb_revenue.get_sheet_by_name('Invoice Numbers')

invoice_numbers = {}
for n in range(3, 25):
    invoice_numbers[f"{ws_invoice_num[f'B{n}'].value}"] = ws_invoice_num[f'F{n}'].value

impressions = {}
for n in range(3, 23):
    impressions[f"{ws_impressions[f'B{n}'].value}"] = ws_impressions[f'F{n}'].value

revenue = {}
for n in range(5, 25):
    revenue[f"{ws_revenue[f'D{n}'].value}"] = round(ws_revenue[f'H{n}'].value)

# programmers = []
# for programmer in impressions:
#     programmers.append(programmer)

programmers = ['CW', 'Epix', 'Genius Brands', 'Kabillion' , 'Reelz' , 'Starz Entertainment', 'TV One']

for programmer in programmers:
    ws_invoice.title = programmer