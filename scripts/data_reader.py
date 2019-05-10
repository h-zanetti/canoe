import openpyxl as xl

# df = input('Dataframe path: ')
revenue = input('Revenue path: ')

# wb_df = xl.load_workbook(f'{df}.xlsx')
wb_revenue = xl.load_workbook(f'{revenue}.xlsx', data_only=True)

ws_revenue = wb_revenue.get_sheet_by_name('Revenue Calc 2019')

revenue = {}
for n in range(5, 25):
    revenue[f"{ws_revenue[f'D{n}'].value}"] = round(ws_revenue[f'H{n}'].value)

