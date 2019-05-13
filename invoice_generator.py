import openpyxl as xl
import datetime as dt

today = dt.date.today()

time_period_start = input('Time period start (mm/dd/yyyy): ')
time_period_end = input('Time period end (mm/dd/yyyy): ')
df_path = input('Dataframe path: ')
df_length = int(input('Dataframe length: '))
revenue_path = input('Revenue path: ')

print("Collecting data...")
wb_df = xl.load_workbook(f'{df_path}', data_only=True)
wb_revenue = xl.load_workbook(f'{revenue_path}', data_only=True)

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

# loop over programmers and create an invoice for each of them
# for programmer in programmers:
#     ws_invoice.title = programmer

# Map invoice locations and columns
invoice_map = {
# Header
    'date': 'L1', 'invoice_num': 'L2', 'period_start': 'D18', 'period_end': 'D19', 'previous_YTD_imp': 'D22', 'YTD_imp': 'K17',

# Body
    # All invoices from CANOE_INVOICE_APR_2019 starts at the column 27, for some weird reason CW starts at 28
    'data_starts': 'C28', 'data_ends': 'J28',
    # Data columns
    'campaign_id': 'C', 'campaign_name': 'D', 'network': 'E', 'start_date': 'F', 'end_date': 'G', 'campaign_goal': 'H', 'total_imp': 'I',
    'month_imp': 'J', 'cpm':  'K', 'total': 'L',
}

# Dataframe Column Map
df_map = {
    'campaign_id': 'A', 'campaign_name': 'B', 'network': 'C',
    'start_date': 'D', 'end_date': 'E',
    'campaign_goal': 'F', 'total_imp': 'G', 'month_imp': 'H'
}

print("[OK]")

file_name_date = today.isoformat().replace('-', '')
file_name = f'invoice-{programmers[0]}-{file_name_date}_P4.xlsx'
print(f'Creating {file_name}')

# Load CW dataframe
ws_CW_df = wb_df.get_sheet_by_name('CW')

# Load last CW invoice
wb_old_CW = xl.load_workbook('../INVOICE_MAR_2019/invoice-CW-20190401_P4.xlsx')
ws_old_CW = wb_old_CW.get_sheet_by_name('Invoice')

wb_old_CW_data = xl.load_workbook('../INVOICE_MAR_2019/invoice-CW-20190401_P4.xlsx', data_only=True)
ws_old_CW_data = wb_old_CW_data.get_sheet_by_name('Invoice')

# Change header
today.strftime('%m/%d/%Y')
ws_old_CW[invoice_map['date']] = today

ws_old_CW[invoice_map['invoice_num']] = invoice_numbers['CW']

ws_old_CW[invoice_map['previous_YTD_imp']] = ws_old_CW_data[invoice_map['YTD_imp']].value # incomplete

ws_old_CW[invoice_map['period_start']] = time_period_start
ws_old_CW[invoice_map['period_end']] = time_period_end

# Past data
i = 1
for n in range(28, df_length + 28):
    if i == ws_old_CW_data[f'B{n}'].value:
        for key in df_map:
                ws_old_CW[f"{invoice_map[key]}{n}"] = ws_CW_df[f"{df_map[key]}{i+3}"].value
    else:
        ws_old_CW.insert_rows(n)
        ws_old_CW[f'B{n}'] = f'=B{n - 1}+1'
        for key in df_map:
                ws_old_CW[f"{invoice_map[key]}{n}"] = ws_CW_df[f"{df_map[key]}{i+3}"].value
    i += 1

print('[OK]')

print('Making format corrections...')

tota_imp = f'=SUM(J28:J{df_length+28-1})'
amount_due = f'=SUM(L28:L{df_length+28-1})'
new_YTD_imp = f"={invoice_map['previous_YTD_imp']}+J{df_length+28+2}"

ws_old_CW[f"{invoice_map['YTD_imp']}"] = new_YTD_imp # Changing YTD impressions
ws_old_CW[f"{invoice_map['month_imp']}{df_length+28+2}"] = tota_imp # Changing total impressions of the month
ws_old_CW[f"{invoice_map['total']}{df_length+28+2}"] = amount_due # Changing total of the month
ws_old_CW[f"{invoice_map['total']}{df_length+28+13}"] = f"={invoice_map['total']}{df_length+28+2}" # Changing amount due

wb_old_CW.save(filename=file_name)
print('[DONE]')