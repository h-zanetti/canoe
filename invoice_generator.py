import openpyxl as xl
import datetime as dt
from colorama import Fore

# Immutable Data

invoice_map = {
# Header
    'date': 'L1', 'invoice_num': 'L2', 'period_start': 'D18', 'period_end': 'D19', 'previous_YTD_imp': 'D22', 'YTD_imp': 'K17',
# Body
    'campaign_id': 'C', 'campaign_name': 'D', 'network': 'E', 'start_date': 'F', 'end_date': 'G', 'campaign_goal': 'H', 'total_imp': 'I',
    'month_imp': 'J', 'cpm':  'K', 'total': 'L',
}

df_map = {
    'campaign_id': 'A', 'campaign_name': 'B', 'network': 'C',
    'start_date': 'D', 'end_date': 'E',
    'campaign_goal': 'F', 'total_imp': 'G', 'month_imp': 'H'
}

today = dt.date.today()
programmers = {
    # 'ABC Disnay': 'P4',
    # 'AMC Networks': 'P4',
    # 'CBS Corporation': 'P4',
    'CW': 'P4',
    'Epix': None,
    'Genius Brands': None,
    'Kabillion': None,
    'Reelz': None,
    'Starz Entertainment': None,
    'TV One': None
}

time_period_start = input('Time period start (mm/dd/yyyy): ')
time_period_end = input('Time period end (mm/dd/yyyy): ')

df_path = input('Dataframe path: ')
revenue_path = input('Revenue path: ')

print("Collecting data...")
# Revenue Workbook
wb_revenue = xl.load_workbook(f'{revenue_path}', data_only=True)
# Revenue Sheets
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

wb_df = xl.load_workbook(f'{df_path}', data_only=True)
for programmer in programmers:
    print(f"Creating invoice for {programmer}")
    # invoice_start = input("Column number where the data will start to be placed: ")
    # invoice_start = int(invoice_start)
    if programmer == 'CW':
        invoice_start = 28
    else:
        invoice_start = 27

    df_length = input('Dataframe length: ')
    df_length = int(df_length)

    # All invoices from CANOE_INVOICE_APR_2019 starts at the column 27, for some weird reason CW starts at 28
    invoice_map['data_starts'] = f"C{invoice_start}"
    invoice_map['data_ends'] = f"J{invoice_start}"

    try:
        ws_df = wb_df.get_sheet_by_name(programmer)
        print(Fore.WHITE + '[' + Fore.GREEN + "OK" + Fore.WHITE + '] Data successfully collected')

        old_invoice_path = input(f"Path of an old invoice for {programmer}: ")

        file_name_date = today.isoformat().replace('-', '')
        file_name_programmer = programmer.replace(' ', '_')
        if programmers[programmer] == 'P4':
            file_name = f'invoice-{file_name_programmer}-{file_name_date}_P4.xlsx'
        else:
            file_name = f'invoice-{file_name_programmer}-{file_name_date}.xlsx'

        print(f'Creating {file_name}')

        wb_invoice = xl.load_workbook(old_invoice_path)
        ws_invoice = wb_invoice.get_sheet_by_name('Invoice')

        wb_invoice_data = xl.load_workbook(old_invoice_path, data_only=True)
        ws_invoice_data = wb_invoice_data.get_sheet_by_name('Invoice')

        # Update header
        ws_invoice[invoice_map['date']] = today.strftime('%m/%d/%Y')

        ws_invoice[invoice_map['invoice_num']] = invoice_numbers[programmer]

        ws_invoice[invoice_map['previous_YTD_imp']] = ws_invoice_data[invoice_map['YTD_imp']].value # incomplete, make it change the rate card when needed

        ws_invoice[invoice_map['period_start']] = time_period_start
        ws_invoice[invoice_map['period_end']] = time_period_end

        # Update data
        i = 1
        for n in range(invoice_start, invoice_start + df_length):
            if i == ws_invoice_data[f'B{n}'].value:
                for key in df_map:
                        ws_invoice[f"{invoice_map[key]}{n}"] = ws_df[f"{df_map[key]}{i+3}"].value
            else:
                ws_invoice.insert_rows(n)
                ws_invoice[f'B{n}'] = f'=B{n - 1}+1'
                for key in df_map:
                        ws_invoice[f"{invoice_map[key]}{n}"] = ws_df[f"{df_map[key]}{i+3}"].value
            i += 1

        # Formatting corrections
        # Still have to change the column numbers - most of invoices have different positions due to multiple networks on the Sub-total field
        # This will work for sub-totals with no networks
        tota_imp = f'=SUM(J{invoice_start}:J{df_length + invoice_start - 1})'
        amount_due = f'=SUM(L{invoice_start}:L{df_length + invoice_start - 1})'
        new_YTD_imp = f"={invoice_map['previous_YTD_imp']}+J{df_length + invoice_start - 2}"

        ws_invoice[f"{invoice_map['YTD_imp']}"] = new_YTD_imp # Changing YTD impressions
        ws_invoice[f"{invoice_map['month_imp']}{df_length + invoice_start - 2}"] = tota_imp # Changing total impressions of the month
        ws_invoice[f"{invoice_map['total']}{df_length + invoice_start - 2}"] = amount_due # Changing total of the month
        ws_invoice[f"{invoice_map['total']}{df_length + invoice_start - 13}"] = f"={invoice_map['total']}{df_length + invoice_start - 2}" # Changing amount due

        wb_invoice.save(filename=file_name)
        print(Fore.WHITE + '[' + Fore.GREEN + "SUCCESS" + Fore.WHITE + f"] {programmer} invoice created")

    except:
        print('[' + Fore.RED + "ERROR" + Fore.WHITE + f'] Fail to generate {programmer} invoice')










