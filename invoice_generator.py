import openpyxl as xl
import datetime as dt
from colorama import Fore

# Immutable Data

invoice_map = {
# Header
    'date': 'L1', 'invoice_num': 'L2', 'period_start': 'D18', 'period_end': 'D19', 'previous_YTD_imp': 'D22',
# Body
    'YTD_imp': 'J', 'campaign_id': 'C', 'campaign_name': 'D', 'network': 'E', 'start_date': 'F', 'end_date': 'G', 'total_imp': 'H',
    'month_imp': 'I', 'cpm':  'J', 'total': 'K',
}

df_map = {
    'campaign_id': 'A', 'campaign_name': 'B', 'network': 'C',
    'start_date': 'D', 'end_date': 'E',
    'campaign_goal': 'F', 'total_imp': 'G', 'month_imp': 'H'
}

today = dt.date.today()
programmers = {
    'NBC Universal': None
    # 'ABC Disnay': 'P4',
    # 'AMC Networks': 'P4',
    # 'CBS Corporation': 'P4',
    # 'CW': 'P4',
    # 'Epix': None,
    # 'Genius Brands': None,
    # 'Kabillion': None,
    # 'Reelz': None,
    # 'Starz Entertainment': None,
    # 'TV One': None
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
    invoice_start = int(input("Column number where the data will start to be placed: "))
    df_length = int(input('Dataframe length: '))

    invoice_map['data_starts'] = f"C{invoice_start}"
    invoice_map['data_ends'] = f"J{invoice_start}"

    try:
        ws_df = wb_df.get_sheet_by_name(programmer)
        print(Fore.WHITE + '[' + Fore.GREEN + "OK" + Fore.WHITE + '] Data successfully collected')

        old_invoice_path = input(f"Path of an old invoice for {programmer}: ")
        marketplace = int(input('How many marketplace fields are in the invoice? '))

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

        # incomplete, make it change the rate card when needed
        for r in range(17, 26):
            if ws_invoice.cell(row=r, column=10).value != None:
                cpm = ws_invoice.cell(row=r, column=9).value
                ws_invoice[invoice_map['previous_YTD_imp']] = ws_invoice_data.cell(row=r, column=10).value

        ws_invoice[invoice_map['period_start']] = time_period_start
        ws_invoice[invoice_map['period_end']] = time_period_end

        # Update data
        i = 1
        networks = {}
        for n in range(invoice_start, invoice_start + df_length):
            if i == ws_invoice_data[f'B{n}'].value and ws_invoice_data[f"{invoice_map['campaign_id']}{n}"].value != 'NA':
                for key in df_map:
                    ws_invoice[f"{invoice_map[key]}{n}"] = ws_df[f"{df_map[key]}{i+3}"].value
            else:
                ws_invoice.insert_rows(n)
                ws_invoice[f'B{n}'] = f'=B{n - 1}+1'
                for key in df_map:
                    ws_invoice[f"{invoice_map[key]}{n}"] = ws_df[f"{df_map[key]}{i+3}"].value
            
            # Check for multiple networks
            if ws_df[f"{df_map['network']}{i+3}"].value in networks:
                networks[ws_df[f"{df_map['network']}{i+3}"].value] += ws_df[f"{df_map['month_imp']}{i+3}"].value
            else:
                networks[ws_df[f"{df_map['network']}{i+3}"].value] = ws_df[f"{df_map['month_imp']}{i+3}"].value
            
            i += 1

    # Formatting corrections
        max_row = ws_invoice.max_row
        max_column = ws_invoice.max_column
        for r in range(1, max_row+1):
            if ws_invoice_data.cell(row=r, column=8).value in networks:
                ws_invoice.cell(row=r, column=9).value = networks[ws_invoice.cell(row=r, column=8).value]
                ws_invoice.cell(row=r, column=11).value = networks[ws_invoice.cell(row=r, column=8).value] * cpm
            
            if ws_invoice_data.cell(row=r, column=7).value == 'Total:':
                ws_invoice.cell(row=r, column=9).value = f"=SUM({invoice_map['month_imp']}{invoice_start}:{invoice_map['month_imp']}{invoice_start + df_length - 1})"
                ws_invoice.cell(row=r, column=11).value = f"=SUM({invoice_map['total']}{invoice_start}:{invoice_map['total']}{invoice_start + df_length - 1})"

            if ws_invoice_data.cell(row=r, column=10).value in networks:
                ws_invoice.cell(row=r, column=11).value = networks[ws_invoice.cell(row=r, column=10).value] * cpm
            
            if ws_invoice_data.cell(row=r, column=10).value == 'Amount Due:':
                ws_invoice.cell(row=r, column=11).value == f"=SUM({invoice_map['total']}{invoice_start}:{invoice_map['total']}{invoice_start + df_length - 1})"

        wb_invoice.save(f"new_invoices/{file_name}")
        print(Fore.WHITE + '[' + Fore.GREEN + "SUCCESS" + Fore.WHITE + f"] {programmer} invoice created")

    except:
        print('[' + Fore.RED + "ERROR" + Fore.WHITE + f'] Fail to generate {programmer} invoice')
