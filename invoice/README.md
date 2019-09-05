# Invoice Generator

## Instructions:

The script will look for a Excel file (.xlsx) in the following directory: 
`invoice/data/<file_name>.xlsx`.
 
It will predict the file name by the month in wich the invoice is running (i.g. if you are 
making June invoices it will look for a dataframe like this: `invoice/data/INVOICE_JUN.xlsx`

Once it has the dataframe, it will look for the base invoices in the following directory:
`invoice/plain_invoices/PLAIN_INVOICE.xlsx`

With that in place, just watch the magic. The script will print into the commad line messages
of [OK] or [ALERT] if the total impressions of a given invoice matches with the total 
impressions in the database (OPERATIONS.ops_stat_all)

Finally, all invoices will be saved in one Excel file in the following directory:
`invoice/new_invoices/INVOICES_MONTH.xlsx`
