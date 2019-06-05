import datetime as dt
import calendar
from oracle import *

today = dt.date.today()
start_date = dt.date(today.year, today.month-1, 1)
end_date = dt.date(today.year, today.month-1, calendar.monthrange(today.year, today.month-1)[1])

invoice_map = {
    # Header
    'date': 'K1', 'invoice_num': 'K2', 'period_start': 'D18', 'period_end': 'D19', 'programmer': 'D20', 'networks': 'D21', 'previous_YTD_imp': 'D22', 'rate_cards': 'I',
    # Body
    'YTD_imp': 'J', 'campaign_id': 'C', 'campaign_name': 'D', 'network': 'E', 'start_date': 'F', 'end_date': 'G', 'total_imp': 'H',
    'month_imp': 'I', 'cpm':  'J', 'total': 'K',
}
df_map = {
    'campaign_id': 'A', 'campaign_name': 'B', 'network': 'C',
    'start_date': 'D', 'end_date': 'E',
    'total_imp': 'F', 'month_imp': 'G'
}

# Programmers Info
query = '''
SELECT id, abbreviation, title, total_imp, rate_card, networks, invoice_num, start_point, bill_to, attention, address, state, contact
FROM invoice_generator_info
ORDER BY TITLE
'''
cur = cursor.execute(query)
programmers = cur.fetchall()
# Total Impressions
query = f'''
SELECT os.PROGRAMMER, sum(os.IMPRESSIONS) FROM OPERATIONS.OPS_STAT_ALL os
WHERE os.EVENT_DATE >= '01-JAN-{today.strftime('%y')}'
AND os.EVENT_DATE <= '{end_date.strftime('%d-%b-%y').upper()}'
GROUP BY os.PROGRAMMER ORDER BY 1
'''
cur = cursor.execute(query)
impressions = cur.fetchall()

rate_card = {
    199999999: {
        'value_1': 0.95,
        'value_2': 1.05,
        'P4_1': 1.28,
        'P4_2':1.42
    },
    399999999: {
        'value_1': 0.84,
        'value_2': 1.00,
        'P4_1': 1.13,
        'P4_2': 1.35
    },
    599999999: {
        'value_1': 0.74,
        'value_2': 0.95,
        'P4_1': 0.99,
        'P4_2': 1.28
    },
    799999999: {
        'value_1': 0.63,
        'value_2': 0.89,
        'P4_1': 0.85,
        'P4_2': 1.21
    }, 
    1999999999: {
        'value_1': 0.53,
        'value_2': 0.84,
        'P4_1': 0.71,
        'P4_2': 1.13
    },
    2999999999: {
        'value_1': 0.45,
        'value_2': 0.79,
        'P4_1': 0.61,
        'P4_2': 1.06
    },
    3999999999: {
        'value_1': 0.43,
        'value_2': 0.76,
        'P4_1': 0.58,
        'P4_2': 1.03
    },
    4999999999: {
        'value_1': 0.41,
        'value_2': 0.73,
        'P4_1': 0.55,
        'P4_2': 0.99
    },
    99999999999: {
        'value_1': 0.41,
        'value_2': 0.73,
        'P4_1': 0.50,
        'P4_2': 0.94
    }
}
