import pickle
from models import Programmer

programmers = {
    'A&E': {
        'title': 'A&E Networks',
        'total_imp': 271_013_543,#353_923_524,
        'revenue': 429_934,
        'rate_card': {
            'value': 1.13,
            'P4': True
        },
        'networks': ['A&E', 'Lifetime', 'History', 'LMN', 'FYI', 'Viceland'],
        'invoice_num': 8471,
        # invoice map
        'invoice_path': 'plain_invoices/A&E.xlsx',
        'start_point': 28,

    },
    'ABC': {
        'title': 'ABC Disney',
        'total_imp': 1_442_490_479,
        'revenue': 1_306_168,
        'rate_card': {
            'value': 0.71,
            'P4': True
        },
        'networks': ['ABC', 'Disney Junior', 'Freeform', 'Disney Channel', 'Disney XD'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/ABC.xlsx',
        'start_point': 28,

    },
    'AMC': {
        'title': 'AMC Networks',
        'total_imp': 264_501_792,
        'revenue': 328_887,
        'rate_card': {
            'value': 1.13,
            'P4': True
        },
        'networks': ['AMC', 'AMC Premiere', 'AMC Premiere Free', 'IFC', 'Sundance Channel', 'BBC America', 'WE TV', 'Backfill Campaigns'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/AMC.xlsx',
        'start_point': 28,

    },
    'CBS': {
        'title': 'CBS Corporation',
        'total_imp': 452_477_098,
        'revenue': 533_952,
        'rate_card': {
            'value': 0.99,
            'P4': True
        },
        'networks': ['CBS', 'POP TV'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/CBS.xlsx',
        'start_point': 28,

    },
    'CW': {
        'title': 'CW',
        'total_imp': 68_567_706,
        'revenue': 87_767,
        'rate_card': {
            'value': 1.28,
            'P4': True
        },
        'networks': ['CW'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/CW.xlsx',
        'start_point': 28,

    },
    'CROWN': {
        'title': 'Crown Media',
        'total_imp': 1_297_873,
        'revenue': 653,
        'rate_card': {
            'value': 1.42,
            'P4': True
        },
        'networks': ['Hallmark Channel', 'Backfill Campaigns'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/CROWN.xlsx',
        'start_point': 28,

    },
    'DISCOVERY': {
        'title': 'Discovery Networks',
        'total_imp': 989_418_996,
        'revenue': 984_336,
        'rate_card': {
            'value': 0.71,
            'P4': True
        },
        'networks': ['American Heroes Channel', 'Animal Planet', 'Destination America', 'Discovery', 'Discovery en Espanol', 'Discovery Familia', 'Discovery Family Channel', 'Discovery Life', 'Investigation Discovery', 'OWN: Oprah Winfrey Network', 'Science Channel', 'TLC', 'Velocity', 'Cooking Channel', 'DIY Network', 'Food Network', 'HGTV', 'Travel Channel'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/DISCOVERY.xlsx',
        'start_point': 28,

    },
    'FOX': {
        'title': 'Fox Cable Networks',
        'total_imp': 596_248_098, #777_483_966,
        'revenue': 830_861,
        'rate_card': {
            'value': 0.99,#0.85,
            'P4': True
        },
        'networks': ['FOX Broadcast', 'FX', 'FXM', 'FXX', 'National Geographic Channel', 'Nat Geo WILD'],
        'invoice_num': 8479,
        # invoice map
        'invoice_path': 'plain_invoices/FOX.xlsx',
        'start_point': 32,

    },
    'KIDGENIUS': {
        'title': 'Genius Brands',
        'total_imp': 0,
        'revenue': 0,
        'rate_card': {
            'value': 1.05,
            'P4': False
        },
        'networks': ['Kid Genius', 'Baby Genius'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/KIDGENIUS.xlsx',
        'start_point': 27,

    },
    'KABILLION': {
        'title': 'Kabillion',
        'total_imp': 1_550_459,
        'revenue': 1_628,
        'rate_card': {
            'value': 1.05,
            'P4': False
        },
        'networks': ['Kabillion', 'Kabillion Girls Rule'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/KABILLION.xlsx',
        'start_point': 27,

    },
    'EPIX': {
        'title': 'Epix',
        'total_imp': 4_427_094,
        'revenue': 4_648,
        'rate_card': {
            'value': 1.05,
            'P4': False
        },
        'networks': ['Epix'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/EPIX.xlsx',
        'start_point': 27,

    },
    'MC': {
        'title': 'Music Choice',
        'total_imp': 17_589_771,#23_630_905,
        'revenue': 30_248,
        'rate_card': {
            'value': 1.28,
            'P4': True
        },
        'networks': ['Music Choice'],
        'invoice_num': 8483,
        # invoice map
        'invoice_path': 'plain_invoices/MC.xlsx',
        'start_point': 28,

    },
    'NBC': {
        'title': 'NBC Universal',
        'total_imp': 1_282_654_718,#1_801_182_698,
        'revenue': 1_560_840,
        'rate_card': {
            'value': 0.71,
            'P4': True
        },
        'networks': ['Bravo', 'E!', 'NBC Broadcast', 'Oxygen', 'Universal Kids', 'SyFy', 'Telemundo', 'USA', 'NBC Sports', 'NBC News', 'NBC Universo', 'MSNBC', 'CNBC', 'Golf Channel'],
        'invoice_num': 8484,
        # invoice map
        'invoice_path': 'plain_invoices/NBC.xlsx',
        'start_point': 28,

    },
    'REELZ': {
        'title': 'Reelz',
        'total_imp': 0,
        'revenue': 0,
        'rate_card': {
            'value': 1.05,
            'P4': False
        },
        'networks': ['Reelz'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/REELZ.xlsx',
        'start_point': 27,

    },
    'SONY': {
        'title': 'Sony',
        'total_imp': 463_423,
        'revenue': 548,
        'rate_card': {
            'value': 1.42,
            'P4': True
        },
        'networks': ['Cine'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/SONY.xlsx',
        'start_point': 28,

    },
    'STARZ': {
        'title': 'Starz Entertainment',
        'total_imp': 39_204_434,
        'revenue': 41_165,
        'rate_card': {
            'value': 1.05,
            'P4': False
        },
        'networks': ['Starz', 'MoviePlex', 'Starz Encore'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/STARZ.xlsx',
        'start_point': 27,

    },
    'TVONE': {
        'title': 'TV One',
        'total_imp': 477_926,
        'revenue': 502,
        'rate_card': {
            'value': 1.05,
            'P4': False
        },
        'networks': ['TV One'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/TVONE.xlsx',
        'start_point': 27,

    },
    'TURNER': {
        'title': 'Turner',
        'total_imp': 1_208_422_815,
        'revenue': 1_131_461,
        'rate_card': {
            'value': 0.71,
            'P4': True
        },
        'networks': ['truTV', 'Adult Swim', 'TBS', 'Boomerang', 'Cartoon Network', 'Cartoon Network ESP', 'CNN', 'HLN', 'TNT', 'March Madness'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/TURNER.xlsx',
        'start_point': 42,

    },
    'UNIVISION': {
        'title': 'Univision',
        'total_imp': 2_329_012,
        'revenue': 2_773,
        'rate_card': {
            'value': 1.42,
            'P4': False
        },
        'networks': ['Univision', 'Galavision', 'Unimas', 'Univision Deportes', 'El Rey', 'Bandamax', 'TuTv (De Pelicula)'],
        'invoice_num': None,
        # invoice map
        'invoice_path': 'plain_invoices/UNIVISION.xlsx',
        'start_point': 28,

    },
    'VIACOM': {
        'title': 'Viacom',
        'total_imp': 716_752_172,#966_052_759,
        'revenue': 977_897,
        'rate_card': {
            'value': 0.71,
            'P4': True
        },
        'networks': ['Nick Jr (Noggin)', 'Nick Mom', 'Nickelodeon', 'CMT', 'TeenNick', 'BET', 'BET Her', 'MTV', 'MTV2', 'TV Land', 'VH1', 'VH1 Classic', 'Comedy Central', 'Paramount', 'Logo'],
        'invoice_num': 8491,
        # invoice map
        'invoice_path': 'plain_invoices/VIACOM.xlsx',
        'start_point': 28,

    }
}

i=1
for p in programmers:
    new_net = ''
    for n in programmers[p]['networks']:
        new_net += f'{n}, '
    new_net = new_net[:-2]

    if programmers[p]['rate_card']['P4'] == True:
        P4 = 1
    else:
        P4 = 0

    if programmers[p]['invoice_num'] == None:
        programmers[p]['invoice_num'] = 0

    programmer = Programmer(
        i,
        p,
        programmers[p]['title'],
        programmers[p]['total_imp'],
        programmers[p]['revenue'],
        programmers[p]['rate_card']['value'],
        P4,
        new_net,
        programmers[p]['invoice_num'],
        programmers[p]['invoice_path'],
        programmers[p]['start_point']
    )
    
    prog_file = open(f'programmers/{p}.p', 'wb')
    pickle.dump(programmer, prog_file)

    i+=1