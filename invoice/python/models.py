class Programmer:
    def __init__(self, id, abbreviation, title, total_imp, revenue, rate_card, P4, networks, invoice_num, invoice_path, start_point):
        self.id = id
        self.abbreviation = abbreviation
        self.title = title
        self.total_imp = total_imp
        self.revenue = revenue
        self.rate_card = rate_card
        self.P4 = P4
        self.networks = networks
        self.invoice_num = invoice_num
        self.invoice_path = invoice_path
        self.start_point = start_point

    def __getitem__(self, network):
        return self.networks[network]
    
    def __str__(self):
        return f"{self.abbreviation}"

    def insert(self, table_name):
        dic = {
            'id': self.id,
            'abbreviation': self.abbreviation,
            'title': self.title,
            'total_imp': self.total_imp,
            'revenue': self.revenue,
            'rate_card': self.rate_card,
            'P4': self.P4,
            'networks': self.networks,
            'invoice_num': self.invoice_num,
            'invoice_path': self.invoice_path,
            'start_point': self.start_point
        }
        string = ['abbreviation', 'title', 'networks', 'invoice_path']

        insert = f'INSERT INTO {table_name} ('
        for key in dic:
            insert += f'{key}, '
        insert = insert[:-2]
        insert += ') VALUES ('
        for key in dic:
            if key in string:
                insert += f"'{dic[key]}', "
            else:
                insert += f'{dic[key]}, '
        insert = insert[:-2]
        insert += ')'

        return insert

    def update(self, table_name, total_imp, revenue, rate_card, invoice_num):
        dic = {
        'total_imp': total_imp,
        'revenue': revenue,
        'rate_card': rate_card,
        'invoice_num': invoice_num,
        }

        update = f'UPDATE {table_name} SET '
        for key in dic:
            update += f'{key} = {dic[key]}, '
        update = update[:-2]
        update += f' WHERE id = {self.id}'

        return update

