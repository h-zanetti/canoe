from config import programmers

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

# i=0
# for p in programmers:
#     programmer = Programmer(
#         i,
#         p.title(),
#         programmers[p]['title'],
#         programmers[p]['total_imp'],
#         programmers[p]['revenue'],
#         programmers[p]['rate_card']['value'],
#         programmers[p]['rate_card']['P4'],
#         programmers[p]['networks'],
#         programmers[p]['invoice_num'],
#         programmers[p]['invoice_path'],
#         programmers[p]['start_point'],
#     )
#     i+=1