class Bill:
    period_str: str
    bill_nbr: int
    customer_nickname: str
    customer_name: str
    customer_address: str
    customer_zip_place: str
    nbr_of_eggs: int
    price_per_egg: float = 0.6

    def __init__(self, period_str: str, bill_nbr: int, customer_nickname: str, information_path: str, nbr_of_eggs: int):
        self.period_str = period_str
        self.bill_nbr = bill_nbr
        self.customer_nickname = customer_nickname
        self.get_postal_information(information_path, customer_nickname)
        self.nbr_of_eggs = nbr_of_eggs

    def get_postal_information(self, information_path: str, customer_name: str):
        information_path = information_path if information_path[-1] == "/" else information_path + "/"
        with open(information_path + customer_name + ".txt", 'r', encoding='utf-8') as f:
            self.customer_name = f.readline()
            self.customer_address = f.readline()
            self.customer_zip_place = f.readline()

    def get_bill_nbr_string(self) -> str:
        return str(self.bill_nbr).zfill(3)
