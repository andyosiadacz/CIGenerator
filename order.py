class Order:
    """Order class must be initialized with an order number."""

    def __init__(self, order_num: int):
        self.order_num = order_num
        self.sold_to_name = ""
        self.sold_to_num = ""
        self.sold_to_address = {
            "Name": self.sold_to_name,
            "Address Line 1": "",
            "Address Line 2": "",
            "City": "",
            "State": "",
            "Postal Code": "",
            "Country": "",
        }
        self.ship_to_address = {
            "Name": self.sold_to_name,
            "Address Line 1": "",
            "Address Line 2": "",
            "City": "",
            "State": "",
            "Postal Code": "",
            "Country": "",
        }

        self.po_num = None
        self.orderlines = []
    
    def load_order(self):
        pass

class OrderLine:
    def __init__(self):
        self.line_number : int = ""
        self.item_number = ""
        self.item_description = ""
        self.item_coo = ""
        self.item_hts_code = ""
        self.item_price = ""
        self.quantity = ""
        self.net_weight = ""
        self.gross_weight = ""
        self.cbm = ""