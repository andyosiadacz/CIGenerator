import pandas as pd
from order import *

df = pd.read_csv('practice.csv')

# ROW COUNTER
row = 0

order_num = None
orders = []

for num in df['Order Number']:
    row_data = df.loc[row]
    
    # CREATE NEW ORDER FOR NEW ORDER NUMBER AND LOAD HEADER DATA
    if order_num != num:
        order = Order(num)


        # LOAD ORDER HEADER DATA
        # order.sold_to_name = row_data['Sold To']
        order.sold_to_num = row_data['Sold To Number']
        order.po_num = row_data['Customer PO']

        orders.append(order)

    # INITIALIZE ORDERLINE, SAVE DATA FROM DF, AND APPEND ORDER LINE
    line = OrderLine()
    
    line.item_number = row_data['2nd Item Number']
    line.item_description = row_data['Concatenation Description ']
    line.item_coo = None
    line.item_hts_code = None
    line.item_price = row_data['Unit Price']
    line.quantity = row_data['Quantity Shipped']
    line.net_weight = None
    line.gross_weight = None
    line.cbm = None

    order.orderlines.append(line)  

    # INCREMENT ROW   
    row += 1

