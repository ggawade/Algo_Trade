import angel_common_functions as f
import pandas as pd
import Algo_excel as a
import xlwings as xl

obj = f.angel_login()
order_param = f.generate_order_param('NIFTY08JUL2116500CE','45003','BUY','NFO','75')
order_id = f.place_Order(obj,order_param)
# if order_data['status'] is True:
#     order_id = order_data['data']['orderid']
# else:
#     order_id = '-1'

print(order_id)
obj.terminateSession('G53356')

# print(order_id):