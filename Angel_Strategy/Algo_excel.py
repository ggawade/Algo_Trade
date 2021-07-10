import angel_common_functions as f
from xlpython import xlfunc
import pandas as pd


# nfo_df = f.get_nfo_df()
# sym_list = f.get_exch_sym_tkn_details()

@xlfunc
# @xlarg("symbol", "nparray", 2)
def get_symbol_details(symbol):
    nfo_df = f.get_nfo_df()
    sym_list = f.get_exch_sym_tkn_details(symbol, nfo_df)
    ret_str = ""
    for i in sym_list:
        ret_str = ret_str + str(i) + ","
    return ret_str


@xlfunc
def generate_nifty_nfo_instruments():
    f.generate_nfo_instruments()
    return "Done!"


@xlfunc
def place_order(symbol, symbol_token, trans_type, exch, qty, client_id):
    update_time = ""
    order_status = ""
    reject_reason = ""

    obj = f.angel_login()
    symbol_token = str(int(symbol_token))
    qty = str(int(qty))
    symbol = str(symbol)
    trans_type = str(trans_type)
    exch = str(exch)
    order_param = f.generate_order_param(symbol, symbol_token, trans_type, exch, qty)
    # print(order_param)
    # order_id = ""
    order_id = f.place_Order(obj, order_param)
    # order_status = "success"
    # traded_price = '1.0'
    if order_id != '-1':
        order_book = f.get_order_book(obj)
        order_df = pd.DataFrame(order_book)
        order_details = order_df[order_df['orderid'] == order_id]
        traded_price = '0.0'
        if order_details['orderid'].count() == 1:
            traded_price = str(order_details.iloc[0]['averageprice'])
            order_status = str(order_details.iloc[0]['orderstatus'])
            reject_reason = str(order_details.iloc[0]['text'])
            update_time = str(order_details.iloc[0]['updatetime'])
    else:
        order_status = "Failed"
        traded_price = '-1'
    obj.terminateSession(client_id)
    return str(order_id) + "," + order_status + "," + traded_price + "," + reject_reason + "," + update_time


@xlfunc
def get_ltp(exchange, symbol, symtoken, client_id):
    ltp_val = 0.0
    obj = f.angel_login()
    ltp_val = f.get_ltp(obj, exchange, symbol, symtoken)
    obj.terminateSession(client_id)
    return str(ltp_val)
