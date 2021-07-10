from datetime import datetime, date
from dateutil.relativedelta import relativedelta, TH
from smartapi import SmartConnect
from pprint import pprint
import requests
import pandas as pd
import numpy as np
import xlwings as xl
from time import sleep


def angel_login():
    wb = xl.Book("Algo_excel.xlsm")
    sht = wb.sheets['CONFIG']

    API_KEY = sht['B3'].value  # tGA5Z6GH
    CLIENT_ID = sht['B1'].value  # "G53356"
    PASSWORD = sht['B2'].value  # "victor3"

    # sht['D1'].value = "TEST SUCCESSFUL"

    obj = SmartConnect(api_key=API_KEY)
    data = obj.generateSession(CLIENT_ID, PASSWORD)
    refreshToken = data['data']['refreshToken']
    sht['B5'].value = str(refreshToken)
    feedToken = obj.getfeedToken()
    sht['B6'].value = str(feedToken)
    return obj



def get_expiry_date(run_date):
    curr_exp_date = run_date + relativedelta(weekday=TH(1))
    return curr_exp_date


def get_nifty_current_strike_price(curr_idx):
    mod = curr_idx % 50
    cur_strk_price = 0
    if mod < 25:
        cur_strk_price = curr_idx - mod
    else:
        cur_strk_price = curr_idx + (50 - mod)

    return cur_strk_price


def generate_nfo_instruments():
    url = 'https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json'

    r = requests.get(url)

    data = r.json()

    inst_df = pd.json_normalize(data).convert_dtypes()

    nfo_df = inst_df.loc[
        ((inst_df['exch_seg'] == 'NFO') | (inst_df['symbol'].isin(['NIFTY']))) & (inst_df['name'] == 'NIFTY')]
    bk_df = inst_df.loc[(inst_df['symbol'] == 'BANKNIFTY')]

    nfo_df = nfo_df.append(bk_df)

    nfo_df.to_csv('nfo_data.csv', index=False)


def get_nfo_df():
    nfo_df = pd.read_csv('C:/Users/lenovo/Desktop/Algo_test/Angel_Strategy/nfo_data.csv')
    return nfo_df


def get_exch_sym_tkn_details(symbol, nfo_df):
    ret_list = list()
    df = nfo_df.loc[nfo_df['symbol'] == symbol]
    if df['token'].count() == 1:
        ret_list.append(str(df.iloc[0]['exch_seg']))  # exchange
        ret_list.append(str(df.iloc[0]['symbol']))  # symbol
        ret_list.append(str(df.iloc[0]['token']))  # token
        ret_list.append(str(df.iloc[0]['lotsize']))  # lot size
    return ret_list


def get_ltp(obj, exchange, symbol, symtoken):
    data = obj.ltpData(exchange, symbol, symtoken)
    if data['status'] is False:
        ltp = "-1"
    else:
        ltp = data['data']['ltp']
    return ltp


def generate_symbol(strk, expry_dt):
    my_time = datetime.min.time()
    dt = datetime.combine(expry_dt, my_time)
    e_dt = dt.strftime('%d%b%y').upper()
    nfo_symbol = "NIFTY" + e_dt + str(int(strk))
    return nfo_symbol


def get_instrument_from_symbol(df, symbol):
    return df.loc[df['symbol'].str.contains(symbol)]


def generate_order_param(symbol, symbol_token, trans_type, exch, qty):
    order_param = {
        "variety": "NORMAL",
        "tradingsymbol": symbol,
        "symboltoken": symbol_token,
        "transactiontype": trans_type,
        "exchange": exch,
        "ordertype": "MARKET",
        "producttype": "INTRADAY",
        "duration": "DAY",
        "price": "0",
        "squareoff": "0",
        "stoploss": "0",
        "quantity": qty
    }
    return order_param


def place_Order(obj, order_param):
    order_data = obj.placeOrder(order_param)
    if order_data['status'] is True:
        order_id = order_data['data']['orderid']
    else:
        order_id = '-1'
    return order_id


def get_order_book(obj):
    order_book = obj.orderBook()['data']
    return order_book


def exit_all_positions(obj, pos):
    order_list=list()
    for p in pos:
        if p['instrumenttype'] == "OPTIDX":
            symbol = p['tradingsymbol']
            symbol_token = p['symboltoken']
            exch = p['exchange']
            net_qty = int(p['netqty'])
            if net_qty < 0:
                trans_type='BUY'
                qty = abs(net_qty)
                order_param = generate_order_param(symbol, symbol_token, trans_type, exch, qty)
                order_id=place_Order(obj, order_param)
                order_list.append(order_id)

    return order_list
