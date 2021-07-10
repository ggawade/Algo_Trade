# import requests
import pandas as pd
from Angel_Strategy import angel_common_functions as f
from datetime import date, datetime
from pprint import pprint
from os import path
import xlwings as xl


# f.generate_nfo_instruments()
nfo_df = f.get_nfo_df()

nfty_details = f.get_exch_sym_tkn_details('NIFTY',nfo_df)

if len(nfty_details) != 4:
    print("ERROR: Not getting Nifty index details! Exiting...")
    exit(10)

obj = f.angel_login()

print(nfty_details)
curr_index = f.get_ltp(obj, nfty_details[0], nfty_details[1], nfty_details[2])

strike_price = f.get_nifty_current_strike_price(curr_index)
expry_date = f.get_expiry_date(date.today())


print(curr_index)
pprint(str(strike_price) + "  " + str(expry_date))

nfo_symbol = f.generate_symbol(strike_price,expry_date)

nfo_list = f.get_instrument_from_symbol(nfo_df, nfo_symbol)

if nfo_list['token'].count() != 2:
    print("ERROR: more than 2 instruments found! exiting...")
    exit(10)

wb = xl.Book("Algo_excel.xlsm")
sht = xl.sheets['Sheet1']
rw = 50
for inst in nfo_list.values:
    xl.Range((rw,5)).value = inst[1] #symbol
    xl.Range((rw, 6)).value = inst[7] #exchange
    xl.Range((rw, 7)).value = inst[0] #token
    xl.Range((rw, 8)).value = inst[5] #lotsize
    rw += 1



