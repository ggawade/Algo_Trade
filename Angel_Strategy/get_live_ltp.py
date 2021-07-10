import xlwings as xl
import angel_common_functions as f
from time import sleep

wb = xl.Book("Algo_excel.xlsm")
sht = wb.sheets['Sheet1']
xl.App.screen_updating = False
obj = f.angel_login()
sht['N2'].value = ""

while True:
    var_stop = wb.sheets['CONFIG']['G1'].value
    for i in range(5, 25):
        if sht['M' + str(i)].value != "" and sht['M' + str(i)].value is not None:
            symbol = sht['E' + str(i)].value
            symtoken = sht['G' + str(i)].value
            exch = sht['F' + str(i)].value
            # sht['N'+str(i)].value = symbol + "  :  " + symtoken
            ltp = f.get_ltp(obj, exch, symbol, symtoken)
            sht['N' + str(i)].value = ltp
            total_pnl = sht['N1'].value
            if total_pnl <= int(sht['Q1'].value):
                # f.exit_all_positions(obj)
                sht['N2'].value = "EXIT"
                var_stop = 'STOP'
                break

    if var_stop == "STOP":
        break

    sleep(1)

print("Completed")
