import angel_common_functions as f
import xlwings as xl
from time import sleep
import os, sys

script_path = os.path.dirname(sys.argv[0])

stop_fname = script_path + '/stop.txt'
stop_loss_fname = script_path + '/stop_loss.txt'

wb = xl.Book(script_path + '/Algo_excel.xlsm')
sht = wb.sheets['Sheet1']
stop_loss = float(sht['Q1'].value)
print("Stop Loss Update: " + str(stop_loss))
obj = f.angel_login()

stop = os.path.isfile(stop_fname)
if stop is True:
    os.remove(stop_fname)

while True:
    sleep(1)

    file_exists = os.path.isfile(stop_loss_fname)
    if file_exists is True:
        try:
            sl_file = open(stop_loss_fname)
            stop_loss = float(sl_file.readline().strip())
            print(str(stop_loss))
            sl_file.close()
            os.remove(stop_loss_fname)
            print("Stop Loss Update: " + str(stop_loss))
        except IOError:
            print("IOError")
        except ValueError:
            print("ValueError")
        except Exception:
            print("Issue")

    stop = os.path.isfile(script_path + '/stop.txt')
    if stop is True:
        print("Exiting!!!")
        break

    try:
        total_pnl = 0
        pos = obj.position()['data']

        for p in pos:
            if p['instrumenttype'] == "OPTIDX":
                try:
                    pnl = float(p['pnl'])
                    total_pnl += pnl
                except Exception as e:
                    break
    except TypeError:
        print("Positions are not available")
    print(total_pnl)

    if total_pnl < stop_loss:
        # order_list = f.exit_all_positions(obj, pos)
        order_list = ()
        print('positions exited!')
        print(order_list)
        break

obj.terminateSession("G53356")
