import win32com.client, os, sys, time

start_range = 0
stop_range = 1000000


def Percent(current_value):
    percent = (current_value / stop_range) * 100
    return int(percent)


Exel = win32com.client.Dispatch("Excel.Application")
log = open(r'C:\MyPython\log.txt', 'w')
filename = r'C:\MyPython\1.xlsx'
current_time = time.time()
defined_value = ' second'
for password in range(start_range, stop_range):
    try:
        wb = Exel.Workbooks.Open(filename, False, True, None, password)
        log.writelines(str(password) + '\n')
        break
    except:
        print('Progress ' + str(Percent(password)) + '%', end='\r', flush=True)
end_time = time.time()
execute_time = (end_time - current_time)
if execute_time > 3600:
    execute_time = execute_time / 3600
    defined_value = ' hours'
print('Password ', password)
print('Execute time ' + str(execute_time) + defined_value)
log.writelines('Execute time ' + str(execute_time) + defined_value)
log.close()
input()
