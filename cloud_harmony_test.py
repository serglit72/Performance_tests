import csv
import io
import sys
import os
import time

import subprocess
import gspread
import pyautogui
import urllib3
import json
from pywinauto.application import Application
from pywinauto import Desktop
from datetime import timedelta

# curl -w "@format" -o /dev/null -s "https://nytimes.com"

global gs_file
gs_file = 'CloudHarmony_test_results'

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 2
# pyautogui.displayMousePosition()
# print(pyautogui.size())

def util():
    # path = "C:\\Users\\i.mamutkina\\Downloads\\Performance_tests\\bin\\"
    cmd = "C:\\Users\\i.mamutkina\\Downloads\\Performance_tests\\bin\\NetworkMeasurement.exe"
    run_util = subprocess.Popen(cmd, stdout=subprocess.PIPE,universal_newlines=True,shell=True)
    # ping = subprocess.Popen(bashCmd, stdout=subprocess.PIPE,universal_newlines=True,shell=True)
    # p2 = subprocess.Popen(ping, stdout=subprocess.PIPE,universal_newlines=True)
    # csv_file = p2.communicate()[0]
    text = run_util.communicate()[0]
    return text

def startHSS():
    print('Start HotSpotShield VPN...')
    pyautogui.hotkey('ctrl', 'shift', 'C')
    pyautogui.sleep(3)
    pyautogui.click(x=690,y=750)
    
   
def startTestApp():
    print('Start TestApp VPN...')
    # app_name = 'HydraTestAppShell'
    backend_url = 'https://api-prod-partner-us-west-2.northghost.com/'
    carrier_id = 'kasperskylab'
    node_ip = '185.209.178.61'

    app = Application(backend="uia").start('TestApp.exe')
    dlg = Desktop(backend="uia").HydraTestAppShell

    # select backend URL
    dlg.child_window(auto_id='BackURL', control_type="ComboBox").select(backend_url)
    # select Carrier ID
    dlg.child_window(auto_id='CarrierId', control_type="ComboBox").select(carrier_id)
    # Login
    dlg.child_window(auto_id='LoginBtn', control_type="Button").click()

    # checking the Login state
    state = dlg.child_window(auto_id='LoginState', control_type="Edit").texts()
    print(state[0])
    
    dlg.child_window(auto_id='CredentialBtn', control_type="Button").click()
    for i in range(5):  # delete all 5 received server's ip
        if dlg.child_window(best_match='REMOVE'+str(i), control_type="Button").exists():
            dlg.child_window(best_match='REMOVE'+str(i), control_type="Button").click()
        # add new server IP address
    dlg.child_window(best_match='Add server', control_type="Button").click()
    dlg.child_window(auto_id='ProviderServerTextBox', control_type="Edit").set_text(node_ip)
    # applying new ip
    dlg.child_window(auto_id='ApplyCredentialBtn', control_type="Button").click()
    # CONNECT 
    dlg.child_window(auto_id='ConnectVpnBtn', control_type="Button").click()
    #check VPN State 
    time.sleep(5)
    state_vpn = dlg.child_window(auto_id='VpnConnectState', control_type="Text").texts()
    print(state_vpn[0])
    # check if IP address is changed correctly
    dlg.child_window(auto_id='IPApiBtn', control_type="Button").click()
    # checking the VL IP address
    time.sleep(5)
    my_ip = dlg.child_window(auto_id='addressLbl', control_type="Edit").texts()
    print(my_ip)
    return my_ip


def stopHSS():
    print('Stop HotSpotShield VPN...')
    pyautogui.click(x=690,y=1275)

def stopTestApp():
    print('Stop TestApp VPN...')
    
    dlg = Desktop(backend="uia").HydraTestAppShell
    # C:\Program Files (x86)\Hydra Test Application>taskkill /im "TestApp.exe" /f
    dlg.child_window(auto_id='DisconnectVpnBtn', control_type="Button").click()
    time.sleep(10)

def ip_checker():
    http = urllib3.PoolManager()
    url = 'http://ip-api.com/json'
    try:
        r = http.request('GET',url)
        json_respond = json.loads(r.data.decode('utf-8'))
        ip = json_respond['query']
        region = json_respond['regionName']
        city = json_respond['city']
        print('IP Checker:.....I\'m here:',ip,region,city)
        return (ip,region,city)
    except Exception as e:
        print(e,"IP address didn't verify")

def latency_check():
    
    # vl = ip_checker()
    latency = []
    
    path = "C:\\Users\\i.mamutkina\\Downloads\\Performance_tests\\bin\\"
    with open(path+'output_.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        line_count = 0
       
        for row in csv_reader:

            if line_count == 0:
                row_list = ["","","","","","",""]
                line_count +=1
            else:
                try:
                    # date = row[1]
                    # protocol_type = row[11]
                    # testServer = row[12]
                    if row[13] == "Latency" and row[11] == 'http' and row[12].startswith('Amazon'):
                        row_list[4] = str(row[17])
                        line_count += 1

                    elif row[13] == "Latency" and row[11] == 'http' and row[12].startswith('Google'):
                        row_list[6] = str(row[17])
                        line_count += 1

                    elif row[13] == "Latency" and row[11] == 'http' and row[12].startswith('Microsoft'):
                        row_list[2] = str(row[17])
                        line_count += 1

                    elif row[13] == "Latency" and row[11] == 'https' and row[12].startswith('Amazon'):
                        row_list[3] = str(row[17])
                        
                    elif row[13] == "Latency" and row[11] == 'https' and row[12].startswith('Google'):
                        row_list[5] = str(row[17])
                        line_count += 1

                    elif row[13] == "Latency" and row[11] == 'https' and row[12].startswith('Microsoft'):
                        row_list[1] = str(row[17])
                        line_count += 1

                    else:
                        row_list[0] = row[0]
                        line_count += 1
               
                    if row_list.count('') == 0:
                        
                        latency.append(row_list)
                        row_list = ["","","","","","",""]
                except IndexError as e: 
                    print("Oops!", e.__class__, "occurred.")
    print(latency)       
    return latency

def sending_csv_to_google(latency,worksheet):
    # file = 'performance-test-4apps-c57c7e47bceb.json"
    # gc = gspread.service_account(filename='C:\\Users\\i.mamutkina\\AppData\\Roaming\\gspread\\performance-testing-302402-ffd75e1559d3.json')
    gc = gspread.service_account(filename='C:\\Users\\i.mamutkina\\AppData\\Roaming\\gspread\\performance-testing-302402-ffd75e1559d3.json')
    sh = gc.open('CloudHarmony_test_results')
    # worksheet = sh.get_worksheet(0)
    # worksheet = sh.worksheet("HSS")
    worksheet = sh.worksheet(worksheet)
    # worksheet2 = sh.worksheet("TestApp")
    
    # values_list = worksheet.row_values(1)
    worksheet.append_rows(latency, table_range='A3:G')
    print(f"Data was delivered to {worksheet} worksheet")
    return

def cleanUp_csv_file(csv_file):
    path = "C:\\Users\\i.mamutkina\\Downloads\\Performance_tests\\bin\\"
    empty_list = []
    with open(path+csv_file, "w") as csv_f:
        writer = csv.writer(csv_f)
        csv_f = writer.writerows(empty_list)
    return csv_f


def test_app_run(d):

    if ip_checker()[0] == "96.86.145.150":
        startTestApp()
    print(f"TestApp is running # {d} ")
    ip_checker()
    
    print(f"Utilite is running # {d} ") 
    util()
    print(f"Utilite is stopping # {d} ")
    time.sleep(10)
    stopTestApp()
    time.sleep(10)
    ip_checker()
    print(f"TestApp is stopped # {d} ")
    latency = latency_check()
    print(f"Sending data to Google Spreadsheet # {d} ")
    sending_csv_to_google(latency,worksheet1)
    cleanUp_csv_file(csv_filename)



def hss_app_run(d):
    
    if ip_checker()[0] == "96.86.145.150":
        startHSS()
        
    print(f"HSS running # {d} ")
    time.sleep(10)
    ip_checker()
    print(f"Utilite is running # {d} ") 
    util()
    print(f"Utilite is stopping # {d} ")
    time.sleep(5)
    stopHSS()
    time.sleep(5)
    print(f"HSS is stopped # {d} .")
    ip_checker()
    latency = latency_check()
    print(f"Sending data to Google Spreadsheet # {d} ")
    sending_csv_to_google(latency,worksheet2)
    cleanUp_csv_file(csv_filename)
    
if __name__ == "__main__":
    print('**************** ++++ START >>>>>>>*****************')
    worksheet1,worksheet2 = "TestApp_by_server","HSS_by_server"
    csv_filename = 'output_.csv'
    cleanUp_csv_file(csv_filename)
    d = 1
    m = 20
    while d<=m:
        timestamp = time.localtime()
        time_str = time.strftime("%m/%d/%Y, %H:%M:%S",timestamp)
        start_time =  time.time()
        print(f"Test run # {d} started at {time_str}")
        if d%2 == 0:
            test_app_run(d)
        hss_app_run(d)
        
        print(f"Test run # {d} is over")
        end_time = time.time()
        run_time = end_time-start_time
        delta ="{}".format(str(timedelta(seconds=int(run_time))))
        print(f"It's been {d} out of {m} test runs completed. It took {delta} ")
        time.sleep(100)
        d+=1
    
    timestamp = time.localtime()
    time_str = time.strftime("%m/%d/%Y, %H:%M:%S",timestamp)
    end_time = time.time()
    total_time = run_time = end_time-start_time
    delta ="{}".format(str(timedelta(seconds=int(run_time))))
    print(f"It took {delta} and stopped at {time_str}" )
    print('***************<<<<<<<<<  END  >>>>>>>*****************')
