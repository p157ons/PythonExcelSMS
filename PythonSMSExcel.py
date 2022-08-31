
from ipaddress import IPv4Address
from pyairmore.request import AirmoreSession
from pyairmore.services.device import DeviceService
from pyairmore.services.messaging import MessagingService  # to send messages
from openpyxl import load_workbook
from msvcrt import getch
from time import sleep
from tkinter import filedialog

def openFile():
  filepath_function = filedialog.askopenfilename()
  return filepath_function


openFile()
set_ip = input("Set IP of Phone [Check in Airmore app] in format: 192.168.1.10: " )
ip = IPv4Address(set_ip)  # whatever server's address is
session = AirmoreSession(ip)  # port is default to 2333
service = DeviceService(session)
details = service.fetch_device_details()
details.power  # 0.65
details.brand  # gm
session.is_server_running  # True if Airmore is running
was_accepted = session.request_authorization()
print("Session Connection True")  # True if accepted
service = MessagingService(session)

loop_variable = 1
while loop_variable == 1:

    # path to file
    excel_file_open = input("Select Excel file - Press T")
    if ord(excel_file_open) == 84 or ord(excel_file_open) == 116 :
       filepath = openFile()
    else:

        break
    
    #filepath = "C:\\Users\\L&I Legal\\Desktop\\" + excel_file_name + ".xlsx"
    # column to read
    column = "I"  # suppose it is under "A"
    columnb = "G"
    # number of cols to get
    start_row = int(input("Please set start row for excel file: "))
    length = start_row
    end_row = int(input("Please set end row for excel file: "))
    length_v1 = end_row
    workbook = load_workbook(filepath, read_only=True)
    worksheet = workbook.active  # we will get the active worksheet


    number_of_send_sms = 0
    for length in range (length_v1):
        cell = "{}{}".format(column, length + 1)
        number = worksheet[cell].value
        cell = "{}{}".format(columnb, length + 1)
        message = worksheet[cell].value


        service.send_message(number, message)
        number_of_send_sms += 1



        print("Sms sent to " + str(number_of_send_sms) + " persons")
        sleep(9)

    print("Continue? Press: T ")
    char = getch()
    print(char)

    if ord(char) == 84 or ord(char) == 116 :
        loop_variable = 1
    else:

        break


