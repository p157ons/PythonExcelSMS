
import csv
from ipaddress import IPv4Address
from pyairmore.request import AirmoreSession
from pyairmore.services.device import DeviceService
from pyairmore.services.messaging import MessagingService  # to send messages
from openpyxl import load_workbook
from msvcrt import getch
from time import sleep
from tkinter import filedialog

def main():
	print("\033[H\033[J")
	mainmenu()

def mainmenu():

  choicest1 = input("""\033[1;32m     
   (`|\/|(` (`[~|\ ||\[~|)
   _)|  |_) _)[_| \||/[_|\           
\033[1;0m                                                  
    === main menu ====
    \033[1;33m
    1. Single SMS:   
    2. Bulk SMS:     
    3. API Settings: 
    4. Exit: \033[1;0m        
    Please enter your choice [1-4]: """)
  if choicest1 == "1":
    print("\033[H\033[J")
    bulk_sending()
  elif choicest1 == "2":
     print("\033[H\033[J")
     bulk_sending()
  elif choicest1 == "3":
    print("\033[H\033[J")
    print("No Api Setting Avaivle for the moment")
    mainmenu() 
    #apisettings()
  elif choicest1 == "4":
     print("		Bye Bye! :)")
     pass     
  else: 
    print("\033[H\033[J")
    print("You must only select either [1-4]")
    print("Please try again")
    mainmenu()  

def openFile():
  filepath_function = filedialog.askopenfilename()
  if filepath_function.endswith('.xlsx') or filepath_function.endswith('.csv'):
    return filepath_function
  else:
    print("Wrong File format - select again")
    filepath_function = filedialog.askopenfilename()
    return filepath_function

def connection():
  print("We need to create connection between Phone and Computer please open Airmore application and follow instruction")
  set_ip = input("Set IP of Phone [Check in Airmore app] in format: '192.168.1.10' or '192.168.1.100'  : [to move back enter '0'] " )
  
  if set_ip == "0":
     print("\033[H\033[J")
     mainmenu()
     return None

  elif set_ip != "0": 

    def validate(ip_validation):
      try:
        a,b,c,d = ip_validation.split(".")
        for v in [a,b]:
          assert len(v) == 3
        assert len(c) == 1
        if  len(d) == 2 or len(d) == 3:
          True
        else:
          assert len(d) == 3

        a, b, c, d = int(a), int(b) ,int(c), int(d) 
      except Exception:
        print("\033[H\033[J")
        print("Wrong Format - Please use ip adress example: '192.168.1.10' or '192.168.1.100' " )
        print("")
        return connection()
      else:
        return True 
    if validate(set_ip) == True:
      try:
        print("\033[H\033[J")
        print("Trying to establish connection with IP " + set_ip + " ...")
        ip = IPv4Address(set_ip)  # whatever server's address is
        session = AirmoreSession(ip)  # port is default to 2333
        service = DeviceService(session)
        details = service.fetch_device_details()
        details.power  # 0.65
        details.brand  # gm
        session.is_server_running  # True if Airmore is running
        was_accepted = session.request_authorization()
        print("Session Connection Established")  # True if accepted
        service = MessagingService(session)
        return service
      except:
        print("\033[H\033[J")
        print("Could not establish connection with IP " + set_ip +"- Try again")
        print("")
        return connection()  
    else:
      return None

def bulk_sending():

  service = connection()

  if service == None:
    return None
  else:  
    choicest2()
    def choicest2():
     input("""\033[1;32m     
                                    (`|\/|(` (`[~|\ ||\[~|)
                                    _)|  |_) _)[_| \||/[_|\           
                                  \033[1;0m                                                  
                                      === main menu ====
                                      \033[1;33m
                                      1. Send SMS using existing file .xlsx or .csv:   
                                      2. Send SMS puting manually number and subject:     
                                      3. Exit      
                                      Please enter your choice [1-3]: """)

     if choicest2 == "1":
        print("\033[H\033[J")
        send_sms_using_file()
     elif choicest2 == "2":
        print("\033[H\033[J")
        print("Not Available Yet")
     elif choicest2 == "3":
        print("		Bye Bye! :)")
        pass       
     else: 
        print("\033[H\033[J")
        print("You must only select either [1-3]")
        print("Please try again")
        choicest2()                                    
      

        #loop_variable = 1
        #while loop_variable == 1:

        def send_sms_using_file():
          excel_file_open = input("To select excel / csv file with numbers, subject - Press '1' / Press '0' to break connection")
          if excel_file_open == "1": #ord(excel_file_open) == 84 or ord(excel_file_open) == 116 :
            filepath = openFile()
          elif excel_file_open == "0":
            return None
          else:
              print("\033[H\033[J")
              print("You must press '1' to select file / Press '0' to break connection- pressing other button will cause back to menu")
              print("Please try again")
              excel_file_open = input("Select Excel file - Press '1'")
              if excel_file_open == "1": #ord(excel_file_open) == 84 or ord(excel_file_open) == 116 :
                filepath = openFile()
              elif excel_file_open == "0":
                return None
              else:
                mainmenu()
          
          # column to read
          print("File uploaded")
          
          number_column = input("Give letter of column which include Phone Numbers") 
          message_column = input("Give letter of column which include Subject / Messages")
          # number of cols to get
          start_row = int(input("Please set start row for excel file: "))
          end_row = int(input("Please set end row for excel file: "))
          workbook = load_workbook(filepath, read_only=True)
          worksheet = workbook.active  # we will get the active worksheet

          number_of_send_sms = 0

          for start_row in range (end_row):
              cell = "{}{}".format(number_column, start_row + 1)
              number = worksheet[cell].value
              cell = "{}{}".format(message_column, start_row + 1)
              message = worksheet[cell].value
              service.send_message(number, message)
              number_of_send_sms += 1

              print("Sms sent to " + str(number_of_send_sms) + " persons")
              sleep(4)

          print("Loop of bulk sms ended - To end Appliction Press: T ")
          char = getch()
          print(char)

          #if ord(char) == 84 or ord(char) == 116:
           #break


main()