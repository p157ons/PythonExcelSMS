
from ipaddress import IPv4Address
from pyairmore.request import AirmoreSession
from pyairmore.services.device import DeviceService
from pyairmore.services.messaging import MessagingService  # to send messages
from openpyxl import load_workbook
#from getch import getch,pause
from msvcrt import getch
from time import sleep


podaj_ip = input("Podaj IP w Formacie: 192.168.1.10: " )
ip = IPv4Address(podaj_ip)  # whatever server's address is
session = AirmoreSession(ip)  # port is default to 2333
service = DeviceService(session)
details = service.fetch_device_details()
details.power  # 0.65
details.brand  # gm
session.is_server_running  # True if Airmore is running
was_accepted = session.request_authorization()
print("Status łączności True")  # True if accepted
service = MessagingService(session)

#service.send_message("+48790342432211", "HELLO WORLD")

powtorz = 1
while powtorz == 1:

    # path to file
    nazwapliku = input("Podaj nazwę pliku który znajduje się na pulpicie:")
    filepath = "C:\\Users\\L&I Legal\\Desktop\\" + nazwapliku + ".xlsx"
    # column to read
    column = "I"  # suppose it is under "A"
    columnb = "G"
    # number of cols to get
    dlugoscplikow = int(input("Podaj numer wiersza OD którego chcesz wysyłać sms: "))
    length = dlugoscplikow
    dlugoscplikow_v1 = int(input("Podaj numer wiersza DO którego chcesz wysłać sms: "))
    length_v1 = dlugoscplikow_v1
    workbook = load_workbook(filepath, read_only=True)
    worksheet = workbook.active  # we will get the active worksheet

    #phone_numbers = []
    #for i in range(length):
       # cell = "{}{}".format(column, i+1)
      #  number = worksheet[cell].value
      #  if number != "" or number is not None:
      #      phone_numbers.append(str(number))

    #MessageB = []
    #for i in range(length):
      #  cell = "{}{}".format(columnb, i+1)
      #  message = worksheet[cell].value
      #  if message != "" or message is not None:
      #      MessageB.append(str(message))

    liczbawyslanych = 0
    for length in range (length_v1):
        cell = "{}{}".format(column, length + 1)
        number = worksheet[cell].value
      #  if number != "" or number is not None:
         #   number.append(str(number))
        cell = "{}{}".format(columnb, length + 1)
        message = worksheet[cell].value
        #if message != "" or message is not None:
           # message.append(str(message))

       # if  number == "" or number is  None:
          #  break


        service.send_message(number, message)
        liczbawyslanych += 1



        print("Sms Wysłane do " + str(liczbawyslanych) + " osób")
        sleep(9)

    print("Czy chcesz kontynuować? Kliknij T ")
    char = getch()
    print(char)

    if ord(char) == 84 or ord(char) == 116 :
        powtorz = 1
    else:

        break


#message = "50% discount at Lorem Ipsum Co. tomorrow."
#for number in phone_numbers :
           # service.send_message(number, message)
