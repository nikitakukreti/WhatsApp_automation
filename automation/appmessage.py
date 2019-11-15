from time import sleep
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
import socket
from xlrd import *

data=[]
row=[]
l=[]
book=input("Enter the Book name without .xlsx: ")
booknm=book+".xlsx"
file=open_workbook(booknm)
sheet=file.sheet_by_index(0)
column=input("Enter the column name whose data is to be retieved: ")
col=column.lower()
message_text=input("Enter the text message you wish you send: ") # message you want to send
frow=sheet.cell_value(0, 0)
for x in range(sheet.ncols):
    row.append(sheet.cell_value(0,x))
for y in row:
    if(y.lower()==col):
            for i in range(sheet.nrows):
                    data.append(sheet.cell_value(i,row.index(y)))
            data.remove('PhoneNo')
            a=1
            for j in data:
                if(len(str(j))>9):
                    s1=str(j)
                    ind=s1[0:2]
                    if(ind=="91"):
                        l.append(int(sheet.cell_value(a,row.index(y))))
                    else:
                        s2=sheet.cell_value(a,row.index(y))
                        s3="91"+str(int(s2))
                        l.append(int(s3))
                a=a+1

no_of_message=1 # no. of time you want the message to be send
moblie_no_list=l # list of phone number can be of any length
def element_presence(by,xpath,time):
    element_present = EC.presence_of_element_located((By.XPATH, xpath))
    WebDriverWait(driver, time).until(element_present)

def is_connected():
    try:
        # connect to the host -- tells us if the host is actually
        # reachable
        socket.create_connection(("www.google.com", 80))
        return True
    except :
        is_connected()
driver = webdriver.Chrome(executable_path="C:\Program Files (x86)/Google/Chrome/Application/chromedriver.exe")
driver.get("http://web.whatsapp.com")
sleep(10) #wait time to scan the code in second

def send_whatsapp_msg(phone_no,text):
    driver.get("https://web.whatsapp.com/send?phone={}&source=&data=#".format(phone_no))
    try:
        driver.switch_to_alert.accept()
    except Exception as e:
        print("Alert errror")

    try:
        element_presence(By.XPATH,'//*[@id="main"]/footer/div[1]/div[2]/div/div[2]',30)
        txt_box=driver.find_element(By.XPATH , '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        global no_of_message
        for x in range(no_of_message):
            txt_box.send_keys(text)
            txt_box.send_keys("\n")

    except Exception as e:
        print("invailid phone no :"+str(phone_no))
for moblie_no in moblie_no_list:
    try:
        send_whatsapp_msg(moblie_no,message_text)

    except Exception as e:
        sleep(10)
        is_connected()
