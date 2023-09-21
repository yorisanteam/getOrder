from telnetlib import SE
from tkinter import SW
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import os,sys,ssl,imaplib,csv,email,datetime
sys.path.append("lib.bs4")
from email.header import decode_header,make_header

todayUtc   = datetime.datetime.utcnow().date()
todayYear  = str(todayUtc.strftime("%Y%m%d")[0:4])
todayMonth = str(todayUtc.strftime("%Y%m%d")[4:6])
todayDay   = str(todayUtc.strftime("%Y%m%d")[6:8])

server = "imap.mail.yahoo.co.jp"
usrNum = "usrNum"
password = "password"
sWord = "支払い完了"
auctionURL = "https://contact.auctions.yahoo.co.jp/seller/top?aid="

orderList = [] #List to store auction IDs and URLs of sold items
lastAucIdList = []

inputDateMonthStr = ''
targetF = '(' #Characters to slice when extracting only the auction ID
targetS = ')'

inputDateYear = input('取得する年を入力してください') #Retrieve information for a given date
inputDateMonth = input('取得する月を入力してください')
inputDateDay = input('取得する日を入力してください')
if inputDateMonth == '1': #Change to format for imap search
    inputDateMonthStr = 'Jan'
elif inputDateMonth == '2':
    inputDateMonthStr = 'Feb'
elif inputDateMonth == '3':
    inputDateMonthStr = 'Mar'
elif inputDateMonth == '4':
    inputDateMonthStr = 'Apr'
elif inputDateMonth == '5':
    inputDateMonthStr = 'May'
elif inputDateMonth == '6':
    inputDateMonthStr = 'Jun'
elif inputDateMonth == '7':
    inputDateMonthStr = 'Jul'
elif inputDateMonth == '8':
    inputDateMonthStr = 'Aug'
elif inputDateMonth == '9':
    inputDateMonthStr = 'Sep'
elif inputDateMonth == '10':
    inputDateMonthStr = 'Oct'
elif inputDateMonth == '11':
    inputDateMonthStr = 'Nov'
elif inputDateMonth == '12':
    inputDateMonthStr = 'Dec'

orderDay = inputDateYear + '/' + inputDateMonth + '/' + inputDateDay

try:
    print(usrNum,"を取得します")
    context = ssl.create_default_context()
    imapclient = imaplib.IMAP4_SSL(server,993,ssl_context=context)
    imapclient.login(usrNum,password)
    imapclient.select()
    typ,data = imapclient.search(None,"ON",f'{inputDateDay}-{inputDateMonthStr}-{inputDateYear}') #Check the mail
    for num in data[0].split():
        typ,data = imapclient.fetch(num,"(RFC822)")
        email_message = email.message_from_bytes(data[0][1])
        subject = str(make_header(decode_header(email_message['Subject'])))
        # Date = str(make_header(decode_header(email_message['Date'])))
        # msg_encoding = 'utf-8_sig'

        # if email_message.is_multipart() == False: # Single
        #     byt  = bytearray(email_message.get_payload(), msg_encoding)
        # else:   # Multi
        #     prt  = email_message.get_payload()[0]
        #     byt  = prt.get_payload(decode=True)

        if sWord in subject: #Extract auctionID
            sliceFir = subject[-13:]
            sliceSec = sliceFir.find(targetF)
            sliceThi = sliceFir[sliceSec+1:]
            sliceFor = sliceThi.find(targetS)
            sliceFif = sliceThi[:sliceFor]
            orderList.append(auctionURL+sliceFif)
            lastAucIdList.append(sliceFif)
            print(sliceFif)
            print('----------------------')

    print("注文数：",len(orderList))
    sleep(5)
    imapclient.close()
    imapclient.logout()

except Exception as ee:
    sys.stderr.write("*** error ***\n")
    sys.stderr.write(str(ee) + '\n')
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)
    lastKey = input("エラーが発生しました。ボタンを押してください")

nextButtonXpath = " " #Describe the required XPath
loginButtonXpath = " "
customNameXpath = " "
customAddressNumXpath = " "
customAddressXpath = " "
customSendXpath = " "
sellPriceXpath = " "
loadPriceXpath = " "
customMessageXpath = " "

infoList = [" "," "," "," "," "," "," "," "]
allInfoLists = []
allInfoHead = ["auctionId","itemName","postNum","address","orderPrice","loadPrice","note","orderDay"]

try:
    for urlNum in orderList:
        sleep(2)
        driver = webdriver.Chrome("./chromedriver.exe")
        driver.implicitly_wait(10)
        sleep(2)
        print(usrNum,"開始")
        driver.get(urlNum)
        sleep(2)
        try:
            element = driver.find_element(By.ID,"username")
        except:
            element = driver.find_element(By.ID,"login_handle")
        element.send_keys(usrNum)
        sleep(2)
        try:
            element = driver.find_element(By.XPATH,"//*[@id='btnNext']").click()
        except:
            element = driver.find_element(By.XPATH,nextButtonXpath).click()
        sleep(2)
        try:
            element = driver.find_element(By.ID,"passwd")
        except:
            element = driver.find_element(By.ID,"password")
        element.send_keys(password)
        sleep(2)
        try:
            element = driver.find_element(By.XPATH,"//*[@id='btnSubmit']").click()
        except:
            element = driver.find_element(By.XPATH,loginButtonXpath).click()
        sleep(2)

        try:
            infoList[1] = driver.find_element(By.XPATH,customNameXpath).text
        except Exception as ee:
            infoList[1] = " "
        try:
            infoList[2] = driver.find_element(By.XPATH,customAddressNumXpath).text
        except Exception as ee:
            infoList[2] = " "
        try:
            infoList[3] = driver.find_element(By.XPATH,customAddressXpath).text
        except Exception as ee:
            infoList[3] = " "
        try:
            infoList[4] = driver.find_element(By.XPATH,sellPriceXpath).text
        except Exception as ee:
            infoList[4] = " "
        try:
            infoList[5] = driver.find_element(By.XPATH,loadPriceXpath).text
        except Exception as ee:
            infoList[5] = " "
        try:
            infoList[6] = driver.find_element(By.XPATH,customMessageXpath).text
        except Exception as ee:
            infoList[6] = " "
        infoList[7] = orderDay
        allInfoLists.append(infoList.copy()) #Make a copy and paste it
        print(urlNum[-11:],"完了")
        sleep(2)
    driver.quit()

    for i in range(len(lastAucIdList)):
        allInfoLists[i][0] = lastAucIdList[i]
    allInfoLists.insert(0,allInfoHead)
    finKey = input("完了。ボタンを押したら終了します")

except Exception as ee:
    sys.stderr.write("*** error ***\n")
    sys.stderr.write(str(ee) + '\n')
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)
    lastKey = input("エラーが発生しました。ボタンを押してください")

with open("./output/output.csv","w",newline='',encoding="utf_8_sig") as f:
    writer = csv.writer(f)
    writer.writerows(allInfoLists) #Make a csv file
