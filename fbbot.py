# Note: For proper working of this Script Good and Uninterepted Internet Connection is Required
# Keep all contacts unique
# Can save contact with their phone Number

# Import required packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime
import time
import openpyxl as excel
import random
import pickle

# function to read contacts from a text file
def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['B']
    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        contact = "\"" + contact + "\""
        lst.append(contact)
    return lst

def readNames(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        contact = "\"" + contact + "\""
        lst.append(contact)
    return lst

# Target Contacts, keep them in double colons
# Not tested on Broadcast
# targets = readContacts("data4.xlsx")
# names = readNames("data4.xlsx")
# can comment out below line
print('starting targets')
# print(targets)
# print('end of targets')
# Driver to open a browser
driver = webdriver.Chrome()
url = driver.command_executor._url
session_id = driver.session_id
print(url)
print(session_id)
print(driver)
# url = 'http://127.0.0.1:57931'
# session_id = '088f24b92dba47f32a2ce05d00e908cd'
# driver = webdriver.Remote(command_executor=url,desired_capabilities={})
# time.sleep(2)
# driver.session_id = session_id

#link to open a site
driver.get("https://web.telegram.org/")

# 10 sec wait time to load, if good internet connection is not good then increase the time
# units in seconds
# note this time is being used below also
wait = WebDriverWait(driver, 10)
wait5 = WebDriverWait(driver, 5)
input("Scan the QR code and then press Enter")
# pickle.dump(driver.get_cookies() , open("QuoraCookies.pkl","wb"))
for cookie in pickle.load(open("QuoraCookies.pkl", "rb")):
    driver.add_cookie(cookie)


# Message to send list
# 1st Parameter: Hours in 0-23
# 2nd Parameter: Minutes
# 3rd Parameter: Seconds (Keep it Zero)
# 4th Parameter: Message to send at a particular time
# Put '\n' at the end of the message, it is identified as Enter Key
# Else uncomment Keys.Enter in the last step if you dont want to use '\n'
# Keep a nice gap between successive messages
# Use Keys.SHIFT + Keys.ENTER to give a new line effect in your Message

msgToSend = [
                [12, 32, 0, ",  We are happy to announce our next meetup in Hyderabad with Binance on 1st June, 2019 at 91springboard, Hi-tech City from 4:00 - 8:00 pm. RSVP now to book your spot now https://forms.gle/CbY8pACpjbtqo28A7  \n https://www.facebook.com/events/678126052624684/" + Keys.SHIFT + Keys.ENTER ]
            ]

msg2 = ", In between all the excitement around India Dapp Fest, we hope you didn't miss the amazing next meetup in Hyderabad where we host the Binance team as they come to the city. In case you didn't know, it's happening this Saturday (1st July) at 91 springboard, Hitech City. \n RSVP now before all slots fill in https://forms.gle/CbY8pACpjbtqo28A7 \n https://www.facebook.com/events/678126052624684/" + Keys.SHIFT + Keys.ENTER
msg3 = ", We are happy to announce our next Hyderabad meetup with Binance on 1st June, 2019 at 91springboard, Hi-tech City from 4:00 - 8:00 pm. Book your spot now at https://forms.gle/CbY8pACpjbtqo28A7 RSVP form \n Find more details about the about the even here: https://www.facebook.com/events/678126052624684/" + Keys.SHIFT + Keys.ENTER
# Count variable to identify the number of messages to be sent

searchBox = driver.find_element_by_class_name('im_dialogs_search_field')
searchBox.send_keys('Core, IndiaDappFest'+ Keys.ENTER)
time.sleep(3)
groupHeader = driver.find_element_by_class_name('tg_head_peer_title')
groupHeader.click()
time.sleep(2)
contactList = driver.find_elements_by_class_name('md_modal_list_peer_name')
i = 0
while i < len(contactList):
    try:    
        contactList[i].click
        i = i+1
        time.sleep(1)
        messageButton = driver.find_elements_by_class_name('md_modal_split_action')
        messageButton[1].click
        time.sleep(1)
        textArea = driver.find_element_by_class_name('composer_rich_textarea')
        textArea.send_keys('hellooo message from bot' + Keys.ENTER)
        searchBox = driver.find_element_by_class_name('im_dialogs_search_field')
        searchBox.send_keys('Core, IndiaDappFest'+ Keys.ENTER)
        time.sleep(2)
        groupHeader = driver.find_element_by_class_name('tg_head_peer_title')
        groupHeader.click()
    except:
        print('except')

# count = 0
# while count<len(msgToSend):

#     # Identify time
#     curTime = datetime.datetime.now()
#     curHour = curTime.time().hour
#     curMin = curTime.time().minute
#     curSec = curTime.time().second

#     # if time matches then move further
#     # if msgToSend[count][0]==curHour and msgToSend[count][1]==curMin and msgToSend[count][2]==curSec:
#     if count == 0:
#         # utility variables to tract count of success and fails
#         success = 0
#         sNo = 1
#         failList = []

#         # Iterate over selected contacts
#         index = 0
#         for target in targets:
#             print(sNo, ". Target is: " + target)
#             print("namee is "+ names[index])
#             sNo+=1
#             try:
#                 # Select the target
#                 x_arg = '//span[contains(@title,' + target + ')]'
#                 try:
#                     wait.until(EC.presence_of_element_located((
#                         By.XPATH, x_arg
#                     )))
#                 except:
#                     # If contact not found, then search for it
#                     print('91')
#                     # searBoxPath = '//*[@id="input-chatlist-search"]'
#                     # wait5.until(EC.presence_of_element_located((
#                         # By.ID, "input-chatlist-search"
#                     # )))
#                     inputSearchBox = driver.find_element_by_css_selector("[title='Search or start new chat']")
#                     print('inputSearchBox')
#                     time.sleep(random.choice([0.5,1,1.5,2]))
#                     inputSearchBox.click()
#                     # click the search button
#                     # driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/div[2]/div/button').click()
#                     time.sleep(random.choice([0.5,1,1.5,2]))
#                     inputSearchBox.clear()
#                     inputSearchBox.send_keys(target[1:len(target) - 1])
#                     print('Target Searched')
#                     # Increase the time if searching a contact is taking a long time
#                     time.sleep(random.choice([2.5,1.5,2]))

#                 # Select the target
#                 driver.find_element_by_xpath(x_arg).click()
#                 print("Target Successfully Selected")
#                 time.sleep(random.choice([1,1.5,2]))

#                 # Select the Input Box
#                 inp_xpath = "//div[@contenteditable='true']"
#                 input_box = wait.until(EC.presence_of_element_located((
#                     By.XPATH, inp_xpath)))
#                 time.sleep(random.choice([1,1.5,2]))

#                 # Send message
#                 # taeget is your target Name and msgToSend is you message
#                 input_box.send_keys('Hey '+ names[index].strip('\"') + random.choice([msgToSend[count][3],msg2,msg3]) + Keys.SPACE) # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
#                 # Link Preview Time, Reduce this time, if internet connection is Good
#                 time.sleep(random.choice([5.5,4,4.5,5]))
#                 input_box.send_keys(Keys.ENTER)
#                 print("Successfully Send Message to : "+ target + '\n')
#                 success+=1
#                 index+=1
#                 time.sleep(1)

#             except:
#                 # If target Not found Add it to the failed List
#                 print("Cannot find Target: " + target)
#                 failList.append(target)
#                 index+=1
#                 pass

#         print("\nSuccessfully Sent to: ", success)
#         print("Failed to Sent to: ", len(failList))
#         print(failList)
#         print('\n\n')
#         index+=1
#         count+=1
