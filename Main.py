# *-- Python import --*
import xlrd
import time

# *-- Third Party imports --*
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.command import Command
from webdriver_manager.chrome import ChromeDriverManager



# f = "d:\\test.xlsx"
# wb = xlrd.open_workbook(f)
# sheet = wb.sheet_by_index(0)
# sheet.cell_value(0,0)
# name = []
# number = []
final_msg = []
# for i in range(sheet.nrows):
#     print(str(sheet.cell_value(i,0)))
#     name.append(sheet.cell_value(i,0))

# for j in range(sheet.nrows):
#     print(int(sheet.cell_value(j, 1)))
#     number.append(str(int(sheet.cell_value(j, 1))))

name = ["gopal"]
number = ["8679764495"]
print(name,number)
msg_count = 0
# msg = input("Enter Your Personalize Message by Giving <user>")
msg = '''hi <user> Nokia online  service on 
Realme online service on'''
for contact_name in name:
    message = msg.replace("<user>",str(contact_name))
    final_msg.append(message)

# *--   Chrome options --*
prefs = {"profile.managed_default_content_settings.images": 2}
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", prefs)
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("log-level=3")


def get_driver_status(driver):
    try:
        driver.execute(Command.STATUS)
        return "Alive"
    except Exception:
        if(driver):
            driver.quit()
        return "Dead"


def intialize_driver():
    driver = None
    attempt = 0
    while driver is None or get_driver_status(driver) == "Dead":
        driver = webdriver.Chrome(ChromeDriverManager().install())
        attempt += 1
        if attempt > 3:
            raise AssertionError("Unable to Initialize driver")
    return driver

driver = intialize_driver()
print(driver)



print(final_msg)
print(final_msg[0])
# print(final_msg[1])
# driver=webdriver.Firefox(executable_path="C:\\webdrivers\\geckodriver.exe")
driver.get('https://messages.google.com/web/authentication?redirected=true')
time.sleep(5)
remember_pc = driver.find_element_by_class_name('mat-slide-toggle-thumb')
remember_pc.click()

input("Press Enter After Scanning QR Code ")
name_index = 0

for num in number:
    # Simulate Keys
    # action = ActionChains(driver)

    continuousMsg = final_msg[name_index].split("\n")
    print(continuousMsg)
    #if len(num)==10:
    time.sleep(3)
    start_chat = driver.find_element_by_class_name('fab-label')
    start_chat.click()
    time.sleep(5)
    phn_num_box = driver.find_element_by_class_name('input')
    phn_num_box.click()
    phn_num_box.send_keys(num)
    time.sleep(1)
    phn_num_box.send_keys(Keys.ENTER)
    time.sleep(4)
    for i in range(len(continuousMsg)) :
        msg_box = driver.find_element_by_xpath('//textarea [@placeholder = "Text message"]')
        msg_box.click()
        msg_box.send_keys(continuousMsg[i])
        msg_box.send_keys(Keys.SHIFT, Keys.ENTER)
        # action.key_down(Keys.SHIFT)
        # action.send_keys(Keys.ENTER)
        # action.key_up(Keys.SHIFT)
        # action.key_up(Keys.Enter)
        # action.perform()
        time.sleep(2)
    msg_box.send_keys(Keys.ENTER)
    time.sleep(2)
    msg_count+=1
    print("Message Successfully Sent To",num)
    name_index = name_index+1

    # else:
    #     print("Please Check This Number",num,"\nThis Is Invalid Number!!\nNumber Must Have Only 10 Digits!!!")

print("Total",msg_count,"Messages Are Sent!!")
