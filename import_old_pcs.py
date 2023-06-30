from selenium import webdriver
from selenium.webdriver import EdgeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time

options = EdgeOptions()
options.use_chromium = True

#importing the datafram
df = pd.read_excel(r"Enter path to dataframe")
#formatting date
df['book_date2'] = df['book_date'].dt.strftime('%m/%d/%Y')
print(df)


#Executable file path
s = Service(executable_path ="C:\\Users\\jberry\\Downloads\\msedgedriver.exe")
driver = webdriver.Edge(service=s, options=options)

#Pulls first instance of the webpage. This will bring up the login page.
driver.get("https://pellacorp.lightning.force.com/lightning/o/Lead/new?count=1&nooverride=1&useRecordTypeCheck=1&navigationLocation=LIST_VIEW&uid=168381270854618407&backgroundContext=%2Flightning%2Fo%2FLead%2Flist%3FfilterName%3DRecent&recordTypeId=0123i000000MdKTAA0")
#Sleep to give the webpage a chance to load
time.sleep(1)

#Enters credentials and, clicks sign in, and waits for the verification email.
un = driver.find_element("xpath",'//*[@id="input28"]')
un.send_keys('*****')

pw = driver.find_element("xpath", '//*[@id="input36"]')
pw.send_keys('*****')

signin = driver.find_element("xpath",'//*[@id="form20"]/div[2]/input')
signin.click()

# Time to wait for for the verification email. 
# ALTERNATIVE
# Use a while loop: while element on form does not exist, wait. When element is present, then iterate over rows.
time.sleep(45)



# Iteration variable for print statements.
count = 1

#Loop to 
for index, row in df.iterrows():

    # Pulls another instance of the webpage form. I use the same link for consistency in the xpaths.
    driver.get(
        "https://pellacorp.lightning.force.com/lightning/o/Lead/new?count=1&nooverride=1&useRecordTypeCheck=1&navigationLocation=LIST_VIEW&uid=168381270854618407&backgroundContext=%2Flightning%2Fo%2FLead%2Flist%3FfilterName%3DRecent&recordTypeId=0123i000000MdKTAA0")
    time.sleep(6)
    
    # Print statement to show which customer is currenty being entered in the form.
    print(str(count) + ". currently entering in " + str(row['first_name']) + " " + str(row['last_name']) + "...")
    count = count + 1

    # Entering row information
    firstname = driver.find_element("xpath",'//*[@id="input-151"]' )
    firstname.send_keys(str(row['first_name']))

    lastname = driver.find_element("xpath",'//*[@id="input-153"]')
    lastname.send_keys(str(row['last_name']))

    actions = ActionChains(driver)
    tab = actions.send_keys(Keys.TAB)
    tab.perform()

    pressq = actions.send_keys('q')
    pressq.perform()

    downarrow = actions.send_keys(Keys.ARROW_DOWN)
    downarrow.perform()

    hitenter = actions.send_keys(Keys.ENTER)
    hitenter.perform()

    address = driver.find_element("xpath", '//*[@id="input-172"]')
    address.send_keys(str(row['address']))

    city = driver.find_element("xpath",'//*[@id="input-173"]')
    city.send_keys(str(row['city']))

    state = driver.find_element("xpath",'//*[@id="input-174"]')
    state.send_keys(str(row['state']))

    zipcode = driver.find_element("xpath",'//*[@id="input-175"]')
    zipcode.send_keys(str(row['zipcode']))

    country = driver.find_element("xpath",'//*[@id="input-176"]')
    country.send_keys('US')

    tab = actions.send_keys(Keys.TAB)
    tab.perform()

    typeolder = actions.send_keys('older')
    typeolder.perform()
    time.sleep(3)
    downarrow = actions.send_keys(Keys.ARROW_DOWN)
    downarrow.perform()
    downarrow = actions.send_keys(Keys.ARROW_DOWN)
    downarrow.perform()
    hitenter = actions.send_keys(Keys.ENTER)
    hitenter.perform()

    email = driver.find_element("xpath",'//*[@id="input-164"]')
    email.send_keys(str(row['email']))

    number = driver.find_element("xpath",'//*[@id="input-232"]')
    number.send_keys(str(row['number']))

    book_date = driver.find_element("xpath", '//*[@id="input-207"]')
    book_date.send_keys(str(row['book_date2']))

    comments = driver.find_element("xpath", '//*[@id="input-277"]')
    comments.send_keys('JB' + '\n'
                       'Sales rep: ' + str(row['sales_rep']) + '\n'
                       'Contracted date: ' + str(row['book_date2']) + '\n'
                       'Product: ' + str(row['product_family']) + '\n'
                       'Price: $' + str(row['price']) + '\n'
                       'Call notes: ' + str(row['call_notes']) + '\n')
    
    # time.sleep to give me time to look over the information to make sure everything was entered in correctly. Then manually save the the information entered in the form.
    time.sleep(30)


