# import Libraries
from selenium import webdriver
from termcolor import colored
import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import csv
import time
import os

########################## Define XPATHS ######################################
# Country BOX Xpath
countrybox = '//*[@id="country-box"]/div/select'
# Country Selector Xpath
countryusa = '//*[@id="country-box"]/div/select/option[1]'
# City BOX Xpath
citybox = '//*[@id="city-state-box"]/div/select'
# City List Xpaths
city = ['//*[@id="city-state-box"]/div/select/optgroup[1]/option[1]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[2]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[3]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[4]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[5]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[6]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[7]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[8]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[9]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[10]',
        '//*[@id="city-state-box"]/div/select/optgroup[1]/option[11]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[1]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[2]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[3]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[4]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[5]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[6]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[7]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[8]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[9]',
        '//*[@id="city-state-box"]/div/select/optgroup[2]/option[10]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[1]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[2]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[3]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[4]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[5]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[6]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[7]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[8]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[9]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[10]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[11]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[12]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[13]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[14]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[15]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[16]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[17]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[18]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[19]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[20]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[21]',
        '//*[@id="city-state-box"]/div/select/optgroup[3]/option[22]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[1]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[2]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[3]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[4]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[5]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[6]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[7]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[8]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[9]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[10]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[11]',
        '//*[@id="city-state-box"]/div/select/optgroup[4]/option[12]']
# Heating Type XPath
heatingtype = ''
heatingelec = '//*[@id="heating-type-box"]/div/label[1]/input'
heatingfuel = '//*[@id="heating-type-box"]/div/label[2]/input'

#  Bulding Type XPath
smalloffice = '//*[@id="tabs-1"]/label[10]/input'
mediumoffice = '//*[@id="tabs-1"]/label[11]/input'
largeoffice = '//*[@id="tabs-1"]/label[12]/input'

# From Date Picker BOX XPath
datefrombox = '//*[@id="from_date"]'
# Month Picker Selector XPath
monthselector = '//*[@id="ui-datepicker-div"]/div/div/select'
# From Month XPath
month = ['//*[@id="ui-datepicker-div"]/div/div/select/option[1]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[2]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[3]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[4]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[5]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[6]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[7]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[8]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[9]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[10]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[11]',
         '//*[@id="ui-datepicker-div"]/div/div/select/option[12]',
         ]

# To Date Picker BOX XPath
datetobox = '//*[@id="to_date"]'
# To Month Picker Selector XPath
#     alredy defined
#  Month XPath
#     alredy defined
#  Day XPath
#     alredy defined

# Click Button to add load shapes XPath
addshapes = '//*[@id="add-shape"]'
# Remove All Button XPath
removeall ='//*[@id="graph-tabs-0"]/div[8]/span[9]/button'

######################################################################################
#########################     Main Script     ########################################
######################################################################################

# Create Google Chrome Options Handler
chromeOptions = webdriver.ChromeOptions()

# Create Safe Browsing
#chromeOptions.add_argument( '— incognito')

# Create Options of A HeadLess Browser
#chromeOptions.add_argument('--headless')
#chromeOptions.add_argument('--disable-gpu')

# Change ChromeDriver Path
chromedriver = 'C:\\Users\\Link To Life\\Python36_64\\chromedriver.exe'

# Change Download Directory
prefs = {"download.default_directory": "C:\CSV"}
chromeOptions.add_experimental_option("prefs", prefs)
chromeOptions.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
download_path = 'C:\CSV'

# Create a HeadLess Browser
driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)

# Define Implicite Waits To Respond
driver.implicitly_wait(1000)

# Driver Initialized LOG
now = datetime.datetime.now()
print(now.strftime("%Y-%m-%d %H:%M:%S"),colored('Driver is Initialized','green'))

# Change Download path
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_path}}
command_result = driver.execute("send_command", params)

#driver.maximize_window()
#driver.minimize_window()

# Define Web Page URL
url = "http://loadshape.epri.com/wholepremise"

#LOADING Page LOG
now = datetime.datetime.now()
print(now.strftime("%Y-%m-%d %H:%M:%S"),colored('Loading EPRI Webpage','yellow'))

# navigate to the page
driver.get(url)
now = datetime.datetime.now()
print(now.strftime("%Y-%m-%d %H:%M:%S"),colored('      ------------- EPRI Webpage Loaded Successfully -----------         ','red'))

tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[1]'

try:
    for i in range(1,56):
        now = datetime.datetime.now()
        print(now.strftime("%Y-%m-%d %H:%M:%S"),colored('City = ', 'green'),i)
        for j in range(1, 3):
            now = datetime.datetime.now()
            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Heathing Type = ', 'yellow'), j)
            if j == 1:
                heatingtype = heatingelec
            elif j == 2:
                heatingtype = heatingfuel
            for m in range(1,13):
                now = datetime.datetime.now()
                print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Month = ', 'blue'), m)
                for d in range(1,32):
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Day = ', 'green'), d)

                    # Month days corrector
                    if m == 2:
                        if d == 29:
                            now = datetime.datetime.now()
                            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('End of month: ', 'red'), m)
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[2]'
                            print('Change To date picker to Opt 2')
                            break
                    if m == 4:
                        if d == 31:
                            now = datetime.datetime.now()
                            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('End of month: ', 'red'), m)
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[2]'
                            print('Change To date picker to Opt 2')
                            break
                    if m == 6:
                        if d == 31:
                            now = datetime.datetime.now()
                            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('End of month: ', 'red'), m)
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[2]'
                            print('Change To date picker to Opt 2')
                            break
                    if m == 9:
                        if d == 31:
                            now = datetime.datetime.now()
                            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('End of month: ', 'red'), m)
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[2]'
                            print('Change To date picker to Opt 2')
                            break
                    if m == 11:
                        if d == 31:
                            now = datetime.datetime.now()
                            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('End of month: ', 'red'), m)
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[2]'
                            print('Change To date picker to Opt 2')
                            break
                    if d == 2:
                        if m > 1:
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[1]'
                            print('Change To date picker to Opt 1')

                    # Select Country
                    driver.find_element_by_xpath(countrybox).click()
                    driver.find_element_by_xpath(countryusa).click()

                    # Select City
                    driver.find_element_by_xpath(citybox).click()
                    driver.find_element_by_xpath(city[i-1]).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('City Selected: ', 'green'), i)

                    # Select Heating Type
                    driver.find_element_by_xpath(heatingtype).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Heating Type Selected: ', 'blue'), j)


                    # Select Bulding Type
                    driver.find_element_by_xpath(smalloffice).click()
                    driver.find_element_by_xpath(mediumoffice).click()
                    driver.find_element_by_xpath(largeoffice).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Bulding Type Selected', 'red'))
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Date Picker Zone ', 'red'))
                    # Open To Date Picker
                    driver.find_element_by_xpath(datetobox).click()

                    # Open To Month Selector Picker
                    driver.find_element_by_xpath(monthselector).click()

                    # Open To Month
                    driver.find_element_by_xpath(tomonthoption).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('To month option', 'green'),tomonthoption)


                    # Select To Day
                    driver.find_element_by_link_text(str(d)).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Selected Day : ', 'green'), d)


                    # Open From Date Picker
                    driver.find_element_by_xpath(datefrombox).click()

                    # Open From Month Selector Picker
                    driver.find_element_by_xpath(monthselector).click()

                    # Open From Month
                    driver.find_element_by_xpath(month[m-1]).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Selected From Month : ', 'yellow'), m)


                    # Select From Day
                    driver.find_element_by_link_text(str(d)).click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Selected From Day : ', 'blue'), d)

                    # Click Buttome
                    driver.find_element_by_xpath(addshapes).click()

                    # Download CSV File
                    driver.find_element_by_link_text('Download plot data (CSV)').click()
                    now = datetime.datetime.now()
                    print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Downloaded plot data (CSV)  ', 'green'))

                    # Remove All Button
                    driver.find_element_by_xpath(removeall).click()
                    if d == 31:
                        tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[2]'
                        now = datetime.datetime.now()
                        print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('Change To date picker to Opt 2', 'red'))
                        if m ==12:
                            tomonthoption = '//*[@id="ui-datepicker-div"]/div/div/select/option[1]'
                            print(now.strftime("%Y-%m-%d %H:%M:%S"), colored('----- End Of The Year -----', 'red'))
                            # Open From Date Picker
                            driver.find_element_by_xpath(datefrombox).click()
                            # Open From Month Selector Picker
                            driver.find_element_by_xpath(monthselector).click()
                            # Open From Month
                            driver.find_element_by_xpath( tomonthoption).click()
                            # Select From Day
                            driver.find_element_by_link_text('1').click()

finally:
    now = datetime.datetime.now()
    print( now.strftime("%Y-%m-%d %H:%M:%S"),colored('Driver Try Not Respond', 'red'))
    driver.quit()
    now = datetime.datetime.now()
    print(now.strftime("%Y-%m-%d %H:%M:%S"),colored('      ------------- Driver Closed Opening Eror Page -----------         ', 'green'))
    chromeOptions = webdriver.ChromeOptions()
    driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)
    url2='https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT8TL9rSqgQNXlXc2ZUgV0Fog0QDx2WPguCEalv5pfvpiGmoXz9'
    driver.maximize_window()
    # navigate to the page
    driver.get(url2)
    now = datetime.datetime.now()
    print(now.strftime("%Y-%m-%d %H:%M:%S"),colored('      ------------- EEEEEEEEERRRRRRRRROOOOOOORRRRR -----------         ', 'red'))