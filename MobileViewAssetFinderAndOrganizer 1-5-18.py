#Mobileview Assistant - Asset Finder and Organizer
#Developed by Ryan Mooney, BMET II, at Mercy Medical Center - Cedar Rapids, IA

#Requires IEDriverSoftware, found here: https://stackoverflow.com/questions/24925095/selenium-python-internet-explorer
    #Use 32 bit one
    #IEDriver must also be added to PATH to function properly
#Additionally requires the Selenium library for python


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Initialize variables
username="rmooney"
password="Moondawg422#"
assets=['1000000', '1000020', '1000500', '1002566', '1000007', '1000008']
possible_locations=['Floor 9', 'Floor 8', 'Floor -2', 'Floor 3', 'Floor -1', 'Floor 6']

#Initialize driver and travel to Mobileview
driver = webdriver.Ie()
driver.get("http://mobileview/asset-manager-web/core/pages/login/login.jsf")

#Login to Mobileview
elem = driver.find_element_by_name("j_username")
elem.clear()
elem.send_keys(username)
elem = driver.find_element_by_name("j_password")
elem.clear()
elem.send_keys(password)
elem.send_keys(Keys.RETURN)

#Navigate to Asset Locator
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "locatorSimpleButton_middle")))
finally:
    elem=driver.find_element_by_id("locatorSimpleButton_middle").click()

#Begin loop for pump locator
for asset in assets:
    print(asset)
    #Enter Asset Number
    try:
        element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:freeTextInput")))
    finally:
        elem=driver.find_element_by_id("mainForm:freeTextInput")
    elem.clear()
    elem.send_keys(asset)
    elem=driver.find_element_by_id("locatorSimpleSearchButton_middle").click()

    #Find floor of asset
    try:
        element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")))
        elem=driver.find_element_by_id("mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")
        print(elem.text) #Will return 'Floor 5'
        elem=driver.find_element_by_id("locatorNavSimpleButton_middle").click()
    except:
        print('CNL')
        elem=driver.find_element_by_id("locatorNavSimpleButton_middle").click()

driver.close()

#To sort a dictionary:
#import operator
#dic={1: 'ray', 2: 'bob', 3: 'felix'}
#sortdic=sorted(dic.items(), key=operator.itemgetter(1))
