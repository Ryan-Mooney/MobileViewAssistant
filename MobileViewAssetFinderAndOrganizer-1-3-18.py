#Mobileview Asset Finder and Organizer
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

#Initialize driver and travel to Mobileview
driver = webdriver.Ie()
driver.get("http://mobileview/asset-manager-web/core/pages/login/login.jsf")

#Login to Mobileview
elem = driver.find_element_by_name("j_username")
elem.clear()
elem.send_keys("rmooney")
elem = driver.find_element_by_name("j_password")
elem.clear()
elem.send_keys("Moondawg422#")
elem.send_keys(Keys.RETURN)

#Navigate to Asset Locator
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "locatorSimpleButton_middle")))
finally:
    elem=driver.find_element_by_id("locatorSimpleButton_middle").click()

#Enter Asset Number
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainForm:freeTextInput")))
finally:
    elem=driver.find_element_by_id("mainForm:freeTextInput")
elem.clear()
elem.send_keys("1000000")
elem.send_keys(Keys.RETURN)
elem.send_keys(Keys.RETURN)

#Find floor of asset
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")))
finally:
    elem=driver.find_element_by_id("mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")
print(elem.text) #Will return 'Floor 5'

#Ensures that the page is loaded before looking for elements
#try:
#    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "content")))
#finally:
#    conventions = driver.find_elements_by_partial_link_text("PyCon")
#    print(conventions)
#    results=[]
#    for each in conventions:
#        print(each.text)
#       text=result.text
#        results=results.append(text)


#driver.close()

#To sort a dictionary:
#import operator
#dic={1: 'ray', 2: 'bob', 3: 'felix'}
#sortdic=sorted(dic.items(), key=operator.itemgetter(1))
