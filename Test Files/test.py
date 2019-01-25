from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random, time

assetList={'1000119'}

driver = webdriver.Ie()
driver.get("http://mobileview/asset-manager-web/core/pages/login/login.jsf")

#Login to Mobileview
elem = driver.find_element_by_name("rmooney")
elem.clear()
elem.send_keys(username)
elem = driver.find_element_by_name("Moondawg422#")
elem.clear()
elem.send_keys(password)
elem.send_keys(Keys.RETURN)

#Navigate to Admin mode
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "systemMenu")))
finally:
    elem=driver.find_element_by_id("systemMenu").click()
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainForm:menuSubview:j_id_y_19:0:j_id_y_1d")))
finally:
    elem=driver.find_element_by_id("mainForm:menuSubview:j_id_y_19:0:j_id_y_1d").click()

floor_counter={}
floor_list={}
i=1

#Enter Asset Number
try:
    element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:adminAssetsSearchValue")))
finally:
    elem=driver.find_element_by_id("mainForm:adminAssetsSearchValue")
elem.clear()
elem.send_keys(str('1000119'))
elem=driver.find_element_by_id("mainForm:assign_assets_selection_search_button_id").click()

#Find floor of asset and battery life
try:
    element=WebDriverWait(driver, 1.5).until(EC.presence_of_element_located((By.ID, "mainForm:adminAssetsTable:0:j_id_58")))
    floor=driver.find_element_by_id("mainForm:j_id_1q_5_i_2_3:2:ViewAreaChainList").text
    #Check if Floor 3 is in HPCC or Mercy Main
    assetList[asset]['Location']=floor
    try:
        battery=driver.find_element_by_class_name("tagBatteryInListRowCss").text
        battery=battery.replace("&nbsp;", "")
        battery=battery.replace(" ","")
        if battery=="" or battery==" ":
            battery="100%"
        assetList[asset]['Battery Status']=battery
    except:
        assetList[asset]['Battery Status']=''
except:
    assetList[asset]['Location']='CNL'
    assetList[asset]['Battery Status']=''
floor=assetList[asset]['Location']

