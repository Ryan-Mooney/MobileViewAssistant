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
import random

#Initialize variables
username="rmooney"
password="Moondawg422#"
assetList=['1000000', '1000020', '1000500', '1002566', '1000007', '1005848']
possible_locations=['Floor 9', 'Floor 8', 'Floor -2', 'Floor 3', 'Floor -1', 'Floor 6']


#Used for testing purposes without needing to have MobileView
def get_asset_locations(username, password, assetList):
    possible_floors= ['Floor Ground', 'Floor 1', 'Floor 2', 'Floor 3', 'Floor 4', 'Floor 5','Floor 6', 'Floor 7', 'Floor 8', 'Floor 9']
    floor_counter={}
    floor_list={}
    for asset in assetList:
        random_floor=possible_floors[random.randrange(len(possible_floors))]
        assetList[asset]['Location']=random_floor
        if random_floor in floor_counter:
            if assetList[asset]['Type'] in floor_counter[random_floor]:
                floor_counter[random_floor][assetList[asset]['Type']]+=1
            else:
                floor_counter[random_floor][assetList[asset]['Type']]=1
        else:
            floor_counter[random_floor]={assetList[asset]['Type']: 1}
        if random_floor in floor_list:
            floor_list[random_floor].append(asset)
        else:
            floor_list[random_floor]=[asset]
    return(assetList, floor_counter, floor_list)
        
    

###Finding assets through admin menus
##def get_asset_locations(username, password, assetList):
##    #Initialize driver and travel to Mobileview
##    driver = webdriver.Ie()
##    driver.get("http://mobileview/asset-manager-web/core/pages/login/login.jsf")
##
##    #Login to Mobileview
##    elem = driver.find_element_by_name("j_username")
##    elem.clear()
##    elem.send_keys(username)
##    elem = driver.find_element_by_name("j_password")
##    elem.clear()
##    elem.send_keys(password)
##    elem.send_keys(Keys.RETURN)
##
##    #Navigate to Admin mode
##    try:
##        element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "systemMenu")))
##    finally:
##        elem=driver.find_element_by_id("systemMenu").click()
##    try:
##        element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mainForm:menuSubview:j_id_y_19:0:j_id_y_1d")))
##    finally:
##        elem=driver.find_element_by_id("mainForm:menuSubview:j_id_y_19:0:j_id_y_1d").click()
##
##    #Initialize counters
##    floor_counter={}
##    floor_list={}
##
##    #Begin asset finding loop
##    for asset in assetList:
##        #Enter Asset Number
##        try:
##            element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:adminAssetsSearchValue")))
##        finally:
##            elem=driver.find_element_by_id("mainForm:adminAssetsSearchValue")
##        elem.clear()
##        elem.send_keys(str(asset))
##        elem=driver.find_element_by_id("mainForm:assign_assets_selection_search_button_id").click()
##
##        #Find floor of asset
##        try:
##            element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:adminAssetsTable:0:j_id_58")))
##            floor=driver.find_element_by_id("mainForm:adminAssetsTable:0:j_id_58")
##            assetList[asset]['Location']=floor
##        except:
##            assetList[asset]['Location']='CNL'
##        floor=assetList[asset]['Location']
##        #Determines if the found floor is in either list, adds it if necessary, then appends the asset and increases counter
##        if floor in floor_counter:
##            if assetList[asset]['Type'] in floor_counter[floor]:
##                floor_counter[floor][assetList[asset]['Type']]+=1
##            else:
##                floor_counter[floor][assetList[asset]['Type']]=1
##        else:
##            floor_counter[floor]={assetList[asset]['Type']: 1}
##        if floor in floor_list:
##            floor_list[floor].append(asset)
##        else:
##            floor_list[floor]=[asset]
##    driver.close()
##    return(assetList, floor_counter, floor_list)


    
##
###Finding assets with regular menu
    
##def get_asset_locations(username, password, assetList):
    #Initialize driver and travel to Mobileview
##    driver = webdriver.Ie()
##    driver.get("http://mobileview/asset-manager-web/core/pages/login/login.jsf")
##
##    #Login to Mobileview
##    elem = driver.find_element_by_name("j_username")
##    elem.clear()
##    elem.send_keys(username)
##    elem = driver.find_element_by_name("j_password")
##    elem.clear()
##    elem.send_keys(password)
##    elem.send_keys(Keys.RETURN)

    ###Navigate to Asset Locator
    ##try:
    ##    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "locatorSimpleButton_middle")))
    ##finally:
    ##    elem=driver.find_element_by_id("locatorSimpleButton_middle").click()
    ##    
    ###Begin loop for pump locator
    ##for asset in assetList:
    ##    #Enter Asset Number
    ##    try:
    ##        element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:freeTextInput")))
    ##    finally:
    ##        elem=driver.find_element_by_id("mainForm:freeTextInput")
    ##    elem.clear()
    ##    elem.send_keys(str(asset))
    ##    elem=driver.find_element_by_id("locatorSimpleSearchButton_middle").click()
    ##
    ##    #Find floor of asset
    ##    try:
    ##        element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")))
    ##        elem=driver.find_element_by_id("mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")
    ##        assetList[asset]['Location']=elem
    ##        elem=driver.find_element_by_id("locatorNavSimpleButton_middle").click()
    ##    except:
    ##        assetList[asset]['Location']='CNL'
    ##        elem=driver.find_element_by_id("locatorNavSimpleButton_middle").click()
    ##driver.close()

#To sort a dictionary:
#import operator
#dic={1: 'ray', 2: 'bob', 3: 'felix'}
#sortdic=sorted(dic.items(), key=operator.itemgetter(1))
