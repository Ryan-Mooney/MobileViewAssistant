from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from credentials import *
import random, time, re, xlwt, datetime, os


#Used for testing purposes without needing to have MobileView
def get_asset_locations_test(username, password, assetList, root, lbl6):
    possible_floors= ['Floor Ground', 'Floor 1', 'Floor 2', 'Floor 3', 'Floor 4', 'Floor 5','Floor 6', 'Floor 7', 'Floor 8', 'Floor 9']
    floor_counter={}
    floor_list={}
    i=1
    try:
        driver=webdriver.Ie()
        driver.close()
    except:
        lbl6.config(text="Driver not set up")
        return()
    for asset in assetList:
        random_floor=possible_floors[random.randrange(len(possible_floors))]
        assetList[asset]['Location']=random_floor
        assetList[asset]['Battery Status']=str(random.randint(1, 101))+'%'
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
        #Progress Bar
        lbl6.config(text="Finding Assets..."+str(int(i/len(assetList)*100))+"% Complete")
        i+=1
        root.update()
    return(assetList, floor_counter, floor_list)
        
    

#Finding assets through admin menus

def get_asset_locations_admin(username, password, assetList, root, lbl6):
    #Initialize driver and travel to Mobileview
    driver = webdriver.Ie()
    driver.get(MV_URL)

    #Login to Mobileview
    elem = driver.find_element_by_name("j_username")
    elem.clear()
    elem.send_keys(username)
    elem = driver.find_element_by_name("j_password")
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

    #Initialize counters
    floor_counter={}
    floor_list={}
    i=1
    #Begin asset finding loop
    for asset in assetList:
        #Enter Asset Number
        try:
            element=WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "mainForm:adminAssetsSearchValue")))
        finally:
            elem=driver.find_element_by_id("mainForm:adminAssetsSearchValue")
        elem.clear()
        elem.send_keys(str(asset))
        elem=driver.find_element_by_id("mainForm:assign_assets_selection_search_button_id").click()

        #Find floor of asset and battery life
        try:
            element=WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "mainForm:adminAssetsTable:0:j_id_58")))
            floor=driver.find_element_by_id("mainForm:adminAssetsTable:0:j_id_58").text
            #Check if Floor 3 is in HPCC or Mercy Main
            if floor=="Floor 3":
                try:
                    check_floor=driver.find_element_by_xpath('//*[@title="Hall Perrine Cancer Center/Floor 3/Floor 3"]').text
                    floor="HPCC"
                except:
                    floor="Floor 3"
            #Check if Floor 1 is in Hospice or Mercy Main
            if floor =="Floor 1":
                try:
                    check_floor=driver.find_element_by_xpath('//*[@title="Hospice House/Floor 1/Floor 1"]').text
                    floor="Hospice"
                except:
                    floor="Floor 1"
            assetList[asset]['Location']=floor
            #Looks for Battery Status
            try:
                battery=driver.find_element_by_class_name("tagBatteryInListRowCss").text
                battery=battery.replace("&nbsp;", "")
                battery=battery.replace(" ","")
                if battery=="" or battery==" ":
                    battery="100%"
                assetList[asset]['Battery Status']=battery
            except:
                assetList[asset]['Battery Status']=''
        #If initial screen times out, the asset is MIA
        except:
            assetList[asset]['Location']='CNL'
            assetList[asset]['Battery Status']=''
        floor=assetList[asset]['Location']
        #Determines if the found floor is in either list, adds it if necessary, then appends the asset and increases counter
        if floor in floor_counter:
            if assetList[asset]['Type'] in floor_counter[floor]:
                floor_counter[floor][assetList[asset]['Type']]+=1
            else:
                floor_counter[floor][assetList[asset]['Type']]=1
        else:
            floor_counter[floor]={assetList[asset]['Type']: 1}
        if floor in floor_list:
            floor_list[floor].append(asset)
        else:
            floor_list[floor]=[asset]
        #Progress Bar
        lbl6.config(text="Finding Assets..."+str(int(i/len(assetList)*100))+"% Complete")
        i+=1
        root.update()
    driver.close()
    return(assetList, floor_counter, floor_list)



#Finding assets with regular menu
    
def get_asset_locations_nonadmin(username, password, assetList, root, lbl6):
    #Initialize driver and travel to Mobileview
    driver = webdriver.Ie()
    driver.get(MV_URL)

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
    
    #Initialize counters
        floor_counter={}
        floor_list={}
        i=1
    #Begin loop for pump locator
    for asset in assetList:
        #Enter Asset Number
        try:
            element=WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "mainForm:freeTextInput")))
        finally:
            elem=driver.find_element_by_id("mainForm:freeTextInput")
        elem.clear()
        elem.send_keys(str(asset))
        elem=driver.find_element_by_id("locatorSimpleSearchButton_middle").click()
    
        #Find floor of asset
        try:
            element=WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList")))
            elem=driver.find_element_by_id("mainForm:j_id_1q_5_i_2_3:3:ViewAreaChainList").text
            #Checks which building Floor 3's are in
            if floor=="Floor 3":
                check_floor=driver.find_element_by_id("mainForm:j_id_1q_5_i_2_3:2:ViewAreaChainList").text
                if check_floor=="Hall Perrine Cancer Center":
                    elem="HPCC"
            assetList[asset]['Location']=elem
            elem=driver.find_element_by_id("locatorNavSimpleButton_middle").click()
        except:
            assetList[asset]['Location']='CNL'
            elem=driver.find_element_by_id("locatorNavSimpleButton_middle").click()
        #Determines if the found floor is in either list, adds it if necessary, then appends the asset and increases counter
        if floor in floor_counter:
            if assetList[asset]['Type'] in floor_counter[floor]:
                floor_counter[floor][assetList[asset]['Type']]+=1
            else:
                floor_counter[floor][assetList[asset]['Type']]=1
        else:
            floor_counter[floor]={assetList[asset]['Type']: 1}
        if floor in floor_list:
            floor_list[floor].append(asset)
        else:
            floor_list[floor]=[asset]
        #Progress Bar
        lbl6.config(text="Finding Assets..."+str(int(i/len(assetList)*100))+"% Complete")
        i+=1
        root.update()
    driver.close()
    return(assetList, floor_counter, floor_list)

def crossCheckAssets(assetList, username, password):
    #Initialize driver and travel to RSQ
    chrome_options=Options()
    chrome_options.add_argument("--headless")
    chrome_options.binary_location = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://trimedx.service-now.com/nav_to.do?uri=%2Fwm_task_list.do%3Fsysparm_query%3Du_work_order_type%253DPreventative%2520Maintenance%255EstateNOT%2520IN3%252C-10%255Ecompany.u_cost_center%253D8380%26sysparm_first_row%3D1%26sysparm_view%3D")
    time.sleep(3)
    driver.switch_to.frame("gsft_main")

    #Login to RSQ
    elem = driver.find_element_by_id(id_="user_name")
    elem.clear()
    elem.send_keys(username)
    elem = driver.find_element_by_id(id_="user_password")
    elem.clear()
    elem.send_keys(password)
    elem = driver.find_element_by_id(id_="sysverb_login").click()
    time.sleep(3)
    driver.switch_to.frame("gsft_main")

    #Search for Asset Info by Serial Numbers
    #driver.switch_to.frame("gsft_main")
    month=datetime.datetime.today().strftime('%m')
    if month[0]=='0':
        month=float(month[1])
    if month=='1':
        lastMonth=float('12')
    else:
        lastMonth=float(str(int(month)-1))
    activeAssets={}
    for asset in assetList:
        if float(assetList[asset]['PM Month Number'])==month or float(assetList[asset]['PM Month Number'])==lastMonth:
            #This part searches the current PM List for the asset using its Serial Number
            elem=driver.find_element_by_xpath("//*[contains(@id, '_text')]")
            elem.clear()
            #print(str(assets[each]['Serial Number']))
            elem.send_keys("="+str(assetList[asset]['Serial Number']))
            elem.send_keys(Keys.ENTER)
            #Searches page for any results
            elem=driver.find_elements_by_xpath('//*[contains(@href, "u_cmdb_ci_equipment")]')
            if elem:
                activeAssets.update({asset: assetList[asset]})
            #Progress Bar
            lbl6.config(text="Cross checking for active PMs..."+str(int(i/len(assetList)*100))+"% Complete")
            i+=1
        root.update()
    driver.close()
    return(activeAssets)
                
    
