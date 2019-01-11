#https://stackoverflow.com/questions/24925095/selenium-python-internet-explorer
#Use 32 bit one

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Ie()
driver.get("http://www.python.org")
#assert "Python.org" in driver.title
elem = driver.find_element_by_name("q")
elem.clear()
elem.send_keys("pycon")
elem.send_keys(Keys.RETURN)
assert "No results found." not in driver.page_source

#Ensures that the page is loaded before looking for elements
try:
    element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "content")))
finally:
    conventions = driver.find_elements_by_partial_link_text("PyCon")
    results=[]
    for each in conventions:
        print(each.text)
#       text=result.text
#        results=results.append(text)
print(results)

driver.close()

#To sort a dictionary:
#import operator
#dic={1: 'ray', 2: 'bob', 3: 'felix'}
#sortdic=sorted(dic.items(), key=operator.itemgetter(1))
