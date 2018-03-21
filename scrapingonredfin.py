"""
    require: python 2.7, selenium module, webdriver

"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
import urllib
import datetime
import sys
import time
import openpyxl
from openpyxl import Workbook

excelfilename = 'David - freelancer.xlsx'
book = openpyxl.load_workbook(excelfilename)
sheet = book.active
cells = sheet['A1': 'B8']

address_list = []
for c1, c2 in cells:
    address_list.append("{0:8} {1:8}".format(c1.value, c2.value))

driver = webdriver.Chrome(executable_path="./webdriver/chromedriver")
driver.get("http://www.redfin.com")
wait = WebDriverWait(driver, 5)

i = 0

for addr in address_list:
    elem = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input#search-box-input")))
    search_button = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "button.SearchButton")))
    elem.clear()
    elem.send_keys(addr)
    search_button.click()

    try:
        desc_div = wait.until(EC.visibility_of_element_located((By.ID, "house-info")))
        desc_str = desc_div.find_element_by_xpath("//div[@class='clear-fix descriptive-paragraph']/div[1]")
        description = desc_str.text
        print (description)
    except Exception as e:
        description = 'No description!'
        print("No description!")

    desc_url = driver.current_url
    print (desc_url)

    i+=1

    sheet['C'+str(i)] = description
    sheet['D'+str(i)] = desc_url

    book.save(excelfilename)

    driver.execute_script("location.href='http://www.redfin.com'")