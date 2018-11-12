# coding: utf-8
 # coding: cp932
 # coding: euc-jp
import xlwings as xw
import urllib
import re
from bs4 import BeautifulSoup
import requests
import ssl
# import wget
# import wget
from selenium import webdriver



driver = webdriver.Chrome('E:/hong/python/chromedriver_win32/chromedriver.exe')

base = u'http://211.109.9.61/view.php?id='
i=1
wb = xw.Book(u'test.xlsx').sheets[0]
for row in range(2, wb.range("b2").end('down').row + 1):
    link = wb.range("b" + str(row)).value
    driver.get(link)
    while i == 1:
        i1 = driver.find_element_by_id("username")
        i1.send_keys('hongnx')
        i2 = driver.find_element_by_xpath("//input[@value='Login']").click()
        i3 = driver.find_element_by_id("password")
        i3.send_keys('123456')
        i4 = driver.find_element_by_xpath("//input[@value='Login']").click()
        i += 1
    i5 = driver.find_element_by_xpath("//input[@value='Edit']").click()
    i6 = driver.find_element_by_xpath("//select[@name='status']/option[contains(text(),'resolved')]").click()
    i7 = driver.find_element_by_xpath("//select[@name='resolution']/option[contains(text(),'fixed')]").click()
    i8 = driver.find_element_by_xpath("//select[@name='custom_field_2']/option[contains(text(),'(VDC)Review')]").click()
    # i9 = driver.find_element_by_xpath("//select[@id='custom_field_5']/option[contains(@value, 'Phase-8')]").click()
    i12 = driver.find_element_by_id("summary").get_attribute('value')
    i12 = i12[i12.find('- ')+2:]
    print(i12)
    # i11 = driver.find_element_by_id("custom_field_3")
    # i11.send_keys(i12)
    i14 = "don't need to set schedule"
    # if str(row) in [124,175]:
    #     i14 = 'update rules'
        # i16 = driver.find_element_by_xpath("//select[@name='custom_field_1']/option[text()='Structures_Change']").click()
    i13 = driver.find_element_by_id("bugnote_text")
    # i13.send_keys(i14)

    i10 = driver.find_element_by_xpath("//input[@value='Update Information']").click()

    print(row)
   

    