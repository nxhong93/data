 # coding: utf-8
 # coding: cp932
 # coding: euc-jp
import xlwings as xw
import urllib
import re
from bs4 import BeautifulSoup
import csv
import unicodecsv as csv
import requests
import ssl
import unicodedata
import html5lib
import wget
import os
from selenium import webdriver

driver = webdriver.Chrome('C:/Users/Admin/Downloads/Compressed/1\chromedriver_win32/chromedriver.exe')


base = u'http://www.bdbiosciences.com'
wb = xw.Book(u'linkweb124_anh.xlsx').sheets[0]
for row in range(2, wb.range("c1").end('down').row + 1):
    link = wb.range("c" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    #req = urllib2.Request(link) 
    #page = urllib2.urlopen(req, context=ctx)
    driver.get(link)
    page = driver.page_source
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    z1 = soup.find('div', class_ = u'img-container ng-scope')
    if z1 == None:
        k = ''
    else:
        k = []
        for z2 in soup.find_all('div', class_ = u'img-container ng-scope'):
            z3 = z2.find_next('img', src = True)
            z4_ = z3['src']
            
            z4 = z4_.replace(u'//doc.coromant.sandvik.com/product/tibp_pic/preview/','')
            filename = wget.download(u'https:' + z4_, out = z4)
            k.append(z4)
            wb.range('e' + str(row)).expand('table').value = [k]
    y1 = soup.find('span', text = u'(CW)')
    if y1 != None:
        h = y1.find_next('strong').text.strip()
        wb.range('d' + str(row)).value = h
    else:
        h = ''
    print (h)
    print (k)
            
    


 
    

    
