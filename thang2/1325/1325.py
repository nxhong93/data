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
#import wget
import os

base = u'https://www.iwata-fa.jp'
x = 2
wb = xw.Book(u'1325.xlsx').sheets[1]
for row in range(9, wb.range("d1").end('down').row + 1):
    link = wb.range("d" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    
    a = soup.find(class_ = u'proTit pkg').find_next('p').text.strip()
    print(a)
    b0 = soup.find(class_ = u'spacBox').find_all('tr')[-1]
    b = b0.find_all('td')[0].text.strip()
    print(b)
    
    c = b0.find_all('td')[-1].text.strip()
    print(c)

    wb.range('a' + str(row)).expand('table').value = [a,b,c]

    
    # x += 1    
    

    
