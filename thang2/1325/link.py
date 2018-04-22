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

base = u'http://www.san-ei-web.co.jp'
x = 2
wb = xw.Book(u'1325.xlsx').sheets[1]
for row in range(2, wb.range("e1").end('down').row + 1):
    link = wb.range("e" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    z0 = soup.find('div', id = u'searchListBox')
    for z in z0.find_all('ul', class_ = u'pkg'):
        for z1 in z.find_all('a', href = True):
            z2 = base + z1['href']
            print(z2)

    
    

            wb.range('d' + str(x)).expand('table').value = [z2]

    
            x += 1    
    

    
