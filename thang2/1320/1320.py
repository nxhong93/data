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
x = 522
wb = xw.Book(u'1320.xlsx').sheets[1]
for row in range(28, wb.range("g1").end('down').row + 1):
    link = wb.range("g" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    for z in soup.find_all('div', class_ = u'product_box'):
        a = z.find_next('h5').text.strip()
        print(a)
        b = z.find_next('p', class_ = u'cd').text.strip()
        print(b)
        z1 = z.find_next('table')
        c = z1.find_all('tr')[1].find_next('td').text.strip().replace(u'厚','').replace(u'mm',u':mm').replace(u'×',u'×')
        d = z1.find_all('tr')[1].find_next('td').find_next('td').text.strip()
        print(c)
        print(d)
        z2 = z.find_next('ul', class_ = u'tags')
        z3 = len(z2.find_all('a'))
        if z3 != 1:
            e = z2.text.replace(u'安全標識','')
        else:
            e = ''
        print(e)

    
    

        # wb.range('a' + str(x)).expand('table').value = [a,b,c,d,e,link]

    
        x += 1    
    

    
