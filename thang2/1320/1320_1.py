 #!/usr/bin/python
 # -*- coding: utf-8 -*- 
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

base = u'https://apmall.jp'
x = 2
wb = xw.Book(u'1320.xlsx').sheets[1]
for row in range(44, wb.range("a1").end('down').row + 1):
    l0 = wb.range("h" + str(row)).value
    l1 = wb.range("a" + str(row)).value
    l2 = wb.range("b" + str(row)).value
    link = u'https://apmall.jp/apmall/search?text=' + urllib.parse.quote_plus(l1, encoding='utf-8') + u'+' + l2
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    z = soup.find(class_ = u'prod_name')
    if z != None:
        z = z.find_next('a', href = True)
    else:
        link = l0
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
        page = urllib.request.urlopen(req, context=ctx)
        soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
        z = soup.find(class_ = u'prod_name').find_next('a', href = True)
    z1 = base + z['href']
    req1 = urllib.request.Request(z1, headers={'User-Agent' : "Magic Browser"}) 
    page1 = urllib.request.urlopen(req1, context=ctx)
    soup1 = BeautifulSoup(page1, 'html.parser', from_encoding='cp932')
    z3 = soup1.find('td', text = u'表示')
    if z3 != None:
        k = z3.find_next('td').text.strip()
    else:
        k = ''
    print(k)

    
    

    wb.range('e' + str(row)).expand('table').value = [k] 
    

    
