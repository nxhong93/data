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

base = u'http://catalog-search.ckd.co.jp'
x = 2
wb = xw.Book(u'入力シート_CKD_画像2.xlsx').sheets[0]
for row in range(2, wb.range("a1").end('down').row + 1):
    link = wb.range("a" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    a0 = soup.find(class_ = u'col_left')
    if a0 != None:
        a = a0.find_next('img', src = True)
        a2 = a['src']
        a1 = base + a2
        a3 = a2[a2.rfind(u'/')+1:]
        wb.range("c" + str(row)).options(transpose=True).value = a3
        wb.range("d" + str(row)).options(transpose=True).value = a1
    

    
    

        # wb.range('a' + str(x)).expand('table').value = [a,b,c,d,e,link]

    
         
    

    
