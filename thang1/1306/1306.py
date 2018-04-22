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
x = 28
wb = xw.Book(u'1306.xlsx').sheets[1]
for row in range(4, wb.range("i1").end('down').row + 1):
    link = wb.range("i" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')
    for z1 in soup.find_all('th', text = u'Cat.No.'):
        z2 = z1.find_previous('table')
        z3 = z2.find_all('tr')[0]
        for z4 in z2.find_all('tr')[1:]:
            if u'Cat.No.' in z3.text:
                for z5 in range(len(z3.find_all('th'))):
                    if z3.find_all('th')[z5].text == u'Cat.No.':
                        z6 = z5
                        break
            b = z4.find_all('td')[z6].text
            if u'液相' in z3.text:
                for z5 in range(len(z3.find_all('th'))):
                    if u'液相' in z3.find_all('th')[z5].text:
                        z6 = z5
                        break
                a = z4.find_all('td')[z6].text.replace(u'®','')
            else:
                a = soup.find(class_ = u'color_title').text.replace(u'®','')
            
           

            if u'内径' in z3.text:
                for z5 in range(len(z3.find_all('th'))):
                    if u'内径' in z3.find_all('th')[z5].text:
                        z6 = z5
                        break
                c = z4.find_all('td')[z6].text.replace(u'mm',u':mm').replace(u'ｍｍ',u':mm')
            else:
                c = ''
         

            z6 = ''
            if u'長さ' in z3.text:
                for z5 in range(len(z3.find_all('th'))):
                    if u'長さ' in z3.find_all('th')[z5].text:
                        z6 = z5
                        break
                if z6 != '':
                    d = z4.find_all('td')[z6].text.replace(u'm',u':m').replace(u'ｍ',u':m')
                else:
                    d = ''
            else:
                d = ''
           

            if u'膜厚' in z3.text:
                for z5 in range(len(z3.find_all('th'))):
                    if u'膜厚' in z3.find_all('th')[z5].text:
                        z6 = z5
                        break
                e = z4.find_all('td')[z6].text.replace(u'µm',u':µｍ').replace(u'µｍ',u':µｍ').replace(u'::',u':')
            else:
                e = ''
       

            z6 = ''
            if u'ガードカラム長さ' in z3.text:
                for z5 in range(len(z3.find_all('th'))):
                    if u'ガードカラム長さ' in z3.find_all('th')[z5].text:
                        z6 = z5
                        break
                if z6 != '':
                    f = z4.find_all('td')[z6].text.replace(u'm',u':m').replace(u'ｍ',u':m')
                else:
                    f = ''
            else:
                f = ''
        

            if u'アミンの分析' in soup.text:
                g = u'アミン類分析用'
            else:
                g = ''
            print (a)
            
    


    
        

    
    
    
            #wb.range('a' + str(x)).expand('table').value = [a,b,c,d,e,f,g,link]

    
            x += 1    
    

    
