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
# import wget
from selenium import webdriver




base = u'https://www.phoenixcontact.com'
x = 2
wb = xw.Book(u'md170.xlsx').sheets[0]
# driver = webdriver.Chrome('C:/Users/Admin/Downloads/Compressed/1/chromedriver_win32/chromedriver.exe')
# driver1 = webdriver.Chrome('C:/Users/Admin/Downloads/Compressed/1/chromedriver_win32/chromedriver.exe')

driver = webdriver.Firefox(executable_path='C:/Users/Admin/Downloads/Compressed/2/geckodriver.exe')

for row in range(10354, wb.range("b1").end('down').row + 1):
    
    link = wb.range("e" + str(row)).value
    
    
    driver.get(link)
    # driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')
    # driver.get(link)(By.ID,"pxc-tab-overview").send_keys(Keys.ESCAPE).perform()
    # driver.get(link).find(By.id("pxc-tab-overview")).sendKeys(Keys.ESCAPE)
    # .find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
    
    page = driver.page_source
    soup = BeautifulSoup(page, 'html.parser', from_encoding='cp932')


    def ham(cla):
        ts = soup.find(class_ = cla)
        return ts
    b0 = ham(u'pxc-benefit')
    if b0 != None:
        b = ''
        for b1 in b0.find_all('li'):
            b2 = u'●' + b1.text
            b = b + b2
    else:
        b = ''
    a = ham(u'pxc-prod-detail-txt')
    f = ''
    if a != None:
        a = a.find_next('p').text.replace(u'、 ',u'、 ').replace(u':',u':').replace(u'、 ',u'、').replace(u': ',u':').replace(u'：',u':').replace(u'\n','')
        for a1 in a.split(u'、'):
            if u':' not in a1:
                f = f + u'●' + a1
            else:
                a1_ts = a1[:a1.find(':')]
                a1_nd = a1[a1.find(':')+1:]
                ga = wb.range("A1").end('right').column+1
                gb6 = 0
                for gb5 in range(15,ga):
                    if wb.range(1,gb5).value == a1_ts:
                        wb.range(row,gb5).options(transpose=True).value = a1_nd
                        gb6 = 1
                if gb6 == 0:
                    wb.range(1,ga).options(transpose=True).value = a1_ts
                    wb.range(row,ga).options(transpose=True).value = a1_nd
                    ga += 1
    f = f + b + u'。'
    if f in [u'●。',u'。']:
        f = ''
    wb.range(row,6).options(transpose=True).value = f.replace(u'\n','')
    print(f)
    j = ham(u'cf pxc-mod-app-imagelist')
    if j != None:
        for j1 in j.find_all('img', src = True):
            j2 = j1['src']
            if u'rohs.gif' in j2:
                wb.range(row,15).options(transpose=True).value = 'Y'
    gb1 = ham(u'pxc-tbl')
    y0 = gb1.find('th', text = u'注意')
    if y0 != None:
        y = y0.find_next('td').text.strip()
        wb.range("y" + str(row)).options(transpose=True).value = y
    aa0 = gb1.find('th', text = u'GTIN')
    if aa0 != None:
        aa = aa0.find_next('td').text.strip()
        wb.range("aa" + str(row)).options(transpose=True).value = aa

    ab0 = gb1.find('th', text = re.compile(u'1個あたりの重量'))
    if ab0 != None:
        ab = ab0.find_next('td').text.strip()
        wb.range("ab" + str(row)).options(transpose=True).value = ab

    ad0 = gb1.find('th', text = u'生産国')
    if ad0 != None:
        ad = ad0.find_next('td').text.strip()
        wb.range("ad" + str(row)).options(transpose=True).value = ad           

    link1 = soup.find('base', href = re.compile(u'https://www.phoenixcontact.com/online/portal/jp/pxc/product_detail_page'))
    link1 = str(link1)[12:-3]
    c = soup.find(id = u'pxc_popup_DialogHandler_0')
    
    if c != None:
        link2 = c['href']
        link_ = link1 + link2
        driver.get(link_)
        page1 = driver.page_source
        soup1 = BeautifulSoup(page1, 'html.parser', from_encoding='cp932')

        c2 = soup1.find(class_ = u'pxc-carousel')
        if c2 != None:
            anh = []
            for c3 in c2.find_all('a', href = True):
                c4 = c3['href']
                link_anh = base + c4
                name_anh = c4.replace(u'/assets/images_pr/product_photos/large/','').replace(u'/assets/images_pr/product_drawings/large/','')
                print(name_anh)
                anh.append(name_anh)
               
            print(anh)
            wb.range('g' + str(row)).expand('table').value = [anh]