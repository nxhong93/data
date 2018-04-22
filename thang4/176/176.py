# coding: utf-8
 # coding: cp932
 # coding: euc-jp
import xlwings as xw
import urllib
import re
from bs4 import BeautifulSoup
import requests
import ssl
import wget
# import wget
# from selenium import webdriver



base0 = u'http://iinavi.inax.lixil.co.jp'
base = u'http://parts-shop.lixil.co.jp'
x = 2
wb = xw.Book(u'176.xlsx').sheets[0]
for row in range(2, wb.range("d1").end('down').row + 1):
    link = wb.range("d" + str(row)).value
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(link, headers={'User-Agent' : "Magic Browser"}) 
    page = urllib.request.urlopen(req, context=ctx)
    soup = BeautifulSoup(page, 'html.parser')
    hinban1 = wb.range("b" + str(row)).value
    hinban2 = wb.range("c" + str(row)).value
    hinban = [hinban1,hinban2]
    if u'http://parts-shop.lixil.co.jp' not in link:
        a0 = soup.find('strong', text = re.compile(u'|'.join(hinban)))
        if a0 != None:
            vung = a0.find_previous('table')
            
        else:
            l = 0
        anh = []
        for a1 in vung.find_all('img', src = re.compile(u'http://iinavi.inax.lixil.co.jp/prod_data/Photo/PREVIEW/')):
            link_anh = a1['src'].replace(u'PREVIEW',u'PRINT')
            anh0 = link_anh[link_anh.rfind(u'/')+1:]
            print(link_anh)
            anh.append(anh0)
            filename = wget.download(link_anh, out = anh0)
        wb.range('f' + str(row)).expand('table').value = [anh]
        for e0 in vung.find_all('strong', class_ = u'txt3')[1:]:
            e1 = e0.text.strip().replace(u'： ',u'：')
            cot0 = wb.range("A1").end('right').column+1
            dem = 0
            for cot in range(7, cot0):
                if e1.split(u'：')[0] == wb.range(1,cot).value:
                    wb.range(row,cot).options(transpose=True).value = e1.split(u'：')[1]
                    dem = 1
            if dem == 0:
                wb.range(1,cot0).options(transpose=True).value = e1.split(u'：')[0]
                wb.range(row,cot0).options(transpose=True).value = e1.split(u'：')[1]
                cot0 += 1
        vung1 = vung.find_next('table')
        so = 0
        for f0 in vung1.find_all('tr'):
            so += 1
            if u'代替品' in f0.text:
                chan = so - 1
        for f1 in vung1.find_all('tr')[:chan]:
            if u'販売期間' not in f1.text:
                if u'税別' not in f1.text:
                    f2 = f1.find_all('td')[0].text.strip()
                    f3 = f1.find_all('td')[1].text.strip()
                    cot0 = wb.range("A1").end('right').column+1
                    dem = 0
                    for cot in range(7, cot0):
                        if f2 == wb.range(1,cot).value:
                            if f3 != u'-':
                                wb.range(row,cot).options(transpose=True).value = f3
                            dem = 1
                    if dem == 0:
                        wb.range(1,cot0).options(transpose=True).value = f2
                        if f3 != u'-':
                            wb.range(row,cot0).options(transpose=True).value = f3
                        cot0 += 1
    else:
        anh = []
        for h0 in soup.find_all('td', class_ = u'goodsimg_'):
            if u'拡大画像' in h0.text:
                h1 = h0.find_next('a', href = True)
                link_anh = base + h1['href']
                anh0 = hinban1 + u'.jpg'
                anh0 = anh0.replace(u'/',u'_')
                anh.append(anh0)
                
            else:
                h1 = h0.find('img', src = True)
                if h1 != None:
                    link_anh = base + h1['src']
                    anh0 = hinban1 + u'_2.jpg'
                    anh0 = anh0.replace(u'/',u'_')
                    anh.append(anh0)
            
            print(anh0)
            filename = wget.download(link_anh, out = anh0)
            wb.range('f' + str(row)).expand('table').value = [anh]
        # for h2 in soup.find_all('th', class_ = u'catrige_goods'):
        #     h3 = h2.text.replace(u':','').replace(u' ','')
        #     if u'品番' not in h3:
        #         if u'販売価格' not in h3:
        #             cot0 = wb.range("A1").end('right').column+1
        #             dem = 0
        #             for cot in range(7, cot0):
        #                 if h3 == wb.range(1,cot).value:
        #                     if h2.find_next('td').text.strip() != u'-':
        #                         wb.range(row,cot).options(transpose=True).value = h2.find_next('td').text.strip()
        #                     dem = 1
        #             if dem == 0:
        #                 wb.range(1,cot0).options(transpose=True).value = h3
        #                 if h2.find_next('td').text.strip() != u'-':
        #                     wb.range(row,cot0).options(transpose=True).value = h2.find_next('td').text.strip()
        #                 cot0 += 1
        # h4 = soup.find(text = re.compile(u'適合商品品番'))
        # if h4 != None:
        #     h4 = h4.strip().replace(u'適合商品品番：','')
        #     cot0 = wb.range("A1").end('right').column+1
        #     dem = 0
        #     for cot in range(7, cot0):
        #         if wb.range(1,cot).value == u'適合商品品番':
        #             wb.range(row,cot).options(transpose=True).value = h4
        #             dem = 1
        #     if dem == 0:
        #         wb.range(row,cot0).options(transpose=True).value = h4
        #         wb.range(1,cot0).options(transpose=True).value = u'適合商品品番'
                        





                

    


   