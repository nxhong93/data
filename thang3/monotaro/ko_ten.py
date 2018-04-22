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
import os
import re, math
from collections import Counter

WORD = re.compile(r'\w+')

def get_cosine(text1, text2):
    count = 0
    for i in range(len(text1)):
        text0 = text1[i:i+1]
        if text0 in text2:
            count += 1
    return (count/ len(text2))


name = open('thong_so.txt', 'r').read().split('\n')
print(len(name))
wb = xw.Book(u'thu.xlsx').sheets[0]
dong = [2,9,14,78,156,337,605]
for row in dong:
    column = wb.range(row, 13).end('right').column
    for col in range(13, column + 1, 2):   
        ts = wb.range(row, col).value
        nd = wb.range(row, col + 1).value
        list_ts = []
        for na in name:
            if ts in na:
                list_ts.append(na)
        if len(list_ts) == 1:
            kq = list_ts[0]
        elif len(list_ts) > 1:
            tyle = []
            text1 = nd
            for i in list_ts:
                cosine = get_cosine(nd, i)
                tyle.append(cosine)
            if max(tyle) != 0:
                index_max = tyle.index(max(tyle))
                kq = list_ts[index_max]
            else:
                kq = ts
        else:
            do_tl = []
            for z in name:
                cosine1 = len(z)*get_cosine(ts, z)
                cosine2 = get_cosine(nd, z)
                cosine = cosine1*cosine2
                do_tl.append(cosine)
                list_ts.append(z)
            if max(do_tl) != 0:
                index_max = do_tl.index(max(do_tl))
                kq = list_ts[index_max]
            else:
                kq = ts
        print(kq)




    
    
    
    

    
