 # coding: utf-8
 # coding: cp932
 # coding: euc-jp
import urllib
import re
from bs4 import BeautifulSoup
import csv
import unicodecsv as csv
import requests
import ssl
import unicodedata
import html5lib



x1 = [u'B',u'S']
x2 = [u'4',u'6',u'10',u'15']
x3 = [u'5',u'10',u'15']
x4 = ['',u'H4', u'H6']
x5 = ['',u'B']


for y1 in x1:
    for y2 in x2:
        for y3 in x3:
            for y4 in x4:
                for y5 in x5:
                    b = u'CJP' + y1 + y2 + u'-' + y3 + y4 + u'-' + y5
                    b = b.replace(u'--',u'-')
                    if b[-1] == u'-':
                        b = b[:-1]
                    print (y2 + u':mm')
       

                    
                        

    


#        with open('933.xls', 'a') as csv_file:
 #           csv_file.write(u'\ufeff'.encode('utf8'))
  #          writer = csv.writer(csv_file, lineterminator='\n')
   #         writer.writerow([a,b,c,d,e,f,g])
#f.close()
