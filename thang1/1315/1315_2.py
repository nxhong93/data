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
x2 = [u'4',u'6',u'10',u'16']
x4 = ['',u'B']


for y1 in x1:
    for y2 in x2:
        if y2 == u'4':
            for y3 in [u'5',u'10',u'15',u'20']:
                for y4 in x4:
                    b = u'CJP2' + y1 + y2 + u'-' + y3 + u'D-' + y4
                    b = b.replace(u'--',u'-')
                    if b[-1] == u'-':
                        b = b[:-1]
                    print (y3 + u':mm')
        elif y2 == u'6':
            for y3 in [u'5',u'10',u'15',u'20', u'25']:
                for y4 in x4:
                    b = u'CJP2' + y1 + y2 + u'-' + y3 + u'D-' + y4
                    b = b.replace(u'--',u'-')
                    if b[-1] == u'-':
                        b = b[:-1]
                    print (y3 + u':mm')
        else:
            for y3 in [u'5',u'10',u'15',u'20', u'25',u'30',u'35',u'40']:
                for y4 in x4:
                    b = u'CJP2' + y1 + y2 + u'-' + y3 + u'D-' + y4
                    b = b.replace(u'--',u'-')
                    if b[-1] == u'-':
                        b = b[:-1]
                    print (y3 + u':mm')

                    
                        

    


#        with open('933.xls', 'a') as csv_file:
 #           csv_file.write(u'\ufeff'.encode('utf8'))
  #          writer = csv.writer(csv_file, lineterminator='\n')
   #         writer.writerow([a,b,c,d,e,f,g])
#f.close()
