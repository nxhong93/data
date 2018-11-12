from pywinauto.application import Application
from pywinauto import clipboard
from xml.dom import minidom
import pywinauto
import os
import time

# path = 'C:\\Program Files (x86)\\Default Company Name\\SetupHocTiengNhat'
path = 'E:\\hong\\Debug'
base = 'C:\\Users\\Admin\\Documents\\rule\\edit_rule'
for link in os.listdir(base):
	if 'txt' in link:
		links = (base + '\\' + link).replace('\\','\\\\')
		with open(links, 'r',encoding='utf-8') as rules:
			data = rules.read()
		rules.close()
# print(data)

for files in os.listdir(path):
    if files == 'Tornado.exe':
        file_name = (path + '\\' + files).replace('\\', '\\\\')
        app = Application(backend="win32").start(file_name)
        time.sleep(1)
        app.top_window_().TypeKeys("tornado{TAB}tornado{TAB}{END}{BACKSPACE}211.109.9.60{TAB}19100{TAB 2}{ENTER}")
        time.sleep(5)
        app.top_window_().TypeKeys("{TAB 9}")
        app.top_window_().TypeKeys("{TAB 3}{DOWN}{RIGHT}{DOWN 12    }{RIGHT}{ENTER}{TAB 2}")
        time.sleep(1)


        app.top_window_().TypeKeys("{DOWN 5}{RIGHT}")
        time.sleep(1)
        app.top_window_().TypeKeys("{ENTER}{TAB 2}{DOWN 6}{RIGHT}")
        time.sleep(1)
        # app.top_window_().TypeKeys("{ENTER}{TAB 2}{DOWN 27}{RIGHT}")
        # time.sleep(1)
        # app.top_window_().TypeKeys("{ENTER}{TAB 2}{DOWN 104}{RIGHT}")
        # time.sleep(2)
        app.top_window_().TypeKeys("{ENTER}{TAB 2}{DOWN 7}")
        time.sleep(2)       






        app.top_window_().TypeKeys("{ENTER}{TAB 2}")
        time.sleep(2)
        app.top_window_().TypeKeys("{ENTER 2}")
        time.sleep(5)   
        app.top_window_().TypeKeys("{DOWN 220}{UP}")
        time.sleep(8)
        last = '1'
        url = ''
        i = 0
        while url != last:
            last = url
            app.top_window_().TypeKeys("+{F10}{DOWN 3}{ENTER}")
            time.sleep(3)
            app.top_window_().TypeKeys("{TAB 5}")
            app.top_window_().TypeKeys("^a^c")
            time.sleep(2)
            url = clipboard.GetData()
            url = url[url.find('<PageUrl>')+9:]
            url = url[:url.find('<')]
            rules_edit = data.replace('link_base', url)
            app.top_window_().TypeKeys("{TAB 4}{ENTER 2}")
            with open(os.path.join(base+'\\check.txt'), 'w',encoding='utf-8') as file_save:
                file_save.write(rules_edit)
            file_save.close()
            if 'http' in url and url != last:
                app.top_window_().TypeKeys("+{F10}{DOWN 4}{ENTER}")
                time.sleep(3)
                app.top_window_().TypeKeys("C:\\Users\\Admin\\Documents\\rule\\edit_rule\\check.txt{ENTER 2}")
                print('Done! with %s' %(i))
                time.sleep(3)
                app.top_window_().TypeKeys("{DOWN}")
                time.sleep(1)
                i += 1
            else:
                print('Error!')
        print('Done!')      