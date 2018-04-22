import shutil
import os


path0 = ".\\1_270"
path = ".\\1_270"
file = os.listdir(path)
for fileName in file:
	   # os.rename(os.path.join(path, fileName), os.path.join(path, fileName.replace("(", "").replace(")", "")))
	   # di = path + '\\' + fileName
	   # shutil.copy2(di,path0)
	   print(fileName)