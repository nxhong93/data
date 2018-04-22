import shutil
import os


path0 = ".\\anh"
path = ".\\anh"
file = os.listdir(path)
for fileName in file:
	   # os.rename(os.path.join(path, fileName), os.path.join(path, fileName.replace("_3.jpg", ".jpg").replace("_2.jpg", ".jpg")))
	   # di = path + '\\' + fileName
	   # shutil.copy2(di,path0)
	   print(fileName)