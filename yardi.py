#!/ms/dist/python/PROJ/core/3.4.4/bin/python

#yardi script

#first part moves files from source directory to destination then converts files from xlsx to csv
import os 
from os import path
import shutil
import pandas as pd,xlrd,glob

src = "/v/region/na/appl/msim/mesawest/data/prod/Yardi/comet/incoming"dst = "/v/region/na/appl/msim/mesawest/data/prod/Yardi/comet/incoming/Yarditextfiles"

files = [i for i in os.listdir(src) if i.startswith("Yardi") and path.isfile(path.join(src, i))]

for f in files:
	shutil.copyfile(path.join(src, f), path.join(dst, f))

#convert files in directory to csv

excel_files = glob.glob(r"/v/region/na/appl/msim/mesawest/data/prod/Yardi/comet/incoming/Yarditextfiles/*.xlsx")

for excel_file in excel_files:

	print("Converting '{}'".format(excel_file))
	try:
		df = pd.read_excel(excel_file)
		output = excel_file.split('.')[0]+'.csv'
		df.to_csv(output,index=False)
	except KeyError:
		print(" Failed to Convert")




