'''
This script copies all of the spread sheet data from the individual sheets into the master sheets
'''

import os
import pandas as pd


source_directory = 'C:/Users/sebas/OneDrive/Software/Python/Excel_Auto_3/Correctly_Arranged_Sheets'
files = os.listdir(source_directory)
print(files)

for filename in os.listdir(source_directory):
	name = os.path.abspath(filename)
	print(name)
	

'''
df = pd.DataFrame()

source_directory = 'C:/Users/sebas/OneDrive/Software/Python/Excel_Auto_3/Correctly_Arranged_Sheets'
for path in os.listdir(source_directory):
	full_path = os.path.join(source_directory, path)
	if os.path.isfile(full_path):
		print(full_path)
		df = df.append(pd.read_excel(full_path), ignore_index=True)
	df.to_excel('master_test.xlsx')

df = pd.DataFrame()

for (dirname, dirs, files) in os.walk('C:/Users/sebas/OneDrive/Software/Python/Excel_Auto_3/Correctly_Arranged_Sheets/'):  #loop through Excel Files
	for file in files:
		if file.endswith('.xlsx'):
			df = df.append(pd.read_excel(file), ignore_index=True)

	df.to_excel('master_test.xlsx')

'''
