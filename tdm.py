from openpyxl import Workbook
from openpyxl import load_workbook
import csv
import argparse
import os, sys


tdm_template = 'tdm.xlsx' # Empty Excel Template for TDMs
wb = load_workbook(tdm_template) # Loding the Excel Template to memory
files = os.listdir('.') # Listing all files in this current directory

# Validating passed command line arguments
parser = argparse.ArgumentParser()
parser.add_argument('objectname', help='The SAP object name')
parser.add_argument('csv_file', help='SAP object spacific csv file to parse in the same folder')
args = parser.parse_args()

# Verifying if the passed csv file exists in the current directory
if args.csv_file not in files:
	sys.exit('csv file is not in this directory, please place the relevant csv file in the same directory')

# Creating Sheet Names list for this particular SAP Object
sheet_names = [
	args.objectname.upper() + ' Master',
	args.objectname.upper() + ' Classification',
	args.objectname.upper() + ' Partner Function',
	args.objectname.upper() + ' Document',
	args.objectname.upper() + ' Long Text'
]

print(sheet_names)

# Assigning the created Sheet Names to the actual spreadsheet[tdm.xlsx]
if len(sheet_names) == len(wb.sheetnames):
	counter = 0
	for sheet in wb:
		sheet.title = sheet_names[counter]
		counter += 1
print(wb.sheetnames)

# Creating WorkSheet objects for: Master, Classification, Partner Function, Document and Long Text
ws_master = wb[sheet_names[0]]
ws_class = wb[sheet_names[1]]
ws_pf = wb[sheet_names[2]]
ws_doc = wb[sheet_names[3]]
wb_lt = wb[sheet_names[4]]
# Setting Master Worksheet as focused worksheet
wb.active = 0



wb.save('tdm_test.xlsx')

# with open(csv_files[0], 'r', newline='') as csv_file:
# 	csv_lines = csv.reader(csv_file, delimiter="\t")
# 	for line in csv_lines:
# 		print(line)
