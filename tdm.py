from openpyxl import Workbook
from openpyxl import load_workbook
import csv

args = ['PROJ MASTER', 'PROJ Classification', 'PROJ Partner Function', 'PROJ Document', 'PROJ Long Text']
tdm_template = 'tdm.xlsx'
csv_files = ['PS_MASTER.csv']

wb = load_workbook(tdm_template)

if len(args) == len(wb.sheetnames):
	counter = 0
	for sheet in wb:
		sheet.title = args[counter]
		counter += 1

# five worksheets: Master, Classification, Partner Function, Document and Long Text
ws_master = wb[args[0]]
ws_class = wb[args[1]]
ws_pf = wb[args[2]]
ws_doc = wb[args[3]]
wb_lt = wb[args[4]]

# Setting Master Worksheet as focused worksheet
wb.active = 0


wb.save('tdm_test.xlsx')

# with open(csv_files[0], 'r', newline='') as csv_file:
# 	csv_lines = csv.reader(csv_file, delimiter="\t")
# 	for line in csv_lines:
# 		print(line)
