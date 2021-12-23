import os
import openpyxl

dir_path = os.path.dirname(os.path.realpath(__file__))

for filename in os.listdir(dir_path):
	if filename.lower().endwith((".xlsx")):
		wb = openpyxl.load_workbook(filename)
		sheet = wb["Sheet1"]
		sheet["F9"] = "=SUM(F5:F8)"
		sheet["F9"].style = "=Currency"
		wb.save(filename)