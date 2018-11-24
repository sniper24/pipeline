import win32com.client
win32c = win32com.client.constants
import pandas as pd
'''
1. load_data
	just load data from csv to df, get a schema list
2. build_cube
	aggragate by different levels and columns then 
'''

def connect_csv(wb, from_path, to_sheet, start_cell=[1,1]):
	Sheet1 = wb.Worksheets(to_sheet)

	# load csv to a nested list
	df = pd.read_csv(from_path)
	TestData =  [df.columns.tolist()] + df.values.tolist()
	
	for i, TestDataRow in enumerate(TestData):
		for j, TestDataItem in enumerate(TestDataRow):
			Sheet1.Cells(i+start_cell[0],j+start_cell[1]).Value = TestDataItem

	cl1 = Sheet1.Cells(start_cell[0],start_cell[1])
	cl2 = Sheet1.Cells(start_cell[0]+len(TestData)-1,start_cell[1]+len(TestData[0])-1)
	PivotSourceRange = Sheet1.Range(cl1,cl2)

	PivotSourceRange.Select()
	return PivotSourceRange

def build_cube(from_df, to_csv, group_by_cols, max_cols, sum_cols):
	#from_df

	# group_by_cols
	# max(max_cols)
	# sum(sum_cols)

	pass


def pivot(wb, PivotTableName, PivotSourceRange, filters, cols, rols, fields):# PivotTargetRange, 
	# Add a new worksheet
	wb.Sheets.Add (After=wb.Sheets(wb.Worksheets.Count))
	sheet_new = wb.Worksheets(wb.Worksheets.Count)
	sheet_new.Name = PivotTableName

	cl3=sheet_new.Cells(len(filters)+2,1)
	PivotTargetRange = sheet_new.Range(cl3,cl3)

	PivotCache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)
	PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)

	print "Generating PivotTable: [{}]|ws.count:[{}]|filters:[{}]|cols:[{}]|rols:[{}]|fields:[{}]".format(PivotTableName, wb.Worksheets.Count, "|".join(filters),"|".join(cols),"|".join(rols),"|".join(fields))

	# wb.Sheets.Add (After=wb.Sheets(wb.Worksheets.Count))
	# sheet_new = wb.Worksheets(2)
	# sheet_new.Name = 'Hello'

	for i in range(len(filters)):
		PivotTable.PivotFields(filters[i]).Orientation = win32c.xlPageField
		PivotTable.PivotFields(filters[i]).Position = i+1
	for i in range(len(rols)):
		PivotTable.PivotFields(rols[i]).Orientation = win32c.xlRowField
		PivotTable.PivotFields(rols[i]).Position = i+1
		# PivotTable.PivotFields(rols[i]).Name = 'ColX'
	for i in range(len(cols)):
		PivotTable.PivotFields(cols[i]).Orientation = win32c.xlColumnField
		PivotTable.PivotFields(cols[i]).Position = i+1
		PivotTable.PivotFields(cols[i]).Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
	for i in range(len(fields)):
		DataField = PivotTable.AddDataField(PivotTable.PivotFields(fields[i]))
		# , "Sum of Uptime", xlSum
		# DataField.NumberFormat = '#\'##0.00'
		DataField.NumberFormat = '###0.00'
		DataField.Name = 'Total Amt'
		DataField.Function = win32c.xlSum#xlCount #win32c.xlAverage # win32c.xlCount #win32c.xlSum
		# DataField.Function
		# https://docs.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction


Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
Excel.Visible = 1# 0

wb = Excel.Workbooks.Add()
print "wb.sheets.count:{}".format(wb.Worksheets.Count)
Sheet1 = wb.Worksheets("Sheet1")

PivotSourceRange = connect_csv(wb, r"D:/GitRepo/pipeline/dev/test.csv", "Sheet1", start_cell = [4,2])
pivot(wb, 'PivotTable1', PivotSourceRange, filters=["Country", "Gender"], cols=["Sign"], rols=["Name"], fields=["Amount"])
pivot(wb, 'PivotTable2', PivotSourceRange, filters=["Sign", "Gender"], cols=["Country"], rols=["Name"], fields=["Amount"])
pivot(wb, 'PivotTable3', PivotSourceRange, filters=["Country"], cols=["Sign", "Gender"], rols=["Name"], fields=["Amount"])
pivot(wb, 'PivotTable4', PivotSourceRange, filters=["Sign"], cols=["Country", "Gender"], rols=["Name"], fields=["Amount"])
pivot(wb, 'PivotTable5', PivotSourceRange, filters=["Gender"], cols=["Country"], rols=["Name"], fields=["Amount", "Amount"])
Excel.DisplayAlerts = False
wb.Worksheets("Sheet1").Delete()
wb.SaveAs(r'D:\GitRepo\pipeline\dev\output.xlsx')

Excel.DisplayAlerts = True
Excel.Application.Quit()