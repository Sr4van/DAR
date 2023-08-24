Sub Main
IgnoreWarning (True)
Call ReportReaderImport()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Discount Report.pdf
Call ReportReaderImport1()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Settlement Report.pdf
Call ReportReaderImport2()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Check Time Scroll Report.pdf
Call ReportReaderImport3()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Void Check Report.pdf
Call ReportReaderImport4()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Void Item Report.pdf
Call ReportReaderImport5()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Item Sale Report.pdf
Call ReportReaderImport6()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Transferred Item Report.pdf
Call ExcelImport()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\F&B User list.xlsx
Call ExcelImport1()	'C:\Users\Admin\Desktop\CMS- IDEA\Front Office\Attendance.xlsx
Call ReportReaderImport7()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Check  KOT Reconciliation Report.pdf
Call ReportReaderImport8()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Recalled Checks.pdf
	Client.CloseDatabase "Recalled Check.IMD"
	Client.CloseDatabase "Check KOT Reconciliation Report raw data.IMD"
Client.CloseDatabase "Attendance raw data-Attendance Details By Date.IMD"
	Client.CloseDatabase "F&B User list raw data-Sheet1.IMD"
	Client.CloseDatabase "Transferred Item Report raw data.IMD"
	Client.CloseDatabase "Item Sale Report raw data.IMD"
	Client.CloseDatabase "Void Item Report raw data.IMD"
	Client.CloseDatabase "Void Check Report raw data.IMD"
	Client.CloseDatabase "Check Time Scroll Report raw data.IMD"
	Client.CloseDatabase "Settlement Report raw data.IMD"
	Client.CloseDatabase "Discount Report raw data.IMD"
	Client.CloseDatabase "Attendance-Sheet1.IMD"




IgnoreWarning (False)

End Sub




' Import Discount Report 
' File - Import Assistant: Report Reader
Function ReportReaderImport
	dbName = "Discount Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Discount Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Discounts   The Claridges.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function


' Import settlement report
' File - Import Assistant: Report Reader
Function ReportReaderImport1
	dbName = "Settlement Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Settlement Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Settlements.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function


' File - Import Assistant: Report Reader
Function ReportReaderImport2
	dbName = "Check Time Scroll Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Check Time Scroll Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Check TIME Scroll.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Import Void Check report 
' File - Import Assistant: Report Reader
Function ReportReaderImport3
	dbName = "Void Check Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Void Check Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Void Checks.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Import Void item report
' File - Import Assistant: Report Reader
Function ReportReaderImport4
	dbName = "Void Item Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Void Item Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Void Items.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Import item sale report
' File - Import Assistant: Report Reader
Function ReportReaderImport5
	dbName = "Item Sale Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Item Sale Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Item Sales Report.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function



' Import transferred item report
' File - Import Assistant: Report Reader
Function ReportReaderImport6
	dbName = "Transferred Item Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Transferred Item Report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Transferred Items.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Import F&B User list
' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\F&B User list.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "F&B User list raw data"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function

'Importing attendance
' File - Import Assistant: Excel
Function ExcelImport1
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Front Office\Attendance.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "Attendance"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function

' File - Import Assistant: Report Reader
Function ReportReaderImport7
	dbName = "Check KOT Reconciliation Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Check kot reconciliation report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Check  KOT Reconciliation Report.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function


' File - Import Assistant: Report Reader
Function ReportReaderImport8
	dbName = "Recalled Check.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Recalled Check One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Recalled Checks.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function


