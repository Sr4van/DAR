'F&B-1-Adherence to happy hour
'PKF Sridhar & Santhanam LLP
'Date:-1 November 2022

Sub Main
                IgnoreWarning(True)
	Call DirectExtraction()	'Discount Report raw data.IMD
	Call DirectExtraction1()	'Discount Report copy.IMD
	Call ReportReaderImport1()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Check kot reconciliation report.pdf
	Call DirectExtraction2()	'Check kot reconciliation report raw data.IMD
	Call JoinDatabase()	'Happy Hour.IMD
	Call DirectExtraction3()	'Time in Happy Hour.IMD
	Client.CloseDatabase "Time in Happy Hour.IMD"
	Client.CloseDatabase "Check kot reconciliation report copy.IMD"
	Client.CloseDatabase "Check kot reconciliation report raw data.IMD"
	Client.CloseDatabase "Happy Hour.IMD"
	Client.CloseDatabase "Discount Report copy.IMD"
	Client.CloseDatabase "Discount Report raw data.IMD"
	Call DeleteDatabase()	'Time in Happy Hour.IMD
	Call DeleteDatabase1()	'Happy Hour.IMD
	Call DeleteDatabase2()	'Discount Report copy.IMD
	Call DeleteDatabase3()	'Check kot reconciliation report copy.IMD
	Call ExportDatabaseXLSX()
	IgnoreWarning(False)
End Sub



' Creating a copy of Discount Report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Discount Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Discount Report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting happy hour from discount report
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Discount Report copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Happy Hour.IMD"
	task.AddExtraction dbName, "", "@Isini(""happy"",  GUEST_DETAILS )"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Import Check kot reconciliation report
' File - Import Assistant: Report Reader
Function ReportReaderImport1
	dbName = "Check kot reconciliation report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Check kot reconciliation report One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Check  Kot Reconciliation Report.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Creating a copy of Check kot reconciliation report 
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Check kot reconciliation report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Check kot reconciliation report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Join Extracted happy hour with check kot reconciliation report to get kot punch time 
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Happy Hour.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Check kot reconciliation report copy.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "TIME"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Time in Happy Hour.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting off timings of happy hours 
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Time in Happy Hour.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Off timing of Happy Hour.IMD"
	task.AddExtraction dbName, "", " TIME  < ""15:00:00"" .AND.   TIME  > ""21:00:00"" .AND.  @Match( OUTLET ,""AURA BAR"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting time in happy hour 
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Time in Happy Hour.IMD"
End Function

' Deleting happy hour
' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Happy Hour.IMD"
End Function

' Deleting discount report copy
' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Discount Report copy.IMD"
End Function

' Deleting check kot reconciliation report
' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Check kot reconciliation report copy.IMD"
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Off timing of Happy Hour.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Off timing of Happy Hour.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
