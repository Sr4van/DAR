'F&B-3-Pax not updated
'PKF Sridhar & Santhanam LLP
'Date:- 1 November 2022

Sub Main
                IgnoreWarning(True)
	Call ReportReaderImport()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Time Difference Report.pdf
	Call JoinDatabase()	'Time Difference Report copy.IMD
	Call DirectExtraction2()	'Item Description in Time Difference Report.IMD
	Call DirectExtraction3()	'Item Description in Time Difference Report.IMD
	Call ModifyField()	'Buffet in Time Difference Report.IMD
	Call AppendField()	'Buffet in Time Difference Report.IMD
	Call DirectExtraction4()	'Buffet in Time Difference Report.IMD
	Client.CloseDatabase "Buffet in Time Difference Report.IMD"
	Client.CloseDatabase "Covers not updated.IMD"
	Client.CloseDatabase "Item Description in Time Difference Report.IMD"
	Client.CloseDatabase "Check kot reconciliation report copy.IMD"
	Client.CloseDatabase "Check kot reconciliation report raw data.IMD"
	Client.CloseDatabase "Time Difference Report copy.IMD"
	Client.CloseDatabase "Time Difference Report raw data.IMD"
	Call DeleteDatabase()	'Buffet in Time Difference Report.IMD
	Call DeleteDatabase1()	'Item Description in Time Difference Report.IMD
	Call DeleteDatabase2()	'Time Difference Report copy.IMD
	Call DeleteDatabase3()	'Check kot reconciliation report copy.IMD
	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
End Sub


' File - Import Assistant: Report Reader
Function ReportReaderImport
	dbName = "Time Difference Report raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Time Difference Report.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Time Difference Report.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Join time difference report with check kot reconciliation report to get item, qty, and amount 
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Time Difference Report raw data.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Check kot reconciliation report raw data.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ITEM"
	task.AddSFieldToInc "QTY"
	task.AddSFieldToInc "AMOUNT"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Item Description in Time Difference Report.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Exytracting Cover="0" from time difference report
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Item Description in Time Difference Report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Covers not updated.IMD"
	task.AddExtraction dbName, "", "@Match(COVERS,""0"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting buffet from item descriptiion 
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Item Description in Time Difference Report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Buffet in Time Difference Report.IMD"
	task.AddExtraction dbName, "", "@Isini(""buffet"",ITEM)"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modifying type of cover to number
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Buffet in Time Difference Report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "COVERS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "COVERS", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Add field computing difference in covers
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Buffet in Time Difference Report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DIFFERENCE_IN_COVERS"
	field.Description = "Added field"
	field.Type = WI_VIRT_NUM
	field.Equation = " COVERS  - QTY"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting difference in cover not equal to 0
' Data: Direct Extraction
Function DirectExtraction4
	Set db = Client.OpenDatabase("Buffet in Time Difference Report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Incorrect Pax in Buffet.IMD"
	task.AddExtraction dbName, "", "DIFFERENCE_IN_COVERS <> 0"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting buffet in time difference report
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Buffet in Time Difference Report.IMD"
End Function

' Deleting item description in time difference report 
' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Item Description in Time Difference Report.IMD"
End Function

' Deleting time diference report copy
' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Time Difference Report copy.IMD"
End Function

' Deleting check kot reconcilition rpeort copy
' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Check kot reconciliation report copy.IMD"
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Incorrect Pax in Buffet.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Incorrect Pax in Buffet.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
