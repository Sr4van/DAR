'F&B-8-Liquor sold after licensed hour
'PKF sridhar & Santhanam LLP
'Date:-2 November 2022

Sub Main
                IgnoreWarning(True)
	Call DirectExtraction()	'Item code in check kot reconciliation report.IMD
	Call JoinDatabase()
	Call DirectExtraction1()	'Item code in check kot reconciliation report.IMD	
	Call DirectExtraction2()	'Appending H Code with C Code.IMD
	Call AppendDatabase()	'C code.IMD
	Call DirectExtraction3()	'Sale of Liquor Without IRD.IMD
	Call ModifyField()	'Sale of Liquor Without IRD.IMD
	Call DirectExtraction4()	
	Client.CloseDatabase "Sale of Liquor Without IRD.IMD"
	Call DeleteDatabase()	'Sale of Liquor Without IRD.IMD
	Client.CloseDatabase "Appending H Code with C Code.IMD"
	Call DeleteDatabase1()	'Appending H Code with C Code.IMD
	Client.CloseDatabase "H Code.IMD"
	Call DeleteDatabase2()	'H Code.IMD
	Client.CloseDatabase "C code.IMD"
	Call DeleteDatabase3()	'C code.IMD
	Client.CloseDatabase "Item code in check kot reconciliation report.IMD"
	Call DeleteDatabase4()	'Item code in check kot reconciliation report.IMD
	Client.CloseDatabase "Item sale without smoke.IMD"
	Call DeleteDatabase5()	'Item sale without smoke.IMD
	Client.CloseDatabase "Check  KOT Reconciliation Report raw data.IMD"
	Client.CloseDatabase "Item Sale Report raw data.IMD"
	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
	
End Sub


' Excluding Type="SMOKE" from item sale report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Item Sale Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Item sale without smoke.IMD"
	task.AddExtraction dbName, "", "@NoMatch(TYPE,""SMOKE"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Join check kot reconciliation report with item sale report to get item code
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Check KOT Reconciliation Report raw data.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Item Sale Report raw data.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ITEM_CODE"
	task.AddMatchKey "ITEM", "ITEM_DESCRIPTION", "A"
	dbName = "Item code in check kot reconciliation report.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

'' Extracting Item Code = "H" from joined check kot reconciliation report
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Item code in check kot reconciliation report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "H Code.IMD"
	task.AddExtraction dbName, "", "@SpanIncluding(ITEM_CODE,""H"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting Item Code="C" from joijned check kot reconciliation report
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Item code in check kot reconciliation report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "C code.IMD"
	task.AddExtraction dbName, "", "@SpanIncluding(ITEM_CODE,""C"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Append H code extraction with C code extraction 
' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("H Code.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "C code.IMD"
	dbName = "Appending H Code with C Code.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Excluding Outlet="IRD" from appended report
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Appending H Code with C Code.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Sale of Liquor Without IRD.IMD"
	task.AddExtraction dbName, "", "@NoMatch(OUTLET,""IRD Bar"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Sale of Liquor Without IRD.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "TIME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting OFF timings of liquor sold 
' Data: Direct Extraction
Function DirectExtraction4
	Set db = Client.OpenDatabase("Sale of Liquor Without IRD.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Liquor Sold after licensed hour.IMD"
	task.AddExtraction dbName, "", "TIME  < ""11:00:00"" .AND. TIME > ""01:00:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting sale of liquor without IRD
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Sale of Liquor Without IRD.IMD"
End Function

' Deleting appended  H code with C code
' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Appending H Code with C Code.IMD"
End Function

' Deleting H code
' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "H Code.IMD"
End Function

' Deleting C code
' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "C code.IMD"
End Function

' Deleting Item code in check kot reconciliation report
' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Item code in check kot reconciliation report.IMD"
End Function

' Deleting Item sale without smoke
' File: Delete Database
Function DeleteDatabase5
	Client.DeleteDatabase "Item sale without smoke.IMD"
End Function


' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Liquor Sold after licensed hour.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Liquor Sold after licensed hour.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
