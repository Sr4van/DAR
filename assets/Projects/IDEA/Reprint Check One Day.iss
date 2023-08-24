'F&B-6-Reprint Check
'PKF Sridhar & Santhanam LLP
'Date:- 2 November 2022

Sub Main
                IgnoreWarning(True)
	Call ReportReaderImport()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Reprint Check.pdf
	Call SortDatabase()	'Check kot reconciliation report copy.IMD
	Client.CloseDatabase "Sorted database.IMD"
	Call RenameDatabase()	'Sorted database.IMD
	Call JoinDatabase()	'Reprint Check copy.IMD
	Client.CloseDatabase "Reprint Check copy.IMD"
	Call DeleteDatabase()	'Reprint Check copy.IMD
	Client.CloseDatabase "Sorted time In decending order.IMD"
	Call DeleteDatabase1()	'Sorted time In decending order.IMD
	Client.CloseDatabase "Check kot reconciliation report copy.IMD"
	Client.CloseDatabase "Check kot reconciliation report raw data.IMD"
	Client.CloseDatabase "Reprint Check raw data.IMD"
	Call JoinDatabase1()	'Reprint Checks.IMD
	Client.CloseDatabase "Reprint Checks.IMD"
	Call DeleteDatabase2()	'Reprint Checks.IMD

	

	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
End Sub


' File - Import Assistant: Report Reader
Function ReportReaderImport
	dbName = "Reprint Check raw data.IMD"
	Client.ImportPrintReportEx "C:\Users\Admin\Documents\My IDEA Documents\IDEA Projects\Continuous Monitoring System\Import Definitions.ILB\Reprint Check One Day.jpm", "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Reprint Check.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Sort time in decending order in check kot reconciliation report
' Data: Sort
Function SortDatabase
	Set db = Client.OpenDatabase("Check kot reconciliation report raw data.IMD")
	Set task = db.Sort
	task.AddKey "TIME", "D"
	dbName = "Sorted database.IMD"
	task.PerformTask dbName
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Rename Database
Function RenameDatabase
	Set ProjectManagement = client.ProjectManagement
	ProjectManagement.RenameDatabase "Sorted database.IMD", "Sorted time In decending order"
	Set ProjectManagement = Nothing
End Function

' Join reprint check with sorted time in decending order to get item desciption by taking Check no. as unique value
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Reprint Check raw data.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Sorted time In decending order.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ITEM"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Reprint Checks.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bringing User name from user list in Reprint Check by taking User iD as a unique value
' File: Join Databases
Function JoinDatabase1
	Set db = Client.OpenDatabase("Reprint Checks.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "F&B User list raw data-Sheet1.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "USER_NAME"
	task.AddMatchKey "Reprint_Cashier_ID", "USER", "A"
	dbName = "Reprint Checks Final.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' Deleting reprint check copy
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Reprint Check copy.IMD"
End Function

' Deleting sorted time in decending order 
' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Sorted time In decending order.IMD"
End Function


' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Reprint Checks.IMD"
End Function




' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Reprint Checks Final.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Reprint Checks.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function



