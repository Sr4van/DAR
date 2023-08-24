'F&B-2-Recalled Check
'PKF Sridhar & Santhanam LLP
'Date:- 1 November 2022

Sub Main
                IgnoreWarning(True)  
	
	Call DirectExtraction()	'Recalled Check copy.IMD
	Client.CloseDatabase "Recalled Check raw data.IMD"
	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
End Sub



' Extractinf Status="VT" from recalled checks report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Recalled Check.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Recalled Check Final.IMD"
	task.AddExtraction dbName, "", "@Isin(""VT"", STATUS )"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Recalled Check Final.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Recalled Check Final.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
