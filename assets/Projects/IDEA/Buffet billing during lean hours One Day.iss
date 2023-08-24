'F&B-4-Buffet timing during lean hours
'PKF Sridhar & Santhanam LLP
'Date:- 1 November 2022

Sub Main
                IgnoreWarning(True)
	Call DirectExtraction()	'Check kot reconciliation report raw data.IMD
	Call JoinDatabase()	'Buffet in Item Description.IMD
	Call DirectExtraction1()	'Cashier ID in Buffet in Time Description.IMD
	Call DirectExtraction2()	'Cashier ID in Buffet in Time Description.IMD
	Call DirectExtraction3()	'Cashier ID in Buffet in Time Description.IMD
	Call AppendDatabase()	'Dinner Buffet.IMD
	Client.CloseDatabase "Lunch Buffet.IMD"
	Call DeleteDatabase()	'Lunch Buffet.IMD
	Client.CloseDatabase "Dinner Buffet.IMD"
	Call DeleteDatabase1()	'Dinner Buffet.IMD
	Client.CloseDatabase "Breakfast Buffet.IMD"
	Call DeleteDatabase2()	'Breakfast Buffet.IMD
	Client.CloseDatabase "Cashier ID in Buffet in Time Description.IMD"
	Call DeleteDatabase3()	'Cashier ID in Buffet in Time Description.IMD
	Client.CloseDatabase "Buffet in Item Description.IMD"
	Call DeleteDatabase4()	'Buffet in Item Description.IMD
	Client.CloseDatabase "Check kot reconciliation report copy.IMD"
	Call DeleteDatabase5()	'Check kot reconciliation report copy.IMD
	Client.CloseDatabase "Check kot reconciliation report raw data.IMD"
	Client.CloseDatabase "Settlement Report raw data.IMD"
	Call JoinDatabase1()	'Buffet billing during lean hours.IMD
	Client.CloseDatabase "Buffet billing during lean hours.IMD"
	Call DeleteDatabase6()	'Buffet billing during lean hours.IMD
	Call ExportDatabaseXLSX()	'Buffet billing during lean hours final.IMD



	IgnoreWarning(False)
End Sub


' Extracting buffet from item description
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Check kot reconciliation report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Buffet in Item Description.IMD"
	task.AddExtraction dbName, "", "@Isini(""buffet"",ITEM)"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Join buffet in item description with settlement report to get cashier ID
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Buffet in Item Description.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Settlement Report raw data.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "CASHIER_ID"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Cashier ID in Buffet in Time Description.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting breakfast buffet from item description
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Cashier ID in Buffet in Time Description.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Breakfast Buffet.IMD"
	task.AddExtraction dbName, "", "@Isini(""fast"",ITEM) .AND. TIME > ""11:00:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting lunch buffet from item description
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Cashier ID in Buffet in Time Description.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Lunch Buffet.IMD"
	task.AddExtraction dbName, "", "@Isini(""lunch"",ITEM) .AND. TIME > ""17:00:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting dinner buffet from item descdription
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Cashier ID in Buffet in Time Description.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Dinner Buffet.IMD"
	task.AddExtraction dbName, "", "@Isini(""dinner"",ITEM) .AND. TIME > ""03:00:00"" .AND. TIME < ""18:00:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Append breakfast buffet with lunch and dinner buffet
' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("Breakfast Buffet.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Lunch Buffet.IMD"
	task.AddDatabase "Dinner Buffet.IMD"
	dbName = "Buffet billing during lean hours.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' Bringing User name from user list in Settlement report by taking User iD as a unique value
' File: Join Databases
Function JoinDatabase1
	Set db = Client.OpenDatabase("Buffet billing during lean hours.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "F&B User list raw data-Sheet1.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "USER_NAME"
	task.AddMatchKey "CASHIER_ID", "USER", "A"
	dbName = "Buffet billing during lean hours final.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function




' Deleting lunch buffet
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Lunch Buffet.IMD"
End Function

' Deleting dinner buffet
' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Dinner Buffet.IMD"
End Function

' Deleting breakfast buffet
' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Breakfast Buffet.IMD"
End Function

' Deleting cashier ID in buffet in time description
' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Cashier ID in Buffet in Time Description.IMD"
End Function

' Deleting buffet in item description
' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Buffet in Item Description.IMD"
End Function

' Deleting check kot reconciliation report copy
' File: Delete Database
Function DeleteDatabase5
	Client.DeleteDatabase "Check kot reconciliation report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase6
	Client.DeleteDatabase "Buffet billing during lean hours.IMD"
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Buffet billing during lean hours final.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Buffet billing during lean hours final.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function






