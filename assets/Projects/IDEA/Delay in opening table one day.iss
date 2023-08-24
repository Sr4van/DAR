'F&B-9-Delay in opening table
'PKF Sridhar & Santhanam LLP
'Date:- 2 November 2022

Sub Main
                IgnoreWarning(True)
	Call JoinDatabase()	'Check kot reconciliation report copy.IMD
	Call JoinDatabase1()	'Last settlement time in kot report.IMD
	Call SortDatabase()	'Settlement Report raw data.IMD
	Call JoinDatabase2()	'Item code in kot report.IMD
	Call AppendField()	'Cashier ID in kot report.IMD
	Call ModifyField()	'Cashier ID in kot report.IMD
	Call ModifyField1()	'Cashier ID in kot report.IMD
	Call DirectExtraction()	'Cashier ID in kot report.IMD
	Call ModifyField3()	'Cashier ID in kot report copy.IMD
	Call AppendField1()	'Cashier ID in kot report copy.IMD
	Call DirectExtraction1()	'Cashier ID in kot report copy.IMD
	Call DirectExtraction2()	'Time Difference less than 15 min.IMD
	Call DirectExtraction3()	'Delay in opening table.IMD
	Client.CloseDatabase "Delay in opening table.IMD"
	Call DeleteDatabase()	'Delay in opening table.IMD
	Client.CloseDatabase "Time Difference less than 15 min.IMD"
	Call DeleteDatabase1()	'Time Difference less than 15 min.IMD
	Client.CloseDatabase "Cashier ID in kot report copy.IMD"
	Call DeleteDatabase2()	'Cashier ID in kot report copy.IMD
	Client.CloseDatabase "Cashier ID in kot report.IMD"
	Call DeleteDatabase3()	'Cashier ID in kot report.IMD
	Client.CloseDatabase "Item code in kot report.IMD"
	Call DeleteDatabase4()	'Item code in kot report.IMD
	Client.CloseDatabase "Last settlement time in kot report.IMD"
	Call DeleteDatabase5()	'Last settlement time in kot report.IMD
	Client.CloseDatabase "Check kot reconciliation report copy.IMD"
	Call DeleteDatabase6()	'Check kot reconciliation report copy.IMD
	Client.CloseDatabase "Check Time Scroll Report copy.IMD"
	Call DeleteDatabase7()	'Check Time Scroll Report copy.IMD
	Client.CloseDatabase "Item Sale Report copy.IMD"
	Call DeleteDatabase8()	'Item Sale Report copy.IMD
	Client.CloseDatabase "Settlement report copy.IMD"
	Call DeleteDatabase9()	'Settlement report copy.IMD
	Client.CloseDatabase "Settlement Report raw data.IMD"
	Client.CloseDatabase "Item Sale Report raw data.IMD"
	Client.CloseDatabase "Check Time Scroll Report raw data.IMD"
	Client.CloseDatabase "Check kot reconciliation report raw data.IMD"
	Call DirectExtraction4()	'Delay in opening table final.IMD
	Call JoinDatabase3()	'Delay in opening table final.IMD
	Call RemoveField()	'Delay in Opening Table.IMD
	Client.CloseDatabase "Buffet in Item.IMD"
	Client.CloseDatabase "Delay in opening table final.IMD"
	Call DeleteDatabase10()	'Buffet in Item.IMD
	Call DeleteDatabase11()	'Delay in opening table final.IMD
	Call RemoveField1()	'Delay in Opening Table.IMD
	Client.CloseDatabase "Settlement Time In Descending Order.IMD"
	Call DeleteDatabase12()	'Settlement Time In Descending Order.IMD
	Call JoinDatabase4()	'Delay in Opening Table Exception.IMD


		Call ExportDatabaseXLSX()

                IgnoreWarning(False)
End Sub



' Join check kot reconciliation report with check time scroll report to get last settlement time 
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Check Kot Reconciliation Report raw data.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Check Time Scroll Report raw data.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "LAST_SETTLEMENT_TIME"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Last settlement time in kot report.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' Join check kot reconciliation having last settlement time with item sale report to get item code
' File: Join Databases
Function JoinDatabase1
	Set db = Client.OpenDatabase("Last settlement time in kot report.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Item Sale Report raw data.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ITEM_CODE"
	task.AddMatchKey "ITEM", "ITEM_DESCRIPTION", "A"
	dbName = "Item code in kot report.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Sort
Function SortDatabase
	Set db = Client.OpenDatabase("Settlement Report raw data.IMD")
	Set task = db.Sort
	task.AddKey "TIME", "D"
	dbName = "Settlement Time In Descending Order.IMD"
	task.PerformTask dbName
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Join check kot reconciliation having last settlement time and item code with settlement report to get cashier id and payment mode
' File: Join Databases
Function JoinDatabase2
	Set db = Client.OpenDatabase("Item code in kot report.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Settlement Time In Descending Order.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "CASHIER_ID"
	task.AddSFieldToInc "PAYMENT_MODE"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Cashier ID in kot report.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Adding field with time = "24:00:00"
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Cashier ID in kot report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_IN_24_HR"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = """24:00:00"""
	field.Length = 26
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modifying field type of time from character to time
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Cashier ID in kot report.IMD")
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

' Modifying field type of last settlement time from character to time
' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("Cashier ID in kot report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "LAST_SETTLEMENT_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "LAST_SETTLEMENT_TIME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Creating a copy of cashier ID in Kot report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Cashier ID in kot report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Cashier ID in kot report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modifying field type of 24 Hr time from character to time
' Modify Field
Function ModifyField3
	Set db = Client.OpenDatabase("Cashier ID in kot report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_IN_24_HR"
	field.Description = "Added field"
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM:SS"
	task.ReplaceField "TIME_IN_24_HR", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Add field to get the time difference between last settlement time and kot punch time 
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("Cashier ID in kot report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_DIFFERENCE"
	field.Description = "Added field"
	field.Type = WI_VIRT_TIME
	field.Equation = "@If(LAST_SETTLEMENT_TIME >= TIME,LAST_SETTLEMENT_TIME-TIME,LAST_SETTLEMENT_TIME + TIME_IN_24_HR-TIME)"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting time difference less than 15 min.
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Cashier ID in kot report copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Time Difference less than 15 min.IMD"
	task.AddExtraction dbName, "", "TIME_DIFFERENCE < ""00:15:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Excluding  Item code start with letter "L,M,O,S,H,C"
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Time Difference less than 15 min.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Delay in opening table.IMD"
	task.AddExtraction dbName, "", "@SpanExcluding(ITEM_CODE,""L,M,O,S,H,C"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Excluding iRD and YOB outlets
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Delay in opening table.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Delay in opening table final.IMD"
	task.AddExtraction dbName, "", "@SpanExcluding(CHECK_NO,""I,Y"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Delay in opening table.IMD"
End Function

' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Time Difference less than 15 min.IMD"
End Function

' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Cashier ID in kot report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Cashier ID in kot report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Item code in kot report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase5
	Client.DeleteDatabase "Last settlement time in kot report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase6
	Client.DeleteDatabase "Check kot reconciliation report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase7
	Client.DeleteDatabase "Check Time Scroll Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase8
	Client.DeleteDatabase "Item Sale Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase9
	Client.DeleteDatabase "Settlement report copy.IMD"
End Function




' Filtering out Buffet from item description
' Data: Direct Extraction
Function DirectExtraction4
	Set db = Client.OpenDatabase("Delay in opening table final.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Buffet in Item.IMD"
	task.AddExtraction dbName, "", "@Isini(""buffet"",ITEM)"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Join delay in opening table with buffet 
' File: Join Databases
Function JoinDatabase3
	Set db = Client.OpenDatabase("Delay in opening table final.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Buffet in Item.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ITEM"
	task.AddMatchKey "ITEM", "ITEM", "A"
	dbName = "Delay in Opening Table.IMD"
	task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Removing Item1 field 
' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("Delay in Opening Table.IMD")
	Set task = db.TableManagement
	task.RemoveField "ITEM1"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' File: Delete Database
Function DeleteDatabase10
	Client.DeleteDatabase "Buffet in Item.IMD"
End Function

' File: Delete Database
Function DeleteDatabase11
	Client.DeleteDatabase "Delay in opening table final.IMD"
End Function


' Remove Field
Function RemoveField1
	Set db = Client.OpenDatabase("Delay in Opening Table.IMD")
	Set task = db.TableManagement
	task.RemoveField "TIME_IN_24_HR"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Bringing User name from user list in Delay in opening table by taking User iD as a unique value
' File: Join Databases
Function JoinDatabase4
	Set db = Client.OpenDatabase("Delay in Opening Table.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "F&B User list raw data-Sheet1.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "USER_NAME"
	task.AddMatchKey "CASHIER_ID", "USER", "A"
	dbName = "Delay in Opening Table Exception.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Delay in Opening Table Exception.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Delay in opening table final.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function



' File: Delete Database
Function DeleteDatabase12
	Client.DeleteDatabase "Settlement Time In Descending Order.IMD"
End Function

