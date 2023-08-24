Sub Main
                IgnoreWarning(True)
	Call DirectExtraction()	'Check Time Scroll Report raw data.IMD
	Call DirectExtraction1()	'Check Time Scroll Report copy.IMD
	Call ModifyField()	'Last print time with no blanks.IMD
	Call ModifyField1()	'Last print time with no blanks.IMD
	Call AppendField()	'Last print time with no blanks.IMD
	Call DirectExtraction2()	'Last print time with no blanks.IMD
	Call ModifyField2()	'Last print time with no blanks copy.IMD
	Call AppendField1()	'Last print time with no blanks copy.IMD
	Call DirectExtraction3()	'Last print time with no blanks copy.IMD
	Call ModifyField3()	'Settlement Report raw data.IMD
	Call SortDatabase()	'Settlement Report raw data.IMD
	Client.CloseDatabase "Settlement Report raw data.IMD"
	
                Call JoinDatabase()	'Delay in settlement more than 1 hr..IMD
	Call DirectExtraction4()	'Delay in settlement .IMD
	Client.CloseDatabase "Delay in settlement .IMD"
	Call DeleteDatabase()	'Delay in settlement .IMD
	Client.CloseDatabase "Delay in settlement more than 1 hr..IMD"
	Call DeleteDatabase1()	'Delay in settlement more than 1 hr..IMD
	Client.CloseDatabase "First print time with no blanks.IMD"
	Call DeleteDatabase2()	'Last print time with no blanks.IMD
	
	Client.CloseDatabase "Settlement Report raw data.IMD"
	Client.CloseDatabase "Check Time Scroll Report raw data.IMD"
	Client.CloseDatabase "Check Time Scroll Report copy.IMD"
	Client.CloseDatabase "Check Time Scroll Report copy.IMD"
	Call DeleteDatabase3()	'Check Time Scroll Report copy.IMD
	Call DeleteDatabase4()	'Last print time with no blanks.IMD
	Call DeleteDatabase5()	'First print time with no blanks.IMD
	Call DeleteDatabase6()	'Check Time Scroll Report with Print Time.IMD
	Call DirectExtraction5()	'Delay in Settlement.IMD
	Client.CloseDatabase "Delay in Settlement.IMD"
	Call DeleteDatabase7()	'Delay in Settlement.IMD
	Client.CloseDatabase "Settlement Time In Descending Order.IMD"
	Call DeleteDatabase8()	'Settlement Time In Descending Order.IMD
	Call JoinDatabase1()	'Delay In Settlement Cases.IMD
	Client.CloseDatabase "Delay In Settlement Exception.IMD"
	Call DeleteDatabase9()	'Delay In Settlement Exception.IMD

	Call ExportDatabaseXLSX()	'Delay In Settlement Cases.IMD
	
	

	IgnoreWarning(False)
End Sub



' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Check Time Scroll Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Check Time Scroll Report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Check Time Scroll Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "First print time with no blanks.IMD"
	task.AddExtraction dbName, "", "@NoMatch(FIRST_PRINT_TIME,"""")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("First print time with no blanks.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "FIRST_PRINT_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "FIRST_PRINT_TIME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("First print time with no blanks.IMD")
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

' Adding column for time in 24 hr.
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("First print time with no blanks.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_IN_24_HR"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = """24:00:00"""
	field.Length = 30
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Creating a copy for void check report 
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("First print time with no blanks.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "First print time with no blanks copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Converting time in 24 hr. field from character to time type
' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("First print time with no blanks copy.IMD")
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

' Difference of void time and last print time
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("First print time with no blanks copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_DIFFERENCE"
	field.Description = "Added field"
	field.Type = WI_VIRT_TIME
	field.Equation = "@If(LAST_SETTLEMENT_TIME  >= FIRST_PRINT_TIME,LAST_SETTLEMENT_TIME  - FIRST_PRINT_TIME,LAST_SETTLEMENT_TIME + TIME_IN_24_HR - FIRST_PRINT_TIME)"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function



' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("First print time with no blanks copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Delay in settlement more than 1 hr..IMD"
	task.AddExtraction dbName, "", "TIME_DIFFERENCE > ""2:00:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modify Field
Function ModifyField3
	Set db = Client.OpenDatabase("Settlement Report raw data.IMD")
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



' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Delay in settlement more than 1 hr..IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Settlement Time In Descending Order.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "CASHIER_ID"
	task.AddSFieldToInc "PAYMENT_MODE"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Delay in settlement .IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction4
	Set db = Client.OpenDatabase("Delay in settlement .IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Delay in Settlement.IMD"
	task.AddExtraction dbName, "", "@NoMatch(STATUS ,""VT"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Delay in settlement .IMD"
End Function

' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Delay in settlement more than 1 hr..IMD"
End Function

' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "First print time with no blanks copy.IMD"
End Function


' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Settlement Time In Descending Order.IMD"
End Function


' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Check Time Scroll Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase5
	Client.DeleteDatabase "Last print time with no blanks.IMD"
End Function

' File: Delete Database
Function DeleteDatabase6
	Client.DeleteDatabase "First print time with no blanks.IMD"
End Function

' File: Delete Database
Function DeleteDatabase7
	Client.DeleteDatabase "Check Time Scroll Report with Print Time.IMD"
End Function

' Data: Direct Extraction
Function DirectExtraction5
	Set db = Client.OpenDatabase("Delay in Settlement.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "CHECK_NO"
	task.AddFieldToInc "DATE"
	task.AddFieldToInc "OPEN_TIME"
	task.AddFieldToInc "FIRST_PRINT_TIME"
	task.AddFieldToInc "PRINT_DATE"
	task.AddFieldToInc "LAST_SETTLEMENT_TIME"
	task.AddFieldToInc "DISCOUNT"
	task.AddFieldToInc "VALUE"
	task.AddFieldToInc "STATUS"
	task.AddFieldToInc "OUTLET"
	task.AddFieldToInc "TIME_DIFFERENCE"
	task.AddFieldToInc "CASHIER_ID"
	task.AddFieldToInc "PAYMENT_MODE"
	dbName = "Delay In Settlement Exception.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File: Delete Database
Function DeleteDatabase8
	Client.DeleteDatabase "Delay in Settlement.IMD"
End Function


' Bringing User name from user list in Settlement report by taking User iD as a unique value
' File: Join Databases
Function JoinDatabase1
	Set db = Client.OpenDatabase("Delay In Settlement Exception.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "F&B User list raw data-Sheet1.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "USER_NAME"
	task.AddMatchKey "Cashier_ID", "USER", "A"
	dbName = "Delay In Settlement Cases.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File: Delete Database
Function DeleteDatabase9
	Client.DeleteDatabase "Delay In Settlement Exception.IMD"
End Function


' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Delay In Settlement Cases.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Delay In Settlement Cases.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function




