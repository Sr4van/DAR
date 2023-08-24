'F&B-11-Delay in voiding bills
'PKF Sridhar & Santhanam LLP
'Date:- 5 November 2022

Sub Main
                IgnoreWarning(True)
	Call DirectExtraction()	'Check Time Scroll Report raw data.IMD

	Call JoinDatabase()	'Void Check Report copy.IMD
	Call ModifyField()	'Print time in void check report.IMD
	Call ModifyField1()	'Print time in void check report.IMD
	Call AppendField()	'Print time in void check report.IMD
	Call DirectExtraction1()	'Print time in void check report.IMD
	Call ModifyField2()	'Print time in void check report copy.IMD
	Call AppendField1()	'Print time in void check report copy.IMD
	Call DirectExtraction2()	'Print time in void check report copy.IMD
	Client.CloseDatabase "Print time in void check report copy.IMD"
	Call DeleteDatabase()	'Print time in void check report copy.IMD
	Client.CloseDatabase "Print time in void check report.IMD"
	Call DeleteDatabase1()	'Print time in void check report.IMD
	Client.CloseDatabase "Check Time Scroll Report raw data.IMD"
	Client.CloseDatabase "Void Check Report raw data.IMD"
	Client.CloseDatabase "Check Time Scroll Report with Print Time.IMD"
	Call RemoveField()	'Delay in voiding bills.IMD

	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
End Sub

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Check Time Scroll Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Check Time Scroll Report with Print Time.IMD"
	task.AddExtraction dbName, "", "@NoMatch( LAST_PRINT_TIME,"""")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function




' Bringing last print time in void check report 
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Void Check Report raw data.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Check Time Scroll Report with Print Time.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "LAST_PRINT_TIME"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	dbName = "Print time in void check report.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Converting void time field from character to time type
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Print time in void check report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VOID_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "VOID_TIME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Converting last print time field from character to time type
' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("Print time in void check report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "LAST_PRINT_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "LAST_PRINT_TIME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Adding column for time in 24 hr.
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Print time in void check report.IMD")
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
Function DirectExtraction1
	Set db = Client.OpenDatabase("Print time in void check report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Print time in void check report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Converting time in 24 hr. field from character to time type
' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("Print time in void check report copy.IMD")
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
	Set db = Client.OpenDatabase("Print time in void check report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_DIFFERENCE"
	field.Description = "Added field"
	field.Type = WI_VIRT_TIME
	field.Equation = "@If(VOID_TIME >= LAST_PRINT_TIME,VOID_TIME - LAST_PRINT_TIME,VOID_TIME + TIME_IN_24_HR - LAST_PRINT_TIME)"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting delay in voiding cases for more than 30 min.
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Print time in void check report copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Delay in voiding bills.IMD"
	task.AddExtraction dbName, "", "TIME_DIFFERENCE > ""00:30:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting unnecessary database
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Print time in void check report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Print time in void check report.IMD"
End Function



' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("Delay in voiding bills.IMD")
	Set task = db.TableManagement
	task.RemoveField "TIME_IN_24_HR"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function
' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Delay in voiding bills.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Delay in voiding bills.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
