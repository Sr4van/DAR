'F&B-10-Delay in voiding kot
'PKF Sridhar & Santhanam LLP
'Date:- 4 November 2022

Sub Main
                IgnoreWarning(True)
	
	Call DirectExtraction()	'Void Item Report raw data.IMD
	Call AppendField()	'Void Item Report copy.IMD
	Call AppendField1()	'Void Item Report copy.IMD
	Call JoinDatabase()	'Void Item Report copy.IMD
	Call ModifyField()	'Kot punch time in void report.IMD
	Call ModifyField1()	'Kot punch time in void report.IMD
	Call AppendField2()	'Kot punch time in void report.IMD
	Call ModifyField2()	'Kot punch time in void report.IMD
	Call ModifyField3()	'Kot punch time in void report.IMD
	Call DirectExtraction1()	'Kot punch time in void report.IMD
	Call ModifyField4()	'Kot punch time in void report copy.IMD
	Call AppendField3()	'Kot punch time in void report copy.IMD
	Call DirectExtraction2()	'Kot punch time in void report copy.IMD
	Call RemoveField()	'Delay in voiding kot.IMD
	Client.CloseDatabase "Kot punch time in void report copy.IMD"
	Call DeleteDatabase()	'Kot punch time in void report copy.IMD
	Call RemoveField1()	'Delay in voiding kot.IMD
	Client.CloseDatabase "Kot punch time in void report.IMD"
	Call DeleteDatabase1()	'Kot punch time in void report.IMD
	Client.CloseDatabase "Void Item Report copy.IMD"
	Call DeleteDatabase2()	'Void Item Report copy.IMD
	Client.CloseDatabase "Check kot reconciliation report copy.IMD"
	Call DeleteDatabase3()	'Check kot reconciliation report copy.IMD
	Client.CloseDatabase "Check kot reconciliation report raw data.IMD"
	Client.CloseDatabase "Void Item Report raw data.IMD"
	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
End Sub


' Creating a copy of void item report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Void Item Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Void Item Report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Splitting check no from date column
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Void Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CHECK_NO"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(DATE,"""",1,"" "",1)"
	field.Length = 39
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Splitting date from date1 column
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("Void Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE1"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(DATE,"""",1,"" "")"
	field.Length = 8
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function


' Bringing kot punch time from check kot reconciliation report in void item report
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Void Item Report copy.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Check Kot Reconciliation Report raw data.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "TIME"
	task.AddMatchKey "CHECK_NO", "CHECK_NO", "A"
	task.AddMatchKey "ITEM", "ITEM", "A"
	dbName = "Kot punch time in void report.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Converting time field from character to time type
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Kot punch time in void report.IMD")
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


' Converting time field from character to time type
' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("Kot punch time in void report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME1"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "TIME1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Adding column for time in 24 hr
' Add Field
Function AppendField2
	Set db = Client.OpenDatabase("Kot punch time in void report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_IN_24_HR"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = """24:00:00"""
	field.Length = 59
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Converting kot punch time field from character to time type
' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("Kot punch time in void report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "KOT_PUNCH_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "TIME1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Converting void time field from character to time
' Modify Field
Function ModifyField3
	Set db = Client.OpenDatabase("Kot punch time in void report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VOID_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = "HH:MM"
	task.ReplaceField "TIME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Creating a copy of kot punch time in void report
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Kot punch time in void report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Kot punch time in void report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Converting Time in 24 hr from character to time type
' Modify Field
Function ModifyField4
	Set db = Client.OpenDatabase("Kot punch time in void report copy.IMD")
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

' Difference of void time and kot punch time
' Add Field
Function AppendField3
	Set db = Client.OpenDatabase("Kot punch time in void report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_DIFFERENCE"
	field.Description = "Added field"
	field.Type = WI_VIRT_TIME
	field.Equation = "@If(VOID_TIME  >= KOT_PUNCH_TIME,VOID_TIME - KOT_PUNCH_TIME,VOID_TIME + TIME_IN_24_HR - KOT_PUNCH_TIME)"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting delay in voiding kot cases of more than 30 Min.
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Kot punch time in void report copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Delay in voiding kot.IMD"
	task.AddExtraction dbName, "", "TIME_DIFFERENCE > ""00:30:00"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Removing unnecessary fields
' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("Delay in voiding kot.IMD")
	Set task = db.TableManagement
	task.RemoveField "TIME_IN_24_HR"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Deleting Unnecessary database
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Kot punch time in void report copy.IMD"
End Function

' Remove Field
Function RemoveField1
	Set db = Client.OpenDatabase("Delay in voiding kot.IMD")
	Set task = db.TableManagement
	task.RemoveField "DATE"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Kot punch time in void report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Void Item Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Check kot reconciliation report copy.IMD"
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Delay in voiding kot.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Delay in voiding kot.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
