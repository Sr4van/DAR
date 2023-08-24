'F&B-13-Sharing of User ID
'PKF Sridhar & Santhanam LLP
'Date:- 5 November 2022
Sub Main
	IgnoreWarning(True)
	
	Call DirectExtraction()	'F&B User list raw data-Sheet1.IMD
	Call DirectExtraction1()	'Settlement Report raw data.IMD
	Call AppendField()	'F&B User list copy.IMD
	Call AppendField1()	'Settlement Report copy.IMD
	Call JoinDatabase()	'Settlement Report copy.IMD
	
	Call DirectExtraction2()	'Attendance raw data-Sheet1.IMD
	Call JoinDatabase1()	'User Name in settlement report.IMD
	Call ModifyField()	'Attendance in settlement report.IMD
	Call DirectExtraction3()	'Attendance in settlement report.IMD
	Call RemoveField()	'F&B sharing of user id.IMD
	Client.CloseDatabase "Attendance in settlement report.IMD"
	Call DeleteDatabase()	'Attendance in settlement report.IMD
	Client.CloseDatabase "User Name in settlement report.IMD"
	Call DeleteDatabase1()	'User Name in settlement report.IMD
	Client.CloseDatabase "Settlement Report copy.IMD"
	Call DeleteDatabase2()	'Settlement Report copy.IMD
	Client.CloseDatabase "F&B User list copy.IMD"
	Call DeleteDatabase3()	'F&B User list copy.IMD
	Client.CloseDatabase "Attendance copy.IMD"
	Call DeleteDatabase4()	'Attendance copy.IMD
	Client.CloseDatabase "Attendance raw data-Sheet1.IMD"
	Client.CloseDatabase "Settlement Report raw data.IMD"
	Client.CloseDatabase "F&B User list raw data-Sheet1.IMD"
	Client.CloseDatabase "Attendance-Sheet1.IMD"

	Call ExportDatabaseXLSX()

	IgnoreWarning(False)
End Sub



' Creating a copy of F&B user list
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("F&B User list raw data-Sheet1.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "F&B User list copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Creating a copy of settlement report
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Settlement Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Settlement Report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Adding field for trim ID in User list report
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("F&B User list copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_ID"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(USER)"
	field.Length = 246
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Adding field for trim cashier id in settlement report
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("Settlement Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_ID"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(CASHIER_ID)"
	field.Length = 400
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Bringing User name from user list in Settlement report by taking User iD as a unique value
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Settlement Report copy.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "F&B User list copy.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "USER_NAME"
	task.AddMatchKey "TRIMMED_ID", "TRIMMED_ID", "A"
	dbName = "User Name in settlement report.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' Creating a copy of attendance report
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Attendance-Sheet1.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Attendance copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bringing attendance from attendance report in settlement report by taking User name as a unique value
' File: Join Databases
Function JoinDatabase1
	Set db = Client.OpenDatabase("User Name in settlement report.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Attendance copy.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "DATE"
	task.AddMatchKey "USER_NAME", "EMPLOYEE_NAME", "A"
	dbName = "Attendance in settlement report.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Renaming date field to attendance 
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Attendance in settlement report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ATTENDANCE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "DATE1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting cases where the user was absent on a particular date
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Attendance in settlement report.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "F&B sharing of user id.IMD"
	task.AddExtraction dbName, "", "@Match(ATTENDANCE,""A"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Removing unnecessary fields
' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("F&B sharing of user id.IMD")
	Set task = db.TableManagement
	task.RemoveField "TRIMMED_ID"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Deleting Unnecessary database
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Attendance in settlement report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "User Name in settlement report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Settlement Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "F&B User list copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Attendance copy.IMD"
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("F&B sharing of user id.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\F&B sharing of user id.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
