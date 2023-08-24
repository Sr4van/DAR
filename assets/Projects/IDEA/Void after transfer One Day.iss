'F&B-12-Void after transfer
'PKF Sridhar & Santhanam LLP
'Date:- 5 November 2022
Sub Main
	IgnoreWarning(True)
	Call DirectExtraction()	'Transferred Item Report raw data.IMD
	Call AppendField()	'Transferred Item Report copy.IMD
	Call AppendField1()	'Transferred Item Report copy.IMD
	Call AppendField2()	'Transferred Item Report copy.IMD
	Call AppendField3()	'Transferred Item Report copy.IMD
	Call AppendField4()	'Transferred Item Report copy.IMD
	Call ModifyField()	'Transferred Item Report copy.IMD
	Call DirectExtraction1()	'Void Check Report raw data.IMD
	Call DirectExtraction2()
	Call RemoveField()	'Transferred Item report second copy.IMD
	Call ModifyField1()	'Void Check Report copy.IMD
	Call ModifyField2()	'Void Check Report copy.IMD
	Call ModifyField3()	'Void Check Report copy.IMD
	Call JoinDatabase()	'Transferred Item report second copy.IMD
	Client.CloseDatabase "Transferred Item report second copy.IMD"
	Client.CloseDatabase "Transferred Item Report copy.IMD"
	Call DeleteDatabase1()	'Transferred Item Report copy.IMD
	Client.CloseDatabase "Void Check Report copy.IMD"
	Call DeleteDatabase2()	'Void Check Report copy.IMD
	Client.CloseDatabase "Void Check Report raw data.IMD"
	Client.CloseDatabase "Transferred Item Report raw data.IMD"
	Call ExportDatabaseXLSX()	'Void after transfer.IMD

	IgnoreWarning(False)
End Sub


' Creating a copy of transferred item report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Transferred Item Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Transferred Item Report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Splitting qty1 in seperate column from qty column
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "QTY1"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(QTY,"""",1,"" "")"
	field.Length = 30
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Splitting amount in seperate column from qty column
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "AMOUNT"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(QTY,"" "",9,"" "",1)"
	field.Length = 30
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Splitting Check no in seperate column from qty column
' Add Field
Function AppendField2
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CHECK_NO1"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(QTY,"" "",8,"" "",1)"
	field.Length = 30
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Splitting date from qty column
' Add Field
Function AppendField3
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(QTY,"" "",4,"" "",1)"
	field.Length = 30
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Splitting time from qty column
' Add Field
Function AppendField4
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(QTY,"""",1,"" "",1)"
	field.Length = 30
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modify Check No1 name with Move to Check name 
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "MOVE_TO_CHECK"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SimpleSplit(QTY,"" "",8,"" "",1)"
	field.Length = 30
	task.ReplaceField "CHECK_NO1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Creating a copy of void check report
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Void Check Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Void Check Report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' Creating a second copy of transferred item report
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Transferred Item Report copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Transferred Item Report Second Copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function



' Removing field naming "Qty"
' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("Transferred Item Report Second Copy.IMD")
	Set task = db.TableManagement
	task.RemoveField "QTY"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Modify field name from Cashier ID to voided by
' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("Void Check Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VOIDED_BY"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 12
	task.ReplaceField "CASHIER_ID", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modify field name from amount to void amount
' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("Void Check Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VOID_AMOUNT"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "AMOUNT", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modify field name from reason to void reason
' Modify Field
Function ModifyField3
	Set db = Client.OpenDatabase("Void Check Report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VOID_REASON"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 23
	task.ReplaceField "REASON", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Bringing "Voided by", "Void time". "Void reason" and "Void amount" in transferred item report  by taking item description and Check No.as unique value
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Transferred Item Report Second Copy.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Void Check Report copy.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "VOID_AMOUNT"
	task.AddSFieldToInc "VOIDED_BY"
	task.AddSFieldToInc "VOID_TIME"
	task.AddSFieldToInc "VOID_REASON"
	task.AddMatchKey "MOVE_TO_CHECK", "CHECK_NO", "A"
	task.AddMatchKey "ITEM", "ITEM", "A"
	dbName = "Void after transfer.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting unnecessary files
'' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Transferred Item Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Void Check Report copy.IMD"
End Function



' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Void after transfer.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Void after transfer.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
