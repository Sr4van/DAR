'F&B-5-Liquor sold on dry day
'PKF Sridhar & Santhanam LLP
'Date:- 1 November 2022

Sub Main
                IgnoreWarning (True)
	Call DirectExtraction()	'Item Sale Report raw data .IMD

	Call ExcelImport()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Dry Days.xlsx

		Call ModifyField()	'Item Sale report copy.IMD
	Call AppendField()	'Item Sale report copy.IMD
	Call AppendField1()	'Item Sale report copy.IMD
	Call DirectExtraction1()	'Item Sale report copy.IMD
	Call ModifyField1()	'Copy of Item sale working.IMD
	Call ModifyField2()	'Copy of Item sale working.IMD
	Call AppendField2()	'Copy of Item sale working.IMD
	Call DirectExtraction2()	'Copy of Item sale working.IMD
	Call DirectExtraction3()	'Copy of Item sale working.IMD
	Call AppendDatabase()	'C Codes.IMD
	Call ModifyField3()	'Sale of liquor.IMD
	Call JoinDatabase()	'Sale of liquor.IMD
	Client.CloseDatabase "C Codes.IMD"
	Call DeleteDatabase()	'C Codes.IMD
	Client.CloseDatabase "H Codes.IMD"
	Call DeleteDatabase1()	'H Codes.IMD
	Client.CloseDatabase "Sale of liquor.IMD"
	Call DeleteDatabase2()	'Sale of liquor.IMD
	Client.CloseDatabase "Item Sale report copy.IMD"
	Call DeleteDatabase3()	'Item Sale report copy.IMD
	Client.CloseDatabase "Dry Days copy.IMD"
	Call DeleteDatabase4()	'Dry Days copy.IMD
	Client.CloseDatabase "Copy of Item sale working.IMD"
	Call DeleteDatabase5()	'Copy of Item sale working.IMD
	Client.CloseDatabase "Item Sale Report raw data.IMD"
	Call DirectExtraction4()	'Liquor sold on dry day.IMD
	Client.CloseDatabase "Liquor sold on dry day.IMD"
	Call DeleteDatabase6()	'Liquor sold on dry day.IMD
	Call RemoveField()	'Liquor sold on dry day.IMD
	Client.CloseDatabase "Dry Days-Sheet1.IMD"


	Call ExportDatabaseXLSX()

                IgnoreWarning (False)
End Sub



' Creating a copy of item sale report
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Item Sale Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Item Sale report copy.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Dry Days.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "Dry Days"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function


' Modifying Date feild from character to date
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Item Sale report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "DD-MM-YYYY"
	task.ReplaceField "DATE", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Add field to get day from date 
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Item Sale report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DAY"
	field.Description = "Added field"
	field.Type = WI_VIRT_NUM
	field.Equation = "@Day(DATE)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Add field to get month from date
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("Item Sale report copy.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "MONTH"
	field.Description = "Added field"
	field.Type = WI_VIRT_NUM
	field.Equation = "@Month(date)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Creating a copy of Item sale report copy
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Item Sale report copy.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Copy of Item sale working.IMD"
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modifying type of day from date to character
' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("Copy of Item sale working.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DAY"
	field.Description = "Added field"
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 8
	task.ReplaceField "DAY", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modifying type of month from date to character
' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("Copy of Item sale working.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "MONTH"
	field.Description = "Added field"
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 8
	task.ReplaceField "MONTH", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Add field to concatenate day and month
' Add Field
Function AppendField2
	Set db = Client.OpenDatabase("Copy of Item sale working.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CONCAT"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "DAY+MONTH"
	field.Length = 20
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting Item code = "H" from Item code 
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Copy of Item sale working.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "H Codes.IMD"
	task.AddExtraction dbName, "", "@SpanIncluding(ITEM_CODE,""H"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Extracting Item Code ="C" from Item Code 
' Data: Direct Extraction
Function DirectExtraction3
	Set db = Client.OpenDatabase("Copy of Item sale working.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "C Codes.IMD"
	task.AddExtraction dbName, "", "@SpanIncluding(ITEM_CODE,""C"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Append H Code with C Code 
' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("H Codes.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "C Codes.IMD"
	dbName = "Sale of liquor.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modifying type of Concatenated column from character to number 
' Modify Field
Function ModifyField3
	Set db = Client.OpenDatabase("Sale of liquor.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CONCAT"
	field.Description = "Added field"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "CONCAT", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Join Concat Dry day with Liquor sold on dry days report
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Sale of liquor.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Dry Days-Sheet1.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "DRY_DAY"
	task.AddMatchKey "CONCAT", "DRY_DAY", "A"
	dbName = "Liquor sold on dry day.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting C Codes
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "C Codes.IMD"
End Function

' Deleting H Codes
' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "H Codes.IMD"
End Function

' Deleting Sale of liquor 
' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Sale of liquor.IMD"
End Function

' Deleting Item sale report copy
' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Item Sale report copy.IMD"
End Function

' Deleting dry days copy
' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Dry Days copy.IMD"
End Function

' Deleting copy of item sale working
' File: Delete Database
Function DeleteDatabase5
	Client.DeleteDatabase "Copy of Item sale working.IMD"
End Function



' Extracting Dry day not equal to "0" from liquor sold on dry days report
' Data: Direct Extraction
Function DirectExtraction4
	Set db = Client.OpenDatabase("Liquor sold on dry day.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Liquor sold on dry days.IMD"
	task.AddExtraction dbName, "", "DRY_DAY <> 0.00"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Deleting liquor sold on dry days
' File: Delete Database
Function DeleteDatabase6
	Client.DeleteDatabase "Liquor sold on dry days.IMD"

End Function

' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("Liquor sold on dry day.IMD")
	Set task = db.TableManagement
	task.RemoveField "CONCAT"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function


' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Liquor sold on dry day.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Liquor sold on dry day.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function
