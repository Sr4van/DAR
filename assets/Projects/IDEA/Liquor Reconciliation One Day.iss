'F&B-14-Liquor Reconciliation
'PKF Sridhar & Santhanam LLP
'Date:- 8 November 2022

Sub Main
                IgnoreWarning(True)
	Call DirectExtraction()	'Item Sale Report copy.IMD
	Call AppendField()	'Liquor Sold.IMD
	Call ModifyField()	'Liquor Sold.IMD
	Call Summarization()	'Liquor Sold.IMD
	Call ExcelImport()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\POS items with recipe codes.xlsx
	Call AppendField1()	'POS items with recipe codes raw data.IMD
	Call AppendField2()	'POS items with recipe codes raw data.IMD
	Call ExcelImport1()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Recipe card.xlsx
	Call AppendField3()	'Recipe card copy.IMD
	Call JoinDatabase()	'Recipe card copy.IMD
	Call JoinDatabase1()	'Pos Item Code in Recipe Card.IMD
	Call AppendField4()	'Sum of Qty sold in Recipe Card.IMD
	Call ExcelImport2()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Consumption report.xlsx
	Call DirectExtraction1()'Consumption Report Copy,IMD

	Call AppendField5()	'Liquor store in consumption report.IMD
	Call Summarization1()	'Liquor store in consumption report.IMD
	Call ExcelImport3()	'C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Liquor Master.xlsx
	Call AppendField6()	'Liquor Master copy.IMD
	Call AppendField7()	'Sum of qty sold in Recipe card copy.IMD
	Call Summarization2()	'Sum of qty sold in Recipe card copy.IMD
	Call JoinDatabase2()	'Liquor Master copy.IMD
	Call JoinDatabase3()	'Product of Qty sold in Liquor Master.IMD
	Call ModifyField1()	'Qty issued sum in Liquor Master.IMD
	Call AppendField8()'Qty issued sum in Liquor Master.IMD
	Call AppendField9()'Qty issued sum in Liquor Master.IMD
	Call DirectExtraction2()	'Qty issued sum in Liquor Master.IMD
	Call RemoveField()	'Liquor Reconciliation.IMD
	Client.CloseDatabase "Qty issued sum in consumption report.IMD"
	Call DeleteDatabase()	'Qty issued sum in consumption report.IMD
	Client.CloseDatabase "Liquor store in consumption report.IMD"
	Call DeleteDatabase1()	'Liquor store in consumption report.IMD
	Client.CloseDatabase "Consumption report copy.IMD"
	Call DeleteDatabase2()	'Consumption report copy.IMD
	Client.CloseDatabase "Sum of Qty sold in item sale report.IMD"
	Call DeleteDatabase3()	'Sum of Qty sold in item sale report.IMD
	Client.CloseDatabase "Liquor Sold.IMD"
	Call DeleteDatabase4()	'Liquor Sold.IMD
	Client.CloseDatabase "Item Sale Report copy.IMD"
	Call DeleteDatabase5()	'Item Sale Report copy.IMD
	Client.CloseDatabase "Qty issued sum in Liquor Master.IMD"
	Call DeleteDatabase6()	'Qty issued sum in Liquor Master.IMD
	Client.CloseDatabase "Product of Qty sold in Liquor Master.IMD"
	Call DeleteDatabase7()	'Product of Qty sold in Liquor Master.IMD
	Client.CloseDatabase "POS items with recipe code copy.IMD"
	Call DeleteDatabase8()	'POS items with recipe code copy.IMD
	Client.CloseDatabase "Product of qty sold sum copy.IMD"
	Call DeleteDatabase9()	'Product of qty sold sum copy.IMD
	Client.CloseDatabase "Product of qty sold sum.IMD"
	Call DeleteDatabase10()	'Product of qty sold sum.IMD
	Client.CloseDatabase "Sum of qty sold in Recipe card copy.IMD"
	Call DeleteDatabase11()	'Sum of Qty sold in Recipe Card.IMD
	Client.CloseDatabase "Pos Item Code in Recipe Card.IMD"
	Call DeleteDatabase12()	'Pos Item Code in Recipe Card.IMD
	Client.CloseDatabase "Recipe card copy.IMD"
	Call DeleteDatabase13()	'Recipe card copy.IMD
	Client.CloseDatabase "Liquor Master copy.IMD"
	Call DeleteDatabase14()	'Liquor Master copy.IMD
	Client.CloseDatabase "Liquor Master raw data-Sheet1.IMD"
	Client.CloseDatabase "Consumption report raw data-detail_s1.IMD"
	Client.CloseDatabase "Recipe card raw data-detail_s1.IMD"
	Client.CloseDatabase "POS items with recipe codes raw data-detail_s1.IMD"
	Client.CloseDatabase "Item Sale Report raw data.IMD"
	Call RemoveField1()	'Liquor Reconciliation.IMD
	Call ModifyField2()	'Liquor Reconciliation.IMD
	Call ModifyField3()	'Liquor Reconciliation.IMD
	Call RemoveField2()	'Liquor Reconciliation.IMD
	Call ExportDatabaseXLSX()

                IgnoreWarning(False)
End Sub


' Extracting Item code with H & C from item sale report to liquor item sold only
' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Item Sale Report raw data.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Liquor Sold.IMD"
	task.AddExtraction dbName, "", "@SpanIncluding(ITEM_CODE,""C,H"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Trim item code in extracted item sale report
' Add Field
Function AppendField
	Set db = Client.OpenDatabase("Liquor Sold.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_ITEM_CODE"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(ITEM_CODE)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modifying quantity field type from character to number
' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("Liquor Sold.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "QUANTITY"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "QUANTITY", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Summarise item sale report to get the total quantity sold per item code
' Analysis: Summarization
Function Summarization
	Set db = Client.OpenDatabase("Liquor Sold.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ITEM_CODE"
	task.AddFieldToInc "ITEM_DESCRIPTION"
	task.AddFieldToInc "WEIGH"
	task.AddFieldToInc "NET_SALE"
	task.AddFieldToInc "DISCOUNT"
	task.AddFieldToInc "TAX"
	task.AddFieldToInc "TYPE"
	task.AddFieldToInc "OUTLET"
	task.AddFieldToInc "DATE"
	task.AddFieldToInc "TRIMMED_ITEM_CODE"
	task.AddFieldToTotal "QUANTITY"
	dbName = "Sum of Qty sold in item sale report.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Import POS item with recipe code report
' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\POS items with recipe codes.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "detail_s1"
	task.OutputFilePrefix = "POS items with recipe codes raw data"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("detail_s1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function

' Trim POS item code
' Add Field
Function AppendField1
	Set db = Client.OpenDatabase("POS items with recipe codes raw data-detail_s1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_POS_ITEM_CODE"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(POS_ITEM_CODE)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Trim recipe id in POS item with recipe code report
' Add Field
Function AppendField2
	Set db = Client.OpenDatabase("POS items with recipe codes raw data-detail_s1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_RECIPE_ID"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(RECIPE_ID)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Import Recipe card report
' File - Import Assistant: Excel
Function ExcelImport1
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Recipe card.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "detail_s1"
	task.OutputFilePrefix = "Recipe card raw data"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("detail_s1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function

' Trim recipe id in recipe card report
' Add Field
Function AppendField3
	Set db = Client.OpenDatabase("Recipe card raw data-detail_s1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_RECIPE_ID"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(RECIPE_ID)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Bringing POS item code in recipe card from POS item with recipe card report
' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Recipe card raw data-detail_s1.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "POS items with recipe codes raw data-detail_s1.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "TRIMMED_POS_ITEM_CODE"
	task.AddMatchKey "TRIMMED_RECIPE_ID", "TRIMMED_RECIPE_ID", "A"
	dbName = "Pos Item Code in Recipe Card.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bringing sum of quantity sold from item sale report in recipe card report
' File: Join Databases
Function JoinDatabase1
	Set db = Client.OpenDatabase("Pos Item Code in Recipe Card.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Sum of Qty sold in item sale report.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "QUANTITY_SUM"
	task.AddMatchKey "TRIMMED_POS_ITEM_CODE", "TRIMMED_ITEM_CODE", "A"
	dbName = "Sum of Qty sold in Recipe Card.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Add field to get the product of quantity sold and quantity in ML in recipe card report
' Add Field
Function AppendField4
	Set db = Client.OpenDatabase("Sum of Qty sold in Recipe Card.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PRODUCT_QTY_SUM_WITH_QTY"
	field.Description = "Added field"
	field.Type = WI_VIRT_NUM
	field.Equation = "QUANTITY_SUM * QUANTITY"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function



' Import Consumption Report
' File - Import Assistant: Excel
Function ExcelImport2
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Consumption.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "detail_s1"
	task.OutputFilePrefix = "Consumption report raw data"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("detail_s1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function


' Extracting Liquor Store type from Consumption report
' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("Consumption report raw data-detail_s1.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Liquor store in consumption report.IMD"
	task.AddExtraction dbName, "", "@SpanIncluding(STORE_TYPE_DESC,""LIQUOR STORE"")"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Trim Item description in consumption report
' Add Field
Function AppendField5
	Set db = Client.OpenDatabase("Liquor store in consumption report.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_ITEM_DESC"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(ITEM_DESC)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Summarise to get the sum of quantity consumed in consumption report
' Analysis: Summarization
Function Summarization1
	Set db = Client.OpenDatabase("Liquor store in consumption report.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "TRIMMED_ITEM_DESC"
	task.AddFieldToInc "STOCK_UNIT"
	task.AddFieldToTotal "QTY_ISSUED"
	dbName = "Qty issued sum in consumption report.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Import Liquor master
' File - Import Assistant: Excel
Function ExcelImport3
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Liquor Master.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "Liquor Master raw data"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function


' Trim item description in liquor master
' Add Field
Function AppendField6
	Set db = Client.OpenDatabase("Liquor Master raw data-Sheet1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_ITEM"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(ITEM)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Trim Item description in recipe card report
' Add Field
Function AppendField7
	Set db = Client.OpenDatabase("Sum of qty sold in Recipe card.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRIMMED_ITEM_DESC"
	field.Description = "Added field"
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Trim(ITEM_DESC)"
	field.Length = 200
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Summarise product of Qty sold with qty in ML to get the total quantity sold in ML
' Analysis: Summarization
Function Summarization2
	Set db = Client.OpenDatabase("Sum of qty sold in Recipe card.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "TRIMMED_ITEM_DESC"
	task.AddFieldToTotal "PRODUCT_QTY_SUM_WITH_QTY"
	dbName = "Product of qty sold sum.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

'' Bringing product of quantity sold sum in liquor master
' File: Join Databases
Function JoinDatabase2
	Set db = Client.OpenDatabase("Liquor Master raw data-Sheet1.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Product of qty sold sum.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "PRODUCT_QTY_SUM_WITH_QTY_SUM"
	task.AddMatchKey "TRIMMED_ITEM", "TRIMMED_ITEM_DESC", "A"
	dbName = "Product of Qty sold in Liquor Master.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bringing quantity issued sum in liquor master from consumption report
' File: Join Databases
Function JoinDatabase3
	Set db = Client.OpenDatabase("Product of Qty sold in Liquor Master.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Qty issued sum in consumption report.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "QTY_ISSUED_SUM"
	task.AddMatchKey "TRIMMED_ITEM", "TRIMMED_ITEM_DESC", "A"
	dbName = "Qty issued sum in Liquor Master.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("Qty issued sum in Liquor Master.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "QTY_ISSUED_SUM"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "QTY_ISSUED_SUM", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Divide Product of qty sold sum with Units
' Add Field
Function AppendField8
	Set db = Client.OpenDatabase("Qty issued sum in Liquor Master.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PRODUCT_OF_QTY_SOLD_UPON_UNIT"
	field.Description = "Added field"
	field.Type = WI_VIRT_NUM
	field.Equation = "PRODUCT_QTY_SUM_WITH_QTY_SUM / UNITS"
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Difference between product of qty sold sum and Qty issued 
' Add Field
Function AppendField9
	Set db = Client.OpenDatabase("Qty issued sum in Liquor Master.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DIFFERENCE"
	field.Description = "Added field"
	field.Type = WI_VIRT_NUM
	field.Equation = "PRODUCT_OF_QTY_SOLD_UPON_UNIT - QTY_ISSUED_SUM"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Extracting difference not equal to "0"
' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("Qty issued sum in Liquor Master.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Liquor Reconciliation.IMD"
	task.AddExtraction dbName, "", "DIFFERENCE <> 0"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Removing unnecessary fields
' Remove Field
Function RemoveField
	Set db = Client.OpenDatabase("Liquor Reconciliation.IMD")
	Set task = db.TableManagement
	task.RemoveField "TRIMMED_ITEM"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Deleting Unnecessary database
' File: Delete Database
Function DeleteDatabase
	Client.DeleteDatabase "Qty issued sum in consumption report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase1
	Client.DeleteDatabase "Liquor store in consumption report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase2
	Client.DeleteDatabase "Consumption report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase3
	Client.DeleteDatabase "Sum of Qty sold in item sale report.IMD"
End Function

' File: Delete Database
Function DeleteDatabase4
	Client.DeleteDatabase "Liquor Sold.IMD"
End Function

' File: Delete Database
Function DeleteDatabase5
	Client.DeleteDatabase "Item Sale Report copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase6
	Client.DeleteDatabase "Qty issued sum in Liquor Master.IMD"
End Function

' File: Delete Database
Function DeleteDatabase7
	Client.DeleteDatabase "Product of Qty sold in Liquor Master.IMD"
End Function

' File: Delete Database
Function DeleteDatabase8
	Client.DeleteDatabase "POS items with recipe code copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase9
	Client.DeleteDatabase "Product of qty sold sum copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase10
	Client.DeleteDatabase "Product of qty sold sum.IMD"
End Function


' File: Delete Database
Function DeleteDatabase11
	Client.DeleteDatabase "Sum of Qty sold in Recipe Card.IMD"
End Function

' File: Delete Database
Function DeleteDatabase12
	Client.DeleteDatabase "Pos Item Code in Recipe Card.IMD"
End Function

' File: Delete Database
Function DeleteDatabase13
	Client.DeleteDatabase "Recipe card copy.IMD"
End Function

' File: Delete Database
Function DeleteDatabase14
	Client.DeleteDatabase "Liquor Master copy.IMD"
End Function


' Removing unnecessary field
' Remove Field
Function RemoveField1
	Set db = Client.OpenDatabase("Liquor Reconciliation.IMD")
	Set task = db.TableManagement
	task.RemoveField "UNITS"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Modifying field name from Qty issued sum to consumption
' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("Liquor Reconciliation.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CONSUMPTION"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "QTY_ISSUED_SUM", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modifying field name from Product of qty sold upon unit to sale
' Modify Field
Function ModifyField3
	Set db = Client.OpenDatabase("Liquor Reconciliation.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "SALE"
	field.Description = "Added field"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "PRODUCT_OF_QTY_SOLD_UPON_UNIT", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Removing unnecessary field
' Remove Field
Function RemoveField2
	Set db = Client.OpenDatabase("Liquor Reconciliation.IMD")
	Set task = db.TableManagement
	task.RemoveField "PRODUCT_QTY_SUM_WITH_QTY_SUM"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Liquor Reconciliation.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\Admin\Desktop\CMS- IDEA\Foods and Beverages\Exceptions\Liquor Reconciliation.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function

