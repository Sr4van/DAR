Sub Main

Set task = Client.GetImportTask("ImportExcel")
dbName = "C:\Users\casewarebangalore\OneDrive - PKF Sridhar & Santhanam LLP\Documents\My IDEA Documents\IDEA Projects\GoDigit_Claims_FY202202\INPUTS\CLAIMSPAYOUT_202202.xlsx"
task.FileToImport = dbName
task.SheetToImport = "INPUTSHEET1"
task.OutputFilePrefix = "CLAIMSPAYOUT_202202"
task.FirstRowIsFieldName = "TRUE"
task.EmptyNumericFieldAsZero = "TRUE"
task.PerformTask
dbName = task.OutputFilePath("INPUTSHEET1")
Set task = Nothing
Client.OpenDatabase(dbName)

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "OFFICE_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "OFFICE_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "PRODUCT_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "PRODUCT_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "IMD_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "IMD_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "CLAIM_NUMBER"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 20
task.ReplaceField "CLAIM_NUMBER", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "LOSS_DATE"
field.Description = ""
field.Type = WI_DATE_FIELD
field.Equation = "YYYYMMDD"
task.ReplaceField "LOSS_DATE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "PAYABLE_LOCATION_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "PAYABLE_LOCATION_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "PRINT_LOCATION_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "PRINT_LOCATION_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "DEBIT_ACCOUNT_NO"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 30
task.ReplaceField "DEBIT_ACCOUNT_NO", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "DUMMY_COL1"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "DUMMY_COL1", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "DUMMY_COL2"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "DUMMY_COL2", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "EXPENSE_GL_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "EXPENSE_GL_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "PAYMENT_GL_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "PAYMENT_GL_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "BANK_GL_CODE"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 8
task.ReplaceField "BANK_GL_CODE", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

End Sub
