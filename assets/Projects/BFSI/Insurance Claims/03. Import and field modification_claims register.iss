Sub Main

Set task = Client.GetImportTask("ImportExcel")
dbName = "C:\Users\casewarebangalore\OneDrive - PKF Sridhar & Santhanam LLP\Documents\My IDEA Documents\IDEA Projects\GoDigit_Claims_FY202202\INPUTS\CLAIMSREGISTER_202202.xlsx"
task.FileToImport = dbName
task.SheetToImport = "INPUTSHEET1"
task.OutputFilePrefix = "CLAIMSREGISTER_202202"
task.FirstRowIsFieldName = "TRUE"
task.EmptyNumericFieldAsZero = "TRUE"
task.PerformTask
dbName = task.OutputFilePath("INPUTSHEET1")
Set task = Nothing
Client.OpenDatabase(dbName)

Set db = Client.OpenDatabase("CLAIMSREGISTER_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "CLAIM_NO"
field.Description = ""
field.Type = WI_CHAR_FIELD
field.Equation = ""
field.Length = 20
task.ReplaceField "CLAIM_NO", field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSREGISTER_202202-INPUTSHEET1.IMD")
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

Set db = Client.OpenDatabase("CLAIMSREGISTER_202202-INPUTSHEET1.IMD")
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

End Sub
