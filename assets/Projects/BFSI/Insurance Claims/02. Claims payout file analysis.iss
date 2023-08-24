Sub Main

'Check: GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS > POLICY_SUM_INSURED_IN_LAKHS'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Claim amount greater than sum insured.IMD"
task.AddExtraction dbName, "", "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS > POLICY_SUM_INSURED_IN_LAKHS"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check: Gross amount is 0 and TDS is deducted '
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claim amount equal to zero.IMD"
task.AddExtraction dbName, "", "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS =0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Gross claim amount equal to zero.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "TDS Not equal to Zero.IMD"
task.AddExtraction dbName, "", "TDS_AMOUNT_IN_LAKHS <> 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("TDS Not equal to Zero.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross amount is 0 and TDS is deducted.IMD"
task.AddExtraction dbName, "", "STATUS <> ""Deleted"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check: Gross claim amount is 0 or negative but GST is positive'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claims less than or equal to 0.IMD"
task.AddExtraction dbName, "", "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS <= 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Gross claims less than or equal to 0.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "GST Amount greater than zero.IMD"
task.AddExtraction dbName, "", "GST_AMOUNT_IN_LAKHS > 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("GST Amount greater than zero.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claim amount is 0 or negative but GST is positive.IMD"
task.AddExtraction dbName, "", "STATUS <> ""Deleted"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check: Length of PAN card should be 10 digits for claims more than INR. 100,000'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claim amount greater than 1Lakh.IMD"
task.AddExtraction dbName, "", "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS >1"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Gross claim amount greater than 1Lakh.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Length of PAN not equal to 10 digits for claims more than INR. 1L.IMD"
task.AddExtraction dbName, "", "@Len(RECEIVER_PAN) <> 10"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check Risk Expiry before Risk Inception'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Risk expiry  before risk inception.IMD"
task.AddExtraction dbName, "", "RISK_EXP_DATE  < RISK_INC_DATE"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check Claim booking date before Risk Inception Date'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Claim booking date Before risk inception date.IMD"
task.AddExtraction dbName, "", "BOOKING_DATE<RISK_INC_DATE"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check Gross amount is positive and TDS is also positive'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claim amount greater than zero.IMD"
task.AddExtraction dbName, "", "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS > 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Gross claim amount greater than zero.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross amount is positive and TDS is also positive.IMD"
task.AddExtraction dbName, "", "TDS_AMOUNT_IN_LAKHS > 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check Gross amount is negative and TDS is also negative'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claims less than or equal to 0.IMD"
task.AddExtraction dbName, "", "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS <= 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Gross claims less than or equal to 0.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross amount is negative and TDS is also negative.IMD"
task.AddExtraction dbName, "", "TDS_AMOUNT_IN_LAKHS < 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Net claims cross validation'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "NET_CLAIM_CHECK"
field.Description = "Added field"
field.Type = WI_NUM_FIELD
field.Equation = " GROSS_AMOUNT_AS_PER_FINANCE_LOGIC_IN_LAK + GST_AMOUNT_IN_LAKHS +TDS_AMOUNT_IN_LAKHS + DEDUCTION_AMOUNT_IN_LAKHS + ADVANCE_AMOUNT_IN_LAKHS"
field.Decimals = 6
task.AppendField field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "DIFFERENCE_IN_NET_CLAIMS"
field.Description = "Added field"
field.Type = WI_NUM_FIELD
field.Equation = "NET_AMOUNT_IN_LAKHS - NET_CLAIM_CHECK"
field.Decimals = 3
task.AppendField field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Exceptions in net claims check.IMD"
task.AddExtraction dbName, "", "DIFFERENCE_IN_NET_CLAIMS <> 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Exceptions in net claims check.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Paid claims exceptions in net claims check.IMD"
task.AddExtraction dbName, "", "STATUS = ""Paid"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check : Bank code/IFSC less than 11 digits'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "IFSC less than 11 digits.IMD"
task.AddExtraction dbName, "", "@Len(BANK_CODE) < 11"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("IFSC less than 11 digits.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "IFSC - Payment type.IMD"
task.AddExtraction dbName, "", "PAYMENT_TYPE = ""Bank transfer"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("IFSC - Payment type.IMD")
Set task = db.KeyValueExtraction
Dim myArray(0,0)
myArray(0,0) = "Paid"
task.IncludeAllFields
task.AddKey "STATUS", "A"
task.DBPrefix = "Exceptions - Paid claims where IFSC is less than 11 digits"
task.CreateMultipleDatabases = TRUE
task.ValuesToExtract myArray
task.PerformTask
dbName = task.DBName
Set task = Nothing
Set db = Nothing
Client.OpenDatabase(dbName)

'Check Recovered claims having positive net amount'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Recovered claims.IMD"
task.AddExtraction dbName, "", "STATUS = ""Recovery"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Recovered claims.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Recovered claims having positive net amount.IMD"
task.AddExtraction dbName, "", "NET_AMOUNT_IN_LAKHS > 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Duplicate paid claims'

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "CONCATENATED_FIELD"
field.Description = "Added field"
field.Type = WI_CHAR_FIELD
field.Equation = "CLAIM_NUMBER + "" "" + INVOICE_NUMBER + "" "" + BENEFICIARY_NAME"
field.Length = 999
task.AppendField field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Paid claims.IMD"
task.AddExtraction dbName, "", "STATUS = ""Paid"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Paid claims.IMD")
Set task = db.DupKeyDetection
task.IncludeAllFields
task.AddKey "CONCATENATED_FIELD", "A"
task.OutputDuplicates = TRUE
dbName = "Duplicate paid claims.IMD"
task.PerformTask dbName, ""
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check: BOOKING_DATE > PAYMENT_DATE_FORMATTED'

Set db = Client.OpenDatabase("Paid claims.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Claim booking date after payment date.IMD"
task.AddExtraction dbName, "", "BOOKING_DATE > PAYMENT_DATE_FORMATTED"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check: BOOKING_DATE < LOSS_DATE_FORMATTED'

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Claim booking date after loss date.IMD"
task.AddExtraction dbName, "", "BOOKING_DATE < LOSS_DATE_FORMATTED"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - LOSS_DATE_FORMATTED > RISK_EXP_DATE'

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Loss date after risk expiry date.IMD"
task.AddExtraction dbName, "", "LOSS_DATE_FORMATTED > RISK_EXP_DATE"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - INVOICE_DATE_FORMATTED > PAYMENT_DATE_FORMATTED'

Set db = Client.OpenDatabase("Paid claims.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Invoice date later than payment date.IMD"
task.AddExtraction dbName, "", "INVOICE_DATE_FORMATTED > PAYMENT_DATE_FORMATTED"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Gross claims - ABS vs Finance logic'

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "GROSS_CLAIMS_ABS_VS_FINANCE"
field.Description = "Added field"
field.Type = WI_NUM_FIELD
field.Equation = "GROSS_AMOUNT_AS_PER_ABS_IN_LAKHS - GROSS_AMOUNT_AS_PER_FINANCE_LOGIC_IN_LAK"
field.Decimals = 6
task.AppendField field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Gross claims - ABS vs Finance logic.IMD"
task.AddExtraction dbName, "", "GROSS_CLAIMS_ABS_VS_FINANCE <> 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Claims booked after 1 year of loss date'

Set db = Client.OpenDatabase("Paid claims.IMD")
Set task = db.TableManagement
Set field = db.TableDef.NewField
field.Name = "DIFF_BET_BOOKING_AND_LOSS_DATE"
field.Description = "Added field"
field.Type = WI_VIRT_NUM
field.Equation = "@Age(BOOKING_DATE,LOSS_DATE_FORMATTED)"
field.Decimals = 0
task.AppendField field
task.PerformTask
Set task = Nothing
Set db = Nothing
Set field = Nothing

Set db = Client.OpenDatabase("Paid claims.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Claims booked after 1 year of loss date.IMD"
task.AddExtraction dbName, "", "DIFF_BET_BOOKING_AND_LOSS_DATE > 365"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Deleted claims pivot'

Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Deleted claims.IMD"
task.AddExtraction dbName, "", "STATUS = ""Deleted"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

End Sub

Sub Main
	Call PivotTable()	'Deleted claims.IMD
End Sub

' Analysis: Pivot Table
Function PivotTable
	Set db = Client.OpenDatabase("Deleted claims.IMD")
	Set task = db.PivotTable
	task.ResultName = "Pivot Table"
	task.AddRowField "CLAIM_NUMBER"
	task.AddDataField "GROSS_AMOUNT_AS_PER_FINANCE_LOGIC", "Sum: GROSS_AMOUNT_AS_PER_FINANCE_LOGIC", 1
	task.ExportToIDEA False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function



