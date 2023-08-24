Sub Main

'Check - Claims having zero sum insured'

Set db = Client.OpenDatabase("CLAIMSREGISTER_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Sum insured zero cases.IMD"
task.AddExtraction dbName, "", "SUM_INSURED_IN_LAKHS = 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Open claims in claims register but no provision amount'

Set db = Client.OpenDatabase("CLAIMSREGISTER_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Open claims.IMD"
task.AddExtraction dbName, "", "CLAIM_STATUS = ""OPEN"""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Open claims.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Open claims with zero or negative provision.IMD"
task.AddExtraction dbName, "", "PROVISION_AMT_IN_LAKHS  <= 0"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Open claims where provision amount is higher than the sum insured'

Set db = Client.OpenDatabase("Open claims.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Provision more than Sum insured.IMD"
task.AddExtraction dbName, "", "PROVISION_AMT_IN_LAKHS > SUM_INSURED_IN_LAKHS"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - RED before RID'

Set db = Client.OpenDatabase("CLAIMSREGISTER_202202-INPUTSHEET1.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "RED before RID.IMD"
task.AddExtraction dbName, "", "RISK_EXP_DATE < RISK_INC_DATE"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

'Check - Claim status check'
Set db = Client.OpenDatabase("CLAIMSPAYOUT_202202-INPUTSHEET1.IMD")
Set task = db.JoinDatabase
task.FileToJoin "CLAIMSREGISTER_202202-INPUTSHEET1.IMD"
task.IncludeAllPFields
task.AddSFieldToInc "CLAIM_NO"
task.AddSFieldToInc "CLAIM_STATUS"
task.AddMatchKey "CLAIM_NUMBER", "CLAIM_NO", "A"
dbName = "Claims status check.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)

Set db = Client.OpenDatabase("Claims status check.IMD")
Set task = db.Extraction
task.IncludeAllFields
dbName = "Claims status check with claims register.IMD"
task.AddExtraction dbName, "", "CLAIM_STATUS <> CLAIM_STATUS1"
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase (dbName)


End Sub
