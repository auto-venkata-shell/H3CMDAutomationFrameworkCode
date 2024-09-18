Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	cERP1 = Get_Dictionary(ParamValDict,"UserERP1" & "_" & iRowCount)
	cERP2 = Get_Dictionary(ParamValDict,"UserERP2" & "_" & iRowCount)
	cERP3 = Get_Dictionary(ParamValDict,"UserERP3" & "_" & iRowCount)
	

importdbName="GFN_SHELL_SPRINTQA_IMPORT"
masterdbName="GFN_SHELL_SPRINTQA_MASTER"
batchdbName = "GFN_SHELL_SPRINTQA_NL_BATCH"
If strExecute = "Yes" Then
'If strExecute = "Yes_"& countryName Then
	compval = "DX026_GFN_TRX_(?<tranProviderID>[0-9]{5})_(?<SequenceNo>[0-9]{6})_(?<DatePart>[0-9]{8})_(?<TimePart>[0-9]{6})(_(?<BusinessDate>[0-9]{8}))?(.dat)"
	Call VerifyRGX_Value(cERP1,importdbName,compval)
	compval = "(?i)DX053_([0-9]{3})_([0-9]{6}_([0-9]{8})_([0-9]{6}))(.json)"
	Call VerifyRGX_Value(cERP2,masterdbName,compval)
	compval = "(?i)DX350_([0-9]{8})_([0-9]{3,4}_([0-9]{6})_([0-9]{8})_([0-9]{6}))(.dat)"
	Call VerifyRGX_Value(cERP3,batchdbName,compval)
	'masterdbName1 = Job_VerifyJobID(cERP2,masterdbName,False)
	'	if correctid = masterdbName1 Then
	'		Append_TestHTML StepCounter,  "Verify the job in Batch DB"," Job "&cERP2&" in Batch DB is in running  state with StatusID='"&masterdbName1&"'", "PASSED"   
	'
	'	Else
	'		Append_TestHTML StepCounter,  "Verify the job in Batch DB"," Job '"&cERP2&" in Batch DB is not in running  state with StatusID='"&masterdbName1&"'", "FAILED"
	'	End If
'	importid1 = Job_VerifyJobID(cERP3,importdbName,False)
'		if correctid = importid1 Then
'			Append_TestHTML StepCounter,  "Verify the job in Import DB"," Job '"&cERP3&" in Import DB is in running  state with StatusID='"&importid1&"'", "PASSED"   
'	
'		Else
'			Append_TestHTML StepCounter,  "Verify the job in Import DB"," Job '"&cERP3&" in Import DB is not in running  state with StatusID='"&importid1&"'", "FAILED"
'		End If
		
End If
next