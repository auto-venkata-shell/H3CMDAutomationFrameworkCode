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

	Call VerifyRGX_Value(cERP1,importdbName)
	Call VerifyRGX_Value1(cERP2,masterdbName)
	Call VerifyRGX_Value2(cERP3,batchdbName)
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