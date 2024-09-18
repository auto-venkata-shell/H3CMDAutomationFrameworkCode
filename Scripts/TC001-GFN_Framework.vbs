Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	cERP1 = Get_Dictionary(ParamValDict,"UserERP1" & "_" & iRowCount)
	cERP2 = Get_Dictionary(ParamValDict,"UserERP2" & "_" & iRowCount)
	cERP3 = Get_Dictionary(ParamValDict,"UserERP3" & "_" & iRowCount)
	cERP4 = Get_Dictionary(ParamValDict,"UserERP4" & "_" & iRowCount)
	cERP5 = Get_Dictionary(ParamValDict,"UserERP5" & "_" & iRowCount)
	cERP6 = Get_Dictionary(ParamValDict,"UserERP6" & "_" & iRowCount)

batchdbName = "GFN_SHELL_SPRINTQA_NL_BATCH"
importdbName="GFN_SHELL_SPRINTQA_IMPORT"
masterdbName="GFN_SHELL_SPRINTQA_MASTER"
oltpdbName = "GFN_SHELL_SPRINTQA_NL_OLTP"



	If strExecute = "Yes_"& countryName Then		
	correctid= Job_VerifyJobID(cERP1,batchdbName,True)
	batchid1 = Job_VerifyJobID(cERP2,batchdbName,False)
		if correctid = batchid1 Then
			Append_TestHTML StepCounter,  "Verify the job in Batch DB"," Job "&cERP2&" in Batch DB is in running  state with StatusID='"&batchid1&"'", "PASSED"   
	
		Else
			Append_TestHTML StepCounter,  "Verify the job in Batch DB"," Job '"&cERP2&" in Batch DB is not in running  state with StatusID='"&batchid1&"'", "FAILED"
		End If
	importid1 = Job_VerifyJobID(cERP3,importdbName,False)
		if correctid = importid1 Then
			Append_TestHTML StepCounter,  "Verify the job in Import DB"," Job '"&cERP3&" in Import DB is in running  state with StatusID='"&importid1&"'", "PASSED"   
	
		Else
			Append_TestHTML StepCounter,  "Verify the job in Import DB"," Job '"&cERP3&" in Import DB is not in running  state with StatusID='"&importid1&"'", "FAILED"
		End If
	masterid1 = Job_VerifyJobID(cERP4,masterdbName,False)
		if correctid = masterid1 Then
			Append_TestHTML StepCounter,  "Verify the job in Master DB"," Job '"&cERP4&" in Master DB is in running  state with StatusID='"&masterid1&"'", "PASSED"   	
		Else
			Append_TestHTML StepCounter,  "Verify the job in Master DB"," Job '"&cERP4&" in Master DB is not in running  state with StatusID='"&masterid1&"'", "FAILED"
		End If
	batchid2 = Job_VerifyJobID(cERP5,batchdbName,False)
		if correctid = batchid2 Then
			Append_TestHTML StepCounter,  "Verify the job in Batch DB"," Job '"&cERP5&" in Batch DB is in running  state with StatusID='"&batchid2&"'", "PASSED"   	
		Else
			Append_TestHTML StepCounter,  "Verify the job in Batch DB"," Job '"&cERP5&" in Batch DB is not in running  state with StatusID='"&batchid2&"'", "FAILED"
		End If	
	Oltpid1 = Job_VerifyJobID(cERP6,oltpdbName,False)
		if correctid = Oltpid1 Then
			Append_TestHTML StepCounter,  "Verify the job in Oltp DB"," Job '"&cERP6&" in Oltp DB is in running  state with StatusID='"& Oltpid1&"'", "PASSED"   
	
		Else
			Append_TestHTML StepCounter,  "Verify the job in Oltp DB"," Job '"&cERP6&" in Oltp DB is not in running  state with StatusID='"& Oltpid1&"'", "FAILED"
		End If
	
	
	
	'call  Job_VerifyJobID(cERP2,batchdbName,True)
	'call  Job_VerifyJobID(cERP3,importdbName,True)
	'call  Job_VerifyJobID(cERP4,masterdbName,True)
	'call  Job_VerifyJobID(cERP5,batchdbName,True)
	'call  Job_VerifyJobID(cERP6,oltpdbName,True)
		
	End If
next