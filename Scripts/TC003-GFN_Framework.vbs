Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'cERP = Get_Dictionary(ParamValDict,"custerERP" & "_" & iRowCount)
	cERP = customerERP_id
sName = "AutomationTest Shayam " 

updateCSName = "Ram"
oltpdbName = "GFN_SHELL_SPRINTQA_NL_OLTP"
batchdbName = "GFN_SHELL_SPRINTQA_NL_BATCH"
wwwdbName = "GFN_SHELL_SPRINTQA_NL_WWW"
reportsdbName = "GFN_SHELL_SPRINTQA_NL_REPORTS"


	If strExecute = "Yes_"& countryName Then
		'Call  DB_check((ShortName,cERP,"GFN_SHELL_SPRINTQA_NL_OLTP")
	newshortName = customer_dbValidation(cERP,oltpdbName,True)
		'msgbox newshortName
		batchsName = customer_dbValidation(cERP,batchdbName,False)
		if newshortName = batchsName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in BATCh DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & batchsName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in BATCh DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & batchsName, "FAILED"		
		End If
		wwwsName = customer_dbValidation(cERP,wwwdbName,False)
		if newshortName = wwwsName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in WWW DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & wwwsName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in WWW DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & wwwsName, "FAILED"		
		End If
		reportssName = customer_dbValidation(cERP,reportsdbName,False)
		if newshortName = reportssName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in REPORTS DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & reportssName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in REPORTS DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & reportssName, "FAILED"		
		End If
		updateCSName = originalcustSName
		newshortName = customer_dbValidation(cERP,oltpdbName,True)
	End If
next