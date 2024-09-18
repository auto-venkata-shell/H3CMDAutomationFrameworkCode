Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	cERP = Get_Dictionary(ParamValDict,"UserERP" & "_" & iRowCount)
'sName = "AutomationTest Shayam " 

updateSSName = "NewTestSite"
masterdbName="GFN_SHELL_SPRINTQA_MASTER"
oltpdbName = "GFN_SHELL_SPRINTQA_NL_OLTP"
batchdbName = "GFN_SHELL_SPRINTQA_NL_BATCH"
wwwdbName = "GFN_SHELL_SPRINTQA_NL_WWW"
reportsdbName = "GFN_SHELL_SPRINTQA_NL_REPORTS"


	'If strExecute = "Yes_"& countryName Then
	If strExecute = "Yes" Then
	cERP = getPreItemForReplication("Select * from Site;",masterdbName,"SiteID")
		'Call  DB_check((ShortName,cERP,"GFN_SHELL_SPRINTQA_NL_OLTP")
	newshortName = site_dbValidation(cERP,masterdbName,True)
		'msgbox newshortName
		
		oltpsName = site_dbValidation(cERP,oltpdbName,False)
		if newshortName = oltpsName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in oltp DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & oltpsName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in oltp DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & oltpsName, "FAILED"		
		End If
		
		
		batchsName = site_dbValidation(cERP,batchdbName,False)
		if newshortName = batchsName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in BATCh DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & batchsName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in BATCh DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & batchsName, "FAILED"		
		End If
		wwwsName = site_dbValidation(cERP,wwwdbName,False)
		if newshortName = wwwsName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in WWW DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & wwwsName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in WWW DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & wwwsName, "FAILED"		
		End If
		reportssName =site_dbValidation(cERP,reportsdbName,False)
		if newshortName = reportssName Then
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in REPORTS DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & reportssName, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Short Name of Customer "& cERP & " in REPORTS DB"  ,"Expected Value: " & newshortName & VBCRLF & "Actual Value: " & reportssName, "FAILED"		
		End If
		updateSSName = originalsiteSName
		newshortName = site_dbValidation(cERP,masterdbName,True)
	End If
next