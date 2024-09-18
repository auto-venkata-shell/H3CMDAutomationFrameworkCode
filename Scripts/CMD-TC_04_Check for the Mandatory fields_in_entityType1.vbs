Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		cust_data.add "custERP", customerERP_id
		cust_data.add "address", Get_Dictionary(ParamValDict,"Address" & "_" & iRowCount)
		cust_data.add "city", Get_Dictionary(ParamValDict,"City" & "_" & iRowCount)
		cust_data.add "region", Get_Dictionary(ParamValDict,"Region" & "_" & iRowCount)
		cust_data.add "zipCode", Get_Dictionary(ParamValDict,"Zip_Code" & "_" & iRowCount)
		cust_data.add "telephone", Get_Dictionary(ParamValDict,"Telephone" & "_" & iRowCount)
		inputfilepathval = "Toplevelhighvalueserrorcheck.txt"
		outputfilepathval = "Toplevelcustomeroutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		
		
		reqType = "POST"
		apiurl = api_url & "/Customer/MasterData"

		asynctype = false
		strType = "file"
		
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		'below statement commented
		Call basicPreConfigData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
		'********************** Scenario specific variables and business functions *********************
		
			
		searchreplaceString = "EntityTypeId-;ColcoCode-;UserName-;ExternalReference-;AccountNumber-"
		Append_TestHTML StepCounter, "Set Empty Values", searchreplaceString , "PASSED"
		Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		'*********** Generic Code for script ****************
			
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		'msgbox headers
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		rcheck = invokeErrorCMDAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Success:" & rcheck , "FAILED"
		End If
		
		inputfilepathval = "ToplevelNobodyErrCheck.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		searchreplaceString = "EntityTypeId-1;AccountNumber-" & newERPCustNo 
		Append_TestHTML StepCounter, "Set Empty BODY for " & newERPCustNo , searchreplaceString , "PASSED"
		Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		
		rcheck = invokeErrorCMDAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Success:" & rcheck , "FAILED"
		End If
		inputfilepathval = "Toplevel-BodyErrorscheck.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		searchreplaceString = "EntityTypeId-1;ColcoCode-" & err_colcoID & ";AccountNumber-" & newERPCustNo 
		Append_TestHTML StepCounter, "Set all Empty BODY vlaues for " & newERPCustNo , "All values of EntityBody is Empty" , "PASSED"
		Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		rcheck = invokeErrorCMDAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Success:" & rcheck , "FAILED"
		End If
	End If
next