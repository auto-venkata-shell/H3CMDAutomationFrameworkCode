Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		'*********** Generic Code for all scripts ****************
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		cust_data.add "custERP", customerERP_id
		cust_data.add "address", Get_Dictionary(ParamValDict,"Address" & "_" & iRowCount)
		cust_data.add "city", Get_Dictionary(ParamValDict,"City" & "_" & iRowCount)
		cust_data.add "region", Get_Dictionary(ParamValDict,"Region" & "_" & iRowCount)
		cust_data.add "zipCode", Get_Dictionary(ParamValDict,"Zip_Code" & "_" & iRowCount)
		cust_data.add "telephone", Get_Dictionary(ParamValDict,"Telephone" & "_" & iRowCount)
		inputfilepathval = "CreditInsurance.txt"
		outputfilepathval = "CreditInsuranceoutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		'topLevelERP ="NL01801002"
		'topLevelERP = customerERP_id
		reqType = "POST"
		apiurl = api_url & "/Customer/MasterData"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		externaRefVal = "CreditInsu01"&lcase(countryCode) & Right(topLevelERP,6)
		
		
		'********************** Scenario specific variables and business functions *********************
		
		'*******Pre Data comparision checks************
		Call validatePreCreditInsurenceRecords(topLevelERP)
		
		'myDateEffective = nextdateDBFormt(Date())
		
		searchreplaceString = "EntityTypeId-14;AccountNumber-" & topLevelERP & ";ExternalReference-" & externaRefVal &";CreditLimit-200000"
		
		'Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
			
		Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		'B - Medium    ,  A - Low Risk   ,   C - High Risk
		mypredISData = basicPreCreditInsurenceAPIData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP","A - Low Risk",0)
		
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		'msgbox headers
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
				Call validatejob2200()
			'Call validatejob2200withstatus(2)
			addresscustid = verifyPostAPIActiveDetails("CreditRatingId",transReqID)
			Call validatePostCreditInsurenceRecords(jsonFPath,addresscustid)
			
			topnewCreditInsref = externaRefVal
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
			'verifyAddressTypecheckFromUI(True)
		End If
	End If
next