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
			inputfilepathval = "InfoSubscriptionDetail.txt"
		outputfilepathval = "InfoSubscriptionDetailoutput.txt"
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
		
		externaRefVal = "InfoSubscDet01"&lcase(countryCode) & Right(topLevelERP,6)
		
		'*********** Generic Code for all scripts ****************
		
	
		'********************** Scenario specific variables and business functions *********************'CI_InfoSubscriptionID
		Call validateCustomerandCustomerInfoSubscriptionEntry(topLevelERP,1)
		searchreplaceString = "EntityTypeId-10;AccountNumber-" & topLevelERP & ";ExternalReference-" & externaRefVal &";InfoSubscriptionId-" & topnewInfoSubref1 & ";ContactId-" & topnewContactref& ";IsPrimary-true;SuppressContact-false;IsActive-false"
		
		'Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
			
		Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		mypredISData = basicPreConfigUpdateInfoSubscriptionDetailData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
		'mypredISData = basicPreConfigInfoSubscriptionDetailData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
		
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		'msgbox headers
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
				Call validatejob2200()
			'Call validatejob2200withstatus(2)
			addresscustid = verifyPostAPIActiveDetails("DistributionMethodId",transReqID)
			Call validateInfoSubscriptionDetailRecords(jsonFPath,addresscustid,externaRefVal,"1","update")
			Call validatemultipleISDIsPrimaryStatus(topnewInfoSubDetref3,externaRefVal,false,true)
			topnewInfoSubDetref2 = externaRefVal
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
			'verifyAddressTypecheckFromUI(True)
		End If
		newIS_DMID_val=""
		newIS_IOT_val = ""
		newIS_ADM_val = ""
		newIS_SC_val = ""
		newIS_IP_val = ""
		newIS_IA_val = ""
	End If
next