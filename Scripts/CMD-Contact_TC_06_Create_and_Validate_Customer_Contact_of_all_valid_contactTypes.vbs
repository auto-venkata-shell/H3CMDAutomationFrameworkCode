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
		inputfilepathval = "Contact.txt"
		outputfilepathval = "Contactoutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		'topLevelERP ="NL01801002"
		
		reqType = "POST"
		apiurl = api_url & "/Customer/MasterData"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		externaRefVal = "contact01"&lcase(countryCode) & Right(topLevelERP,6)
		
		'*********** Generic Code for all scripts ****************
		
		
		
		'********************** Scenario specific variables and business functions *********************

	
		Call basicPreConfigContactData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
		'********************** Scenario specific variables and business functions *********************
		Append_TestHTML StepCounter, "Prepare INPUT JSON to add all contacttypeIDs", "Update and Insertion with all contacttypeids of"& topLevelERP & " and existing contact-"& externaRefVal, "PASSED"
		searchreplaceString = "EntityTypeId-7;IsPrimaryContact-true;IsActive-true;AccountNumber-" & topLevelERP & ";"& "ContactTypeId-[1,2,3,6,14,18,19];ExternalReference-" & externaRefVal & ";AddressId-" & topnewAddressref
		
		'Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
		Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		'msgbox headers
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
			Call validatejob2200()
			addresscustid = verifyPostCustomerAddressActiveDetails(topLevelERP,transReqID)
			Call validateContactLines(jsonFPath,addresscustid,externaRefVal,"1,2,3,6,14,18,19","insertionupdate")
			topnewContactref = externaRefVal
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