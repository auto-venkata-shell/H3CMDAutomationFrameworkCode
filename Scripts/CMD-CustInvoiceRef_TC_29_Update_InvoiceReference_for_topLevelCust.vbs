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
		inputfilepathval = "InvoiceReference.txt"
		outputfilepathval = "InvoiceReferenceoutput.txt"
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
		
		externaRefVal = "InvoiceRef03"&lcase(countryCode) & Right(topLevelERP,6)
		
		'*********** Generic Code for all scripts ****************
		
	
		'********************** Scenario specific variables and business functions *********************
		query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
		Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
		Set dictDbResultSet = execute_db_query(query, 1, dbName)
		wait 3
		
		db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
		
		Set dictDbResultSet = Nothing
		Set query = Nothing

		myDateEffective = nextdateDBFormt(Date())
		myDateTer_Year = cint(Split(myDateEffective,"-")(0)) + 1
		myDateTerminated =  myDateTer_Year & "-" & Split(myDateEffective,"-")(1) & "-" & Split(myDateEffective,"-")(2)
		'myDateTerminated ="9999-12-31"
		searchreplaceString = "EntityTypeId$8;AccountNumber$" & topLevelERP & ";ExternalReference$" & externaRefVal & ";ColcoCode$" & db_ColcoCode  &";InvoiceReference$UPDATEFIRSTINVOICEREFTEXT;DateEffective$" & myDateEffective & ";DateTerminated$"& myDateTerminated
		
		'Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
		Call searchandReplaceMultipleStringwithdollar(jsonFPath, searchreplaceString)	
		'Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		'mypredISData = basicPreConfigInvoiceReferenceData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
		
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		'msgbox headers
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
				Call validatejob2200()
			'Call validatejob2200withstatus(2)
			addresscustid = verifyPostAPIActiveDetails("InvoiceReference",transReqID)
			Call validateCustomerInvoiceRefRecords(jsonFPath,addresscustid,externaRefVal)
			
			topinvoiceCRef1 = externaRefVal
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