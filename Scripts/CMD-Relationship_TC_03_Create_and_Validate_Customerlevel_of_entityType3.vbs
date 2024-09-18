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
		
		inputfilepathval = "Relationship.txt"
		outputfilepathval = "Relationshipoutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		'Below flag for checking default price profile and fee values
		defaultcheckflag = True
		
		'topLevelERP = "PH01891003"
		If topLevelERP = Empty Then
			Append_TestHTML StepCounter, "Create Sublevel Customer via API", "Error: Fail to create Top level Customer", "FAILED"
			bRunFlag = False
		Else
		
			reqType = "POST"
			apiurl = api_url & "/Customer/MasterData"

			asynctype = false
			strType = "file"
			strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
			'headers = "Content-Type==>application/json,charset==>utf-8,ClientID==>" & client_ID & ",ClientSecretKey==>" & client_SecretKey & ",Authorization==>Basic c2hlbGxfaDNfdGVzdDozZGFjOWQyNy1iZTkwLTRiYzgtYjQ2YS03N2I3MmJhODdiNDg=,RequestId==>1544def3-bdad-4d5e-ac38-33er43,UserCulture==>en-GB"
			newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
			jsonFPath = strJson
			
			Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
			
			'*********** Generic Code for all scripts ****************
			externaRefVal = "Relationship01"&lcase(countryCode) & Right(subLevelERP,6)
			
			query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			Set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
			wait 3
			db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
			db_CompanyID = dictDbResultSet("CompanyID")
			Set dictDbResultSet = Nothing
			Set query = Nothing
			
			
			'********************** Scenario specific variables and business functions *********************
		
			'newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
			'msgbox newERPCustNo
			'Call basicPreConfigData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
			
			
			'********************** Scenario specific variables and business functions *********************
			searchreplaceString = "ColcoCode-" & db_ColcoCode &";EntityTypeId-3;AccountNumber-" & subLevelERP & ";"& "ExternalReference-" & externaRefVal & ";"& "TopLevelCustomerNumber-" & topLevelERP & ";"& "ParentCustomerNumber-" & topLevelERP & ";"& "PaymentCustomerNumber-" & topLevelERP & ";"& "InvoiceCustomerNumber-" & topLevelERP & ";"& "CardDeliveryCustomerNumber-" & topLevelERP & ";"& "PricingCustomerNumber-" & topLevelERP & ";"& "ContractCustomerNumber-" & topLevelERP
			
			Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
			headers = getapiConfigInfoFromCSV(Apconfigpath)
			'msgbox headers
				'Call verifyPreCustomerActiveDetails(newERPCustNo)
			rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
			
			If rcheck = Empty Then
				Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
			Else
				Call validatejob2200()
				Call verifyPostCustomerActiveDetails(subLevelERP)
				RelationshipLevelERP = subLevelERP
				Call validateCustomerRelationshipData(jsonFPath,subLevelERP)
				Call navigateToRelationshipManagementScreen(subLevelERP)
				Set fileSysObj = createObject("Scripting.FileSystemObject")
				fileSysObj.CreateTextFile(newfile)
				Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
				DXwrite.Write rcheck
				DXwrite.Close
				Set fileSysObj = Nothing
				if cint(db_ecustStatusID) = cint("1") Then
				Append_TestHTML StepCounter, "Verify Customer Status ", "Expected-1" & VBCRLF & "Actual-"& db_ecustStatusID , "PASSED"
			Else
				Append_TestHTML StepCounter, "Verify Customer Status ", "Expected-1" & VBCRLF & "Actual-"& db_ecustStatusID , "FAILED"
			End IF
			End If
		
		End IF
	End If
next