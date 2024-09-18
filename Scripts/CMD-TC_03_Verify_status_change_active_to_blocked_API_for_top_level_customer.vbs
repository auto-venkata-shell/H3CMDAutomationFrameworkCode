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
		'Below flag for checking default price profile and fee values
		defaultcheckflag = False
		inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		
		'topLevelERP = "PH01891003"
		If topLevelERP = Empty Then
			Append_TestHTML StepCounter, "Create Sublevel Customer via API", "Error: Fail to create Top level Customer", "FAILED"
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
			
			
			
			'********************** Scenario specific variables and business functions *********************

			newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
			'msgbox newERPCustNo
			Call basicPreConfigData(jsonFPath,"GFN_SHELL_SPRINTQA_PH_OLTP")
			'********************** Scenario specific variables and business functions *********************
			searchreplaceString = "EntityTypeId-2;AccountNumber-" & newERPCustNo & ";"& "CustomerERP-" & newERPCustNo & ";"& "TopLevelCustomerNumber-" & topLevelERP & ";"& "ParentCustomerNumber-" & topLevelERP & ";"& "RegistrationNumber-Reg" & newERPCustNo & ";"& "VATRegNumber-Reg" & newERPCustNo
			
			Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
			headers = getapiConfigInfoFromCSV(Apconfigpath)
			'msgbox headers
				'Call verifyPreCustomerActiveDetails(newERPCustNo)
				cstatusval =  getCustomerStatus(newERPCustNo)
			rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
			
			If rcheck = Empty Then
				Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
			Else
				Call validatejob2200()
				Call verifyPostCustomerActiveDetails(newERPCustNo)
				Set fileSysObj = createObject("Scripting.FileSystemObject")
				fileSysObj.CreateTextFile(newfile)
				Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
				DXwrite.Write rcheck
				DXwrite.Close
				Set fileSysObj = Nothing
			End If
		
		End IF
	End If
next