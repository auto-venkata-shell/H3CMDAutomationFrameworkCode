Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase("Yes") Then
		
		set err_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		err_data.add "envval", Get_Dictionary(ParamValDict,"Environment" & "_" & iRowCount)
		err_data.add "userName", Get_Dictionary(ParamValDict,"Username" & "_" & iRowCount)
		err_data.add "page", Get_Dictionary(ParamValDict,"pageno" & "_" & iRowCount)
		err_data.add "pageSize", Get_Dictionary(ParamValDict,"pagesizeval" & "_" & iRowCount)

		inputfilepathval = "ErrorAPImand.txt"
		outputfilepathval = "ErrorAPIoutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		'Below flag for checking default price profile and fee values
		defaultcheckflag = True
		
		reqType = "POST"
		apiurl = api_url & "/Customer/MasterDataError"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
		
		
		
		'********************** Scenario specific variables and business functions *********************
		'Call getErrorapiprerequisits(err_data,jsonFPath)
		 
		'searchreplaceString = "Page-1;PageSize-1000"
		'Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		
		rcheck = invokeErrorCMDAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Success:" & rcheck , "FAILED"
		Else
			Call errorcodevalidation(rcheck, "E0001", "Validation Error" )
		End If
		
		
	End If
next