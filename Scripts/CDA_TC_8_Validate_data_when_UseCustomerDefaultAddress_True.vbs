Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		set cardapi_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		cardapi_data.add "custcardERP", customerERP_id
		
		'cardapi_data.add "PinAdviceTypeidval", Get_Dictionary(ParamValDict,"PinAdviceTypeID" & "_" & iRowCount)
		'cardapi_data.add "PinContactTypeIDval", Get_Dictionary(ParamValDict,"PinContactTypeID" & "_" & iRowCount)
		inputfilepathval = "CardAddressjsonTrue.txt"
		outputfilepathval = "CardAddressjsonoutput.txt"
			
		'*********** Initialize Global Data for JSON Request *****************************************
		
		'XML_Cadrid = "154629"
		'XML_Cardpanid="16931845"
		
		reqType = "POST"
		apiurl = api_url & "/Card/CardDeliveryAddress_Update"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CSSReg\CardAddressChange\" & inputfilepathval 
		'headers = "Content-Type==>application/json,charset==>utf-8,ClientID==>" & client_ID & ",ClientSecretKey==>" & client_SecretKey & ",Authorization==>Basic c2hlbGxfaDNfdGVzdDozZGFjOWQyNy1iZTkwLTRiYzgtYjQ2YS03N2I3MmJhODdiNDg=,RequestId==>1544def3-bdad-4d5e-ac38-33er43,UserCulture==>en-GB"
		newfile = sCurrentDirectory & "Test Data\CSSReg\PinReminderoutput\" & outputfilepathval
		newoutputxmlfile = sCurrentDirectory & "Test Data\CSSReg\CardAddressChange\UpdateXML.xml" 
		newcreationxmlfile = sCurrentDirectory & "Test Data\CSSReg\CardAddressChange\CreationInfoXML.xml" 
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
		
		
		Call validateprePINReminderValues()
		
		
		panCarddet = pinReminderPrerequestData()
		pcustomerID = Split(panCarddet,"|")(0)
		pcardID = Split(panCarddet,"|")(1)
		ppanID = Split(panCarddet,"|")(2)
		pcardpan = Split(panCarddet,"|")(3)
		pcardexpirydate = Split(panCarddet,"|")(4)
		pcolcoid = Split(panCarddet,"|")(5)
		'pcardexpirydate = null
		'********************** Scenario specific variables and business functions *********************

		'newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		'msgbox newERPCustNo
		
		'********************** Scenario specific variables and business functions *********************
			
		searchreplaceString = "ColCoID-" & pcolcoid &";CustomerID-"& pcustomerID &";CardID-" & pcardID 
		 
		
		
		Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck <> Empty Then
			
			Call validateAPIjob("2013",newoutputxmlfile,newcreationxmlfile)
			Call validatepostCADValues(10,false,13,true,false,"") 'First two CDStatu, Next two CPStatus, Next two New records
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
		Else
		Append_TestHTML StepCounter, "Verify Response JSON ", "Error messages displayed" & rcheck , "PASSED"
		End If
	End If
next