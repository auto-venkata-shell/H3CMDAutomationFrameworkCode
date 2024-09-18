Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		set cardapi_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		'cardapi_data.add "custcardERP", customerERP_id
		
		cardapi_data.add "PinAdviceTypeidval", Get_Dictionary(ParamValDict,"PinAdviceTypeID" & "_" & iRowCount)
		cardapi_data.add "PinContactTypeIDval", Get_Dictionary(ParamValDict,"PinContactTypeID" & "_" & iRowCount)
		inputfilepathval = "PinReminderwithaddress.txt"
		outputfilepathval = "PinReminderwithaddressooutput.txt"
			
		'*********** Initialize Global Data for JSON Request *****************************************
		
		'XML_Cadrid = "159433"
		'XML_Cardpanid="16936649"
		
		reqType = "POST"
		apiurl = api_url & "/Card/Create_PinReminder"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CSSReg\PinReminder\" & inputfilepathval 
		'headers = "Content-Type==>application/json,charset==>utf-8,ClientID==>" & client_ID & ",ClientSecretKey==>" & client_SecretKey & ",Authorization==>Basic c2hlbGxfaDNfdGVzdDozZGFjOWQyNy1iZTkwLTRiYzgtYjQ2YS03N2I3MmJhODdiNDg=,RequestId==>1544def3-bdad-4d5e-ac38-33er43,UserCulture==>en-GB"
		newfile = sCurrentDirectory & "Test Data\CSSReg\PinReminderoutput\" & outputfilepathval
		newoutputxmlfile = sCurrentDirectory & "Test Data\CSSReg\PinReminder\UpdateXML.xml" 
		newcreationxmlfile = sCurrentDirectory & "Test Data\CSSReg\PinReminder\CreationInfoXML.xml" 
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
		
		pinreqcongdet = getpinReminderprerequestValues(cardapi_data)
		
		pinReminder_ATypeID = Split(pinreqcongdet,"|")(1)
		
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
			
		searchreplaceString = "ColCoID-" & pcolcoid &";CustomerID-"& pcustomerID &";CardID-" & pcardID &";PANID-"& ppanID &";CardPan-"& pcardpan & ";ExpiryDate-"& pcardexpirydate & ";PINAdviceTypeID-"& cardapi_data("PinAdviceTypeidval") & ";PINContactTypeID-"& cardapi_data("PinContactTypeIDval") 
		 
		
		
		Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
			
			Call validateAPIjob("2012",newoutputxmlfile,newcreationxmlfile)
			XML_Addressid = findValuefromString(newcreationxmlfile,"<AddressID>","</AddressID>")
			Call validatePostAPIpinReminderDetails()
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
		End If
	End If
next