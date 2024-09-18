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
		inputfilepathval = "PinReminderwithaddressfalse.txt"
		outputfilepathval = "PinReminderwithaddressooutput.txt"
			
		'*********** Initialize Global Data for JSON Request *****************************************
		
		'XML_Cadrid = "154629"
		'XML_Cardpanid="16931845"
		
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
		
		'*********** Generic Code for all scripts ****************'
		'Change Status from 0 to 1
		Call doSysVarColcoKeydataFlagAction("APILogging",0,1)  'Change Status from 0 to 1
		Call doAPIOutboundControlAction("APIAuthMethodID",1,2)
		
		pinreqcongdet = getpinReminderprerequestValues(cardapi_data)
		Call validateprePINReminderValues()
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
		
		If rcheck <> Empty Then
			'Append_TestHTML StepCounter, "Verify Response JSON ", "Error messages displayed" , "PASSED"
			Call validateAPIjob("2012",newoutputxmlfile,newcreationxmlfile)
			Call checkJobStatus("1013","SFN_SHELL_SPRINTQA_ID_BATCH",XML_Cadrid)
			'Call validatepostPINReminderValues()
			Call verifyCardPINMailerRequestDetails(cardapi_data("PinAdviceTypeidval"),1,1)
			JobRundate = Date()
			dtsnow = Now()
			'Individual date components
			dd = Right("00" & Day(dtsnow), 2)
			mm = Right("00" & Month(dtsnow), 2)
			yy = Year(dtsnow)
			hh = Right("00" & Hour(dtsnow), 2)
			nn = Right("00" & Minute(dtsnow), 2)
			ss = Right("00" & Second(dtsnow), 2)
			'Build the date string in the format yyyy-mm-dd
			customdatevalue = dd & "-"& mm & "-" & Mid(yy,3,2)
			
			filePath = appoutbPath & "\OutboundAPILog\" & customdatevalue
			filetype ="OutboundAPILog"

			fileoutboundlog = verifyFileExistance(filePath,filetype)
			cAddressDet = returnCustomerAddressDetails(cardapi_data)
			if fileoutboundlog <> "" and cAddressDet <> "" Then
				Call validateOutboundFileData(fileoutboundlog,cAddressDet,"txt")
			'	Call validateDX501filedata(fileoutboundlog,"txt",cNum,"Pin Option",5,2,"01")
			'	Call validateDX501filedata(fileoutboundlog,"txt",cNum,"Option SPecific data",37,35," ")
			Else
					Append_TestHTML StepCounter, "OutboundAPI file data validation" , "file/Card not created" , "FAILED"
			End if
			
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
		Else
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		End If
	End If
next