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
		'cardapi_data.add "custcardERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		cardapi_data.add "ColcoIDval", Get_Dictionary(ParamValDict,"ColcoID" & "_" & iRowCount)
		cardapi_data.add "CardCatIDval", Get_Dictionary(ParamValDict,"CardCatID" & "_" & iRowCount)
		cardapi_data.add "VehRegNoval", Get_Dictionary(ParamValDict,"VehRegNo" & "_" & iRowCount)
		cardapi_data.add "isVRNumval", Get_Dictionary(ParamValDict,"isVRNval" & "_" & iRowCount)
		cardapi_data.add "PinTypeIDval", Get_Dictionary(ParamValDict,"PinTypeID" & "_" & iRowCount)
		cardapi_data.add "PinAdviceTypeidval", Get_Dictionary(ParamValDict,"PinAdviceval" & "_" & iRowCount)
		inputfilepathval = "CardDeliveryAddress.txt"
		outputfilepathval = "CardDeliveryAddress.txt"
			
		'*********** Initialize Global Data for JSON Request *****************************************
		customerERP_id = cardapi_data("custcardERP")
		
		reqType = "POST"
		apiurl = api_url & "/Card/Create_CardOrder"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CSSReg\CardOrder\" & inputfilepathval 
		'headers = "Content-Type==>application/json,charset==>utf-8,ClientID==>" & client_ID & ",ClientSecretKey==>" & client_SecretKey & ",Authorization==>Basic c2hlbGxfaDNfdGVzdDozZGFjOWQyNy1iZTkwLTRiYzgtYjQ2YS03N2I3MmJhODdiNDg=,RequestId==>1544def3-bdad-4d5e-ac38-33er43,UserCulture==>en-GB"
		newfile = sCurrentDirectory & "Test Data\CSSReg\CardOrderoutput\" & outputfilepathval
		newoutputxmlfile = sCurrentDirectory & "Test Data\CSSReg\CardOrder\UpdateXML.xml" 
		newcreationxmlfile = sCurrentDirectory & "Test Data\CSSReg\CardOrder\CreationInfoXML.xml" 
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
		
		
		
		custCarddet = getCustomerValue(cardapi_data)
		custCardNo = Split(custCarddet,"|")(0)
		custColcoval = Split(custCarddet,"|")(1)
		cardTypeval = Split(custCarddet,"|")(2)
		tokenTypeval = Split(custCarddet,"|")(3)
		'tokentypeval = getRequirePINTokenTypeIDvalue(0)
		Call changeRequirePINvalue(tokenTypeval,1)
		
		'********************** Scenario specific variables and business functions *********************

		'newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		'msgbox newERPCustNo
		
		'********************** Scenario specific variables and business functions *********************
		searchreplaceString = "ColCoID-" & custColcoval &";CustomerID-"& custCardNo &";CardCategoryID-" & cardapi_data("CardCatIDval") &";PINTypeID-"& cardapi_data("PinTypeIDval") &";PINAdviceType-"& cardapi_data("PinAdviceTypeidval") &";CardTypeID-"& cardTypeval &";TokenTypeID-"& tokentypeval &";EmbossingID-3"
		
		Call searchandReplaceMultipleString(jsonFPath, searchreplaceString)
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		
			'Call verifyPreCustomerActiveDetails(newERPCustNo)
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
			
			Call validateAPIjob("2010",newoutputxmlfile,newcreationxmlfile)
			XML_Cadrid = findValuefromString(newcreationxmlfile,"<Identity>","</Identity>")
			XML_Cardpanid = findValuefromString(newcreationxmlfile,"<Identity2>","</Identity2>")
			Call validatePostAPICardDetailswithoutAddress()
			Call doSysVarColcoKeydataFlagAction("AwaitCardSecurityData",1,0)
			wait 10
			Call RunJob196()
			wait 10
			'Call doSysVarColcoKeydataFlagAction("AwaitCardSecurityData",0,1)
			PAN_num = cardPANNum
			custERP = customerERP_id
			Call backdateToCarddetails(custERP, PAN_num)
			
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
			
		End If
	End If
next