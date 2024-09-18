Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	'CardTypeName,TokenTypeval,EmbossingTypeval,MinCardLife,MinReissuePeriod,ExpiryPeriod,EmbossingName
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		set ct_data = CreateObject("Scripting.Dictionary")
		custERP = customerERP_id
		custVRN = Get_Dictionary(ParamValDict,"CustVRN" & "_" & iRowCount)
		
		cardType = Get_Dictionary(ParamValDict,"CardType" & "_" & iRowCount)
		
		embossType = Get_Dictionary(ParamValDict,"EmbossType" & "_" & iRowCount)
		'ct_data.add "CardType", Get_Dictionary(ParamValDict,"CardTypeName" & "_" & iRowCount)
		'ct_data.add "TokenType", Get_Dictionary(ParamValDict,"TokenTypeval" & "_" & iRowCount)
		
		
		custERP = customerERP_id
		'cardPANNum = ""
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		Call verifysysVarcolcoandupdate("CardProductionSystem",1, "AwaitCardSecurityData",1)
		Call OpenApplication(url)
		Call navigateCustomerSummaryscreenforOtherScreens()
		
		Call CreateCard(custERP, custVRN, cardType, EmbossType)
		cNum = cardPANNum
		Call cardCreationchecks(custERP,cNum)
		if cNum <>"" Then
			Call validateCardStatus(cNum,"10")
			'msgbox "Test"
			
			Call RunJob196()
			wait 10
			wait 2
			Call validateCardStatus(cNum,"45")		
			wait 2
			Call verifyActivityLogdetails("200")
			Call cardStatusHistoryDetails()
			cNum = cardPANNum
			statusT = "Awaiting Card Security Data"
			Call navigateCMscreenviaSearchforCards(cNum, statusT )
			Call verifyCardHistoryActivityLog()
		End If 
		
		'Call verifysysVarcolcoandupdate("CardProductionSystem",1, "AwaitCardSecurityData",0)
		set ct_data = Nothing
	End If
next