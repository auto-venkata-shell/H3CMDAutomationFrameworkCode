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
		Call verifysysVarcolcoandupdate("CardProductionSystem",1, "AwaitCardSecurityData",0)
		Call OpenApplication(url)
		Call updateDatesforRenewcard()
		Call navigateCustomerSummaryscreenforOtherScreens()
		
		cNum = cardPANNum
		statusT = "Active"
		Call navigateCMscreenviaSearchforCards(cNum, statusT )
		Call searchCardsinCardMaitenance(statusT)
		Call performCardActions("CardsWebLink_reissueButton")
		Call renewationdet()
		cNum = cardPANNum
		Call cardCreationchecks(custERP,cNum)
		if cNum <>"" Then
			Call validateCardStatus(cNum,"23")
			'msgbox "Test"
			
			Call RunJob196()
			wait 10
			
			
			Call validateCardStatus(cNum,"1")		'return card id
			Call verifyActivityLogdetails("7")
		End If 
		
		set ct_data = Nothing
	End If
next