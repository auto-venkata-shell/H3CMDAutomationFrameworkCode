Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	'CardTypeName,TokenTypeval,EmbossingTypeval,MinCardLife,MinReissuePeriod,ExpiryPeriod,EmbossingName
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		set ct_data = CreateObject("Scripting.Dictionary")
		ct_data.add "CardType", Get_Dictionary(ParamValDict,"CardTypeName" & "_" & iRowCount)
		ct_data.add "TokenType", Get_Dictionary(ParamValDict,"TokenTypeval" & "_" & iRowCount)
		ct_data.add "EmbossingType", Get_Dictionary(ParamValDict,"EmbossingTypeval" & "_" & iRowCount)
		ct_data.add "PurchaseCategory", Get_Dictionary(ParamValDict,"PurchaseCategoryval" & "_" & iRowCount)
		ct_data.add "MinCardLifeval", Get_Dictionary(ParamValDict,"MinCardLife" & "_" & iRowCount)
		ct_data.add "MinReissuePeriodval", Get_Dictionary(ParamValDict,"MinReissuePeriod" & "_" & iRowCount)
		ct_data.add "ExpiryPeriodval", Get_Dictionary(ParamValDict,"ExpiryPeriod" & "_" & iRowCount)
		ct_data.add "EmbossingNameval", Get_Dictionary(ParamValDict,"EmbossingName" & "_" & iRowCount)
		ct_data.add "Cardcategory", Get_Dictionary(ParamValDict,"Cardcategoryval" & "_" & iRowCount)
		ct_data.add "CustVRNval", Get_Dictionary(ParamValDict,"CustVRN" & "_" & iRowCount)
		
		custERP = customerERP_id
		'cardPANNum = ""
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		Call OpenApplication(url)
		CardNo = cardPANNum
		statusT = "Active"
		Call navigateCMscreenviaSearchforCards(CardNo, statusT )
		Call performCardActions("Cards_Link_copycard")
		Call copyCardsValidation()
		Call navigateToCMSearchTab("Cards_WebEdit_DriverNamedfield","NewCDriver","1")
		'TC001 & TC002 & TC005
		'Call OpenApplication(url)
'		Call customerSearch()
'		Call fillCardTypeDetails(ct_data)
'		Call CreateandSaveCard(ct_data)
'		
'		Call navigateCustomerSummaryscreenforOtherScreens()
'		Call navigateCardMaintanceScreen()
'		Call searchCardsinCardMaitenance("New")
'								'Call searchCardsinCardMaitenance() -- Function not defined
'		Call performCardActions("Cards_Link_chagestatus")
'		Call cardCancellAction()
'		Call searchCardsinCardMaitenance("Cancelled")
		'********************** Scenario specific variables and business functions *********************
		set ct_data = Nothing
	End If
next