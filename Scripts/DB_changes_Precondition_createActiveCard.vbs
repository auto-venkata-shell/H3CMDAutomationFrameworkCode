Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'strExecutevalue = "Yes" & "_" & countryName
	

	If strExecute = "Yes" Then
		'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
					
		cardTypeval = Get_Dictionary(ParamValDict,"CardType" & "_" & iRowCount)		
		embossTypeval = Get_Dictionary(ParamValDict,"EmbossType" & "_" & iRowCount)
		embossName = "DRIVERNAME" & randomnumber(100,999)
		cardGroup = "#1"
	
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call GFN_CreateCard(cardGroup, embossName, cardTypeval, embossTypeval)
		cNum = cardPANNum
		if cNum <>"" Then
			Call verifysysVarcolcoandupdate("CardProductionSystem",1, "AwaitCardSecurityData",0)
			Call validateCardStatus(cNum,"10")

			'Call RunJob("196")
			Call RunJob196()
			wait 10
			'Call GetCardID(cNum,1)
			Call validateCardStatus(cNum,"1")		'return card id
		End If 
	
	End If
next