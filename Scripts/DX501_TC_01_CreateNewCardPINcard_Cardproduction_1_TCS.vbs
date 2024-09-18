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
		Call navigateCustomerSummaryscreenforOtherScreens()
		
		Call CreateCard(custERP, custVRN, cardType, EmbossType)
		cNum = cardPANNum
		Call cardCreationchecks(custERP,cNum)
		if cNum <>"" Then
			Call validateCardStatus(cNum,"10")
			
			
			Call RunJob196()
			wait 10
			
			
			filePath = appDataExchange
			filetype ="CRA"

			fileCRA = verifyFileExistance(filePath,filetype)
			
			if fileCRA <> "" Then
				Call validateFiledata(fileCRA,cNum,"txt")
				Call validateDX501filedata(fileCRA,"txt",cNum,"Pin Option",5,2,"01")
				Call validateDX501filedata(fileCRA,"txt",cNum,"Option SPecific data",37,35," ")
			End if
			filetype ="PMA"
			filePMA = verifyFileExistance(filePath,filetype)
			
			if filePMA<> "" Then
				Call validateFiledata(filePMA,cNum,"txt")
				Call validateDX501filedata(filePMA,"txt",cNum,"Dispatch Method",216,2,"M")
			End If
			
			Call validateCardStatus(cNum,"1")		'return card id
		End If 
		set ct_data = Nothing
	End If
next