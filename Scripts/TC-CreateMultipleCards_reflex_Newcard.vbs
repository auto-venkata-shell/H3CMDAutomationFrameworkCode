Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
MCarddetails = ""
multiplecustERPs=""
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
			
		custERP = Get_Dictionary(ParamValDict,"CustomerERP" & "_" & iRowCount)
		cardType = Get_Dictionary(ParamValDict,"CardType" & "_" & iRowCount)
		custCategory = Get_Dictionary(ParamValDict,"CardCategory" & "_" & iRowCount)
		embossType = Get_Dictionary(ParamValDict,"EmbossType" & "_" & iRowCount)
		embossingval = Get_Dictionary(ParamValDict,"Embossing" & "_" & iRowCount)
		CardGrpName = Get_Dictionary(ParamValDict,"CardGrpName" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		if multiplecustERPs <> "" Then
			custERPcompval = split(multiplecustERPs, ";")
			custERPcomplastval = custERPcompval(ubound(custERPcompval)-1)
		Else
			custERPcomplastval = multiplecustERPs
		End If
		
		'********************** Scenario specific variables and business functions *********************
		if instr(custERPcomplastval,custERP)= 0 Then
			Call openAppln(url)
			Call navigateExistingcustomerSummaryScreen(custERP)
			multiplecustERPs = multiplecustERPs & custERP & ";"
		Else
		
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_searchtab", "Click", ""
		
		End If
		
		Carddetails =  createAnyCardType_reflex(cardType,custCategory,embossType,embossingval)
		if iRowCount < iTotalRows Then
			MCarddetails = MCarddetails & Carddetails & ";"
			
		Else
			MCarddetails = MCarddetails & Carddetails
			multiplecustERPs = Replace(custERP & ";", custERP)
		End If
		'********************** Scenario specific variables and business functions *********************
	
	End If
next