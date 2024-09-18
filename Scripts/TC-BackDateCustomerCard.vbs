Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		msgbox customerERP_id
		custVRN = Get_Dictionary(ParamValDict,"CustVRN" & "_" & iRowCount)
		
		cardType = Get_Dictionary(ParamValDict,"CardType" & "_" & iRowCount)
		
		embossType = Get_Dictionary(ParamValDict,"EmbossType" & "_" & iRowCount)
		
		'PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		msgbox cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		Call BackDateCardDetails(custERP, PANNum)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next