Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		custVRN = Get_Dictionary(ParamValDict,"CustVRN" & "_" & iRowCount)
		
		cardType = Get_Dictionary(ParamValDict,"CardType" & "_" & iRowCount)
		
		embossType = Get_Dictionary(ParamValDict,"EmbossType" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		Call CreateCard(custERP, custVRN, cardType, EmbossType)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next