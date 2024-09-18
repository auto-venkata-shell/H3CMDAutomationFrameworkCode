Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		'msgbox customerERP 
		'msgbox customerERP_id
		custERP = customerERP_id
		outputType = Get_Dictionary(ParamValDict,"OutputType" & "_" & iRowCount)
		
		contact = Get_Dictionary(ParamValDict,"Contact" & "_" & iRowCount)
		
		altDistMethod = Get_Dictionary(ParamValDict,"AltDistMethod" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		Call CreateCustInfoSub(custERP, outputType, contact, altDistMethod)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next