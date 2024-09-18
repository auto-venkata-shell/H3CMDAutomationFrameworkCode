Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		Call CreateCustCardTypes(custERP)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next