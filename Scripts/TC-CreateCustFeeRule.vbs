Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName & "_" & appEnvName
	
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		'custERP = customerERP_id
		
		
		'--- New lines
		set fee_data = CreateObject("Scripting.Dictionary")
		fee_data.add "custERP", customerERP_id
		fee_data.add "items", Get_Dictionary(ParamValDict,"NoOfSerchItems" & "_" & iRowCount)
		fee_data.add "item1", Get_Dictionary(ParamValDict,"searchitem1" & "_" & iRowCount)
		fee_data.add "item2", Get_Dictionary(ParamValDict,"searchitem2" & "_" & iRowCount)
		fee_data.add "item3", Get_Dictionary(ParamValDict,"searchitem3" & "_" & iRowCount)
		fee_data.add "langauge", Get_Dictionary(ParamValDict,"Language" & "_" & iRowCount)
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		'Call CreateCustFeeRule(custERP)
		
		'New calling statement
		Call addCustFeeRule(fee_data)
		
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next