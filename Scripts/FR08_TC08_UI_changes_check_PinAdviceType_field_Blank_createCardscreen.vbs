Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)		
		
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call validate_pinAdviceType_dropdown_createCard(customerERP_id)
		Call pinAdviceType_emptyValue()
		
	End If
next