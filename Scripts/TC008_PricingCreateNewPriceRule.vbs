Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	strDesc = Get_Dictionary(ParamValDict,"strDesc" & "_" & iRowCount)
	
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call create_NewPriceRule(strDesc)
		effDate = getEffectiveDate()
		Call verifyPrePopulatedData_customerPriceRuleCreatePage(strDesc,effDate)

		Call verifyEnteredData_customerPriceRuleCreatePage(strDesc,effDate)

	End If
next