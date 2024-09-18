Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	prodGroup = Get_Dictionary(ParamValDict,"prodGroup" & "_" & iRowCount)
	Product = Get_Dictionary(ParamValDict,"Product" & "_" & iRowCount)
	
	If strExecute = "Yes_"& countryName Then
		strDateEff = getEffectiveDate()
		Call fees_navigateFeeRuleProduct(prodGroup,product)
		Call validate_feeRuleProductTableData(prodGroup,product,strDateEff)
		
	End If
next