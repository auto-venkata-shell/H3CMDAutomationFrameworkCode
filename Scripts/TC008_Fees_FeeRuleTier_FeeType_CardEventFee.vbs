Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	feeValue = Get_Dictionary(ParamValDict,"feeValue" & "_" & iRowCount)
	
	If strExecute = "Yes_"& countryName Then
		strDateEff = getEffectiveDate()
		Call navigate_feeRuleTierScreen(feeValue)
		Call validate_feeRuleTierTableData(strDateEff,feeValue)
		Call fees_navigateFeeRuleTier_maintainFeeRule()
		
	End If
next