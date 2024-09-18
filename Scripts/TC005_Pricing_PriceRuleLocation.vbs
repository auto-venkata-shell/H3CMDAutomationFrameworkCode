Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	locScope = Get_Dictionary(ParamValDict,"locScope" & "_" & iRowCount)
	cmyName = Get_Dictionary(ParamValDict,"cmyName" & "_" & iRowCount)
	fuelNetwork = Get_Dictionary(ParamValDict,"fuelNetwork" & "_" & iRowCount)
	
		
	If strExecute = "Yes_"& countryName Then
		Call enterInput_priceRuleLocation(locScope,cmyName)
		Call validate_PriceRuleLocationTabledata(cmyName,fuelNetwork)
	End If
next