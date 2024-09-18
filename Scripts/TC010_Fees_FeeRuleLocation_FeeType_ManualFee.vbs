Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	cmyName = Get_Dictionary(ParamValDict,"cmyName" & "_" & iRowCount)
	fuelNetwork = Get_Dictionary(ParamValDict,"fuelNetwork" & "_" & iRowCount)
	SiteGroup = Get_Dictionary(ParamValDict,"SiteGroup" & "_" & iRowCount)
	SiteID = Get_Dictionary(ParamValDict,"SiteID" & "_" & iRowCount)
	

	If strExecute = "Yes_"& countryName Then
		strDateEff = getEffectiveDate()
		Call fees_navigateFeeRuleLocation(cmyName,fuelNetwork,SiteGroup,SiteID)
		Call validate_feeRuleLocationTableData(cmyName,fuelNetwork,SiteGroup,SiteID,strDateEff)
		
	End If
next