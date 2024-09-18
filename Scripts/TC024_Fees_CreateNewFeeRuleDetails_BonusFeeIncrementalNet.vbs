Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 

	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	feeType = Get_Dictionary(ParamValDict,"feeType" & "_" & iRowCount)
	freq = Get_Dictionary(ParamValDict,"freq" & "_" & iRowCount)
	productGroup = Get_Dictionary(ParamValDict,"productGroup" & "_" & iRowCount)
	feeBasis = Get_Dictionary(ParamValDict,"feeBasis" & "_" & iRowCount)
	volInclRate = Get_Dictionary(ParamValDict,"volInclRate" & "_" & iRowCount)
	minValue = Get_Dictionary(ParamValDict,"minValue" & "_" & iRowCount)
	maxValue = Get_Dictionary(ParamValDict,"maxValue" & "_" & iRowCount)
	waiveOnPastDue = Get_Dictionary(ParamValDict,"waiveOnPastDue" & "_" & iRowCount)
	waiveIfCancelled = Get_Dictionary(ParamValDict,"waiveIfCancelled" & "_" & iRowCount)
	strCurrency = Get_Dictionary(ParamValDict,"strCurrency" & "_" & iRowCount)
	strDesc = "BonusFeeIncNet" & randomNumber(100,999)
	availableFrom = getEffectiveDate()
	cmyName = Get_Dictionary(ParamValDict,"cmyName" & "_" & iRowCount)
	fuelNetwork = Get_Dictionary(ParamValDict,"fuelNetwork" & "_" & iRowCount)
	SiteGroup = Get_Dictionary(ParamValDict,"SiteGroup" & "_" & iRowCount)
	SiteID = Get_Dictionary(ParamValDict,"SiteID" & "_" & iRowCount)
	strDateEff = getEffectiveDate()
	prodGroup = Get_Dictionary(ParamValDict,"prodGroup" & "_" & iRowCount)
	Product = Get_Dictionary(ParamValDict,"Product" & "_" & iRowCount)
	feeValue = Get_Dictionary(ParamValDict,"feeValue" & "_" & iRowCount)
	
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call fees_navigateMaintainFeeRules(feeType,strDesc,freq,productGroup,feeBasis,volInclRate,waiveOnPastDue,waiveIfCancelled,minValue,maxValue,availableFrom,strCurrency)
		Call fees_navigateFeeRuleLocation(cmyName,fuelNetwork,SiteGroup,SiteID)
		Call validate_feeRuleLocationTableData(cmyName,fuelNetwork,SiteGroup,SiteID,strDateEff)
		Call fees_navigateFeeRuleProduct(prodGroup,product)
		Call validate_feeRuleProductTableData(prodGroup,product,strDateEff)
		Call navigate_feeRuleTierScreen(feeValue)
		Call validate_feeRuleTierTableData(strDateEff,feeValue)
		Call fees_navigateFeeRuleTier_maintainFeeRule()
	End If
next