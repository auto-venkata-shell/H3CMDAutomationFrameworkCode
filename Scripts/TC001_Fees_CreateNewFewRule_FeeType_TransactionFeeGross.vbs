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
	
	
	If strExecute = "Yes_"& countryName Then
		'Call pricing_updateRecordStatusinCSV(strDesc)
		availableFrom = getEffectiveDate()
		strDesc = "TransFeeGross" & randomNumber(100,999)
		Call launchApplicationnewGFN()
		Call fees_navigateMaintainFeeRules(feeType,strDesc,freq,productGroup,feeBasis,volInclRate,waiveOnPastDue,waiveIfCancelled,minValue,maxValue,availableFrom,strCurrency)
	End If
next