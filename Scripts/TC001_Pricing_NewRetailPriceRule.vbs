Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	locScope = Get_Dictionary(ParamValDict,"locScope" & "_" & iRowCount)
	proScope = Get_Dictionary(ParamValDict,"proScope" & "_" & iRowCount)
	priceList = Get_Dictionary(ParamValDict,"priceList" & "_" & iRowCount)
	cmyName = Get_Dictionary(ParamValDict,"cmyName" & "_" & iRowCount)
	availableFrom = Get_Dictionary(ParamValDict,"availableFrom" & "_" & iRowCount)
	refPrice = Get_Dictionary(ParamValDict,"refPrice" & "_" & iRowCount)
	priceRuleBasis = Get_Dictionary(ParamValDict,"priceRuleBasis" & "_" & iRowCount)
	lstcurrency = Get_Dictionary(ParamValDict,"lstcurrency" & "_" & iRowCount)
	funder = Get_Dictionary(ParamValDict,"funder" & "_" & iRowCount)
	discount = Get_Dictionary(ParamValDict,"discount" & "_" & iRowCount)
	IsCustomer = Get_Dictionary(ParamValDict,"IsCustomer" & "_" & iRowCount)
	priceRuleCat = Get_Dictionary(ParamValDict,"priceRuleCat" & "_" & iRowCount)
	strDesc = "PriceRule" & randomNumber(100,999)

	If strExecute = "Yes_"& countryName Then
		Call pricing_updateRecordStatusinCSV(strDesc)
		Call launchApplicationnewGFN()
		Call enterInput_searchPriceRulePage(strDesc,locScope,proScope,priceList,cmyName,availableFrom,refPrice,priceRuleBasis,lstcurrency,funder,IsCustomer,discount,priceRuleCat)
		Call validate_resultTable(strDesc)
		Call validate_DB(strDesc)
	End If
next