Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'billReportDate = Get_Dictionary(ParamValDict,"BillReportDate" & "_" & iRowCount)
		
		'PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		Call VerifyBillingJob64(billReportDate, PANNum)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next