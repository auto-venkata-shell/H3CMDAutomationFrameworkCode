Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		setReportDate = Get_Dictionary(ParamValDict,"SettlementReportDate" & "_" & iRowCount)
		
		'PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		bSingleSignOnFlag = True
		bSignOnFlag = False
		Call OpenApplication(url)
		
		Call VerifySettlementTxn(setReportDate, PANNum)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next