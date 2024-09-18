Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		stl_report_date = Get_Dictionary(ParamValDict,"SettlementReportDate" & "_" & iRowCount)
		
		PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		
		Call VerifySettlementJob155(stl_report_date, PAN_num)
		'********************** Scenario specific variables and business functions *********************
	
	End If
next