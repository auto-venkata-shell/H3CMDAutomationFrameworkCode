Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		'PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		fileName = Get_Dictionary(ParamValDict,"FileName" & "_" & iRowCount)
		'msgbox dx026FinalFName
		fileName = dx026FinalFName
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		Call VerifyFileWatcherJob(custERP, PANNum, fileName)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next