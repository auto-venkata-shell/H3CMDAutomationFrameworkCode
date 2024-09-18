Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		'jobno = Get_Dictionary(ParamValDict,"jobid" & "_" & iRowCount)
		'fileNamePrefix = Get_Dictionary(ParamValDict,"FileNamePrefix" & "_" & iRowCount)
		
		
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		
		Call createSession()
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next