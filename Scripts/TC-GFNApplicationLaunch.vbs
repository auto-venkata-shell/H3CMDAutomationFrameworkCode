Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	 strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		'set app_data = CreateObject("Scripting.Dictionary")
		'app_data.add "appurl", Get_Dictionary(ParamValDict,"URL" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call launchApplication_GFN()
		Call launchApplicationnewGFN()
		
		'********************** Scenario specific variables and business functions *********************
		Set app_data = Nothing
	End If
	
next
