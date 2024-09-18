Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
			
	If strExecute = "Yes_"& countryName Then		
		Call launchApplicationnewGFN()
		Call verifyFields_cardTypeScreen()
	End If
next