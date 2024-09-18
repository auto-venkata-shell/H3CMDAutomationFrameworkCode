Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	enfdelLock1 = Get_Dictionary(ParamValDict,"enfdelLock1" & "_" & iRowCount)	
	DelSusDays = Get_Dictionary(ParamValDict,"DelSusDays" & "_" & iRowCount)
	
	If strExecute = "Yes_"& countryName Then			

		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call validateFields_customerCreditLimitManagement(enfdelLock1,DelSusDays)
	else
		Print "StrExecute value is false"	
	End If
	
next