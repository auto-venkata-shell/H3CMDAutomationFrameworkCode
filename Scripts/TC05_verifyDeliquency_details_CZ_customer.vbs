Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	enfdelLock1 = Get_Dictionary(ParamValDict,"enfdelLock1" & "_" & iRowCount)	
	enfdelLock2 = Get_Dictionary(ParamValDict,"enfdelLock2" & "_" & iRowCount)
	
	If strExecute = "Yes_"& countryName Then			

		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call customerCreditLimitMgt_fieldEnabled(enfdelLock1,enfdelLock2)
	else
		Print "StrExecute value is false"	
	End If
	
next