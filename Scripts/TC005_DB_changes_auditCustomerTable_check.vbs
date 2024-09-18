Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)	
	fullName = "FullName" & randomNumber(100,999)

		
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		wait 2		
		Call customerSearch1(customerERP_id)
		ui_beforeFullName = dbChanges_UIauditCustomerTable_check(fullName)	
		Call dbChanges_DBauditCustomerTable_DBcheck(ui_beforeFullName,fullName,customerERP_id)
		
	End If
next