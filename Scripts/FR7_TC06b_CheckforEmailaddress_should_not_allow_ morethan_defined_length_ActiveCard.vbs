Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	status = "Active"
	invalidEmail1 = "SuperUserinvalidemailSuperUserinvalidemailSuperUserinvalidemailmdw@gmail.com"
	
	If strExecute = "Yes_"& countryName Then							
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call validate_activeCard(invalidEmail1,status,"0")
	End If
	
next