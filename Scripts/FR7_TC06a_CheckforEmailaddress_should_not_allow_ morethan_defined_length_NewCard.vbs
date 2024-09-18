Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	embossName = Get_Dictionary(ParamValDict,"embossName" & "_" & iRowCount)
	cardType = Get_Dictionary(ParamValDict,"cardType" & "_" & iRowCount)
	embossType = Get_Dictionary(ParamValDict,"embossType" & "_" & iRowCount)
	invalidEmail1 = "SuperUserinvalidemailSuperUserinvalidemailSuperUserinvalidemailmdw@gmail.com"	




	If strExecute = "Yes_"& countryName Then					
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call GFN_CreateCard("#1", embossName, cardType, embossType)
		Call validate_newCard_email_field(invalidEmail1,"0")
	End If
next