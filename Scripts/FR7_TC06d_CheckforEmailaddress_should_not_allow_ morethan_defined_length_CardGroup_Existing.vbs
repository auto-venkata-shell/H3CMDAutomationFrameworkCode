Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	invalidEmail1 = "SuperUserinvalidemailSuperUserinvalidemailSuperUserinvalidemailmdw@gmail.com"
	
	If strExecute = "Yes_"& countryName Then			
		cardGrpName = "CARDGROUPNEW" & randomNumber(100,999)
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call validate_cardGroup_existingCard(invalidEmail1,"1")	
	End If
	
next