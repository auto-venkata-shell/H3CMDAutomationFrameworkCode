Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	embossName = Get_Dictionary(ParamValDict,"embossName" & "_" & iRowCount)
	cardType = Get_Dictionary(ParamValDict,"cardType" & "_" & iRowCount)
	embossType = Get_Dictionary(ParamValDict,"embossType" & "_" & iRowCount)
	invalidPhoneNo1 = "+31 9876543210"
	invalidPhoneNo2 = "+3109876543210"
	
	If strExecute = "Yes_"& countryName Then					
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call GFN_CreateCard_of_PINAdviceType("#1", embossName, cardType, embossType,"SMS")
		Call validate_newCard_PhoneNo_field(invalidPhoneNo1,invalidPhoneNo2,"0")
		Set app_data = Nothing
	End If
	
next