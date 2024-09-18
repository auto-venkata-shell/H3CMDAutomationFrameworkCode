Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	invalidPhone1 = "+31 9876543210"
	invalidPhone2 = "+3109876543210"		 

	If strExecute = "Yes_"& countryName Then			
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call validate_cardGroup_existingCard_phoneNo(invalidPhone1,invalidPhone2,"1")
	End If

next