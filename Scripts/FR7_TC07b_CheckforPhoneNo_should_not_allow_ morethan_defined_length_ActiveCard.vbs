Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)	
	invalidPhoneNo1 = "+31 9876543210"
	invalidPhoneNo2 = "+3109876543210"

	If strExecute = "Yes_"& countryName Then			
		Call launchApplicationnewGFN()
		Call navigateCustomerSummaryscreenforOtherScreens()
		
		'Call CreateCard(custERP, custVRN, cardType, EmbossType)
		cNum = cardPANNum
		statusT = "Active"
		Call navigateCMscreenviaSearchforCards(cNum, statusT )
		'Call customerSearch1(customerERP_id)
		Call validate_phoneNo(invalidPhoneNo1,invalidPhoneNo2,"Active","0")
		Set app_data = Nothing
	End If
	
next