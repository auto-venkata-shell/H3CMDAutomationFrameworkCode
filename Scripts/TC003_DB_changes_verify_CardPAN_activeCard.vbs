Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)	
	driverName = "Driver" & randomNumber(100,999)
	strStatus = "Active"

	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call navigate_cardMaintenance()
		Call db_changes_verify_cardPAN_activeStatus(strStatus,driverName)
		db_CardID = dbChanges_verifyDB_cardID_activeStatus(cardPANNum)
		Call dbChanges_verifyDB_CardPANID_activeStatus(cardPANNum,db_CardID)
		
	End If
next