Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)	
	strStatus = "Blocked Card"
	strNewDriverName = "DriverName" & randomNumber(100,999)
	reason = "Lost"
		
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call navigate_cardMaintenance()
		Call db_changes_verify_cardPAN_blockedStatus(strStatus,1)
		Call click_on_copyCard_tableItem(strNewDriverName)
		Call dbChanges_verifyDB_cardID_blockedstatus(cardPANNum)
		Call dbChanges_verifyDB_CardPANID_blockedstatus(db_PAN)
		
	End If
next