Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)	
	strStatus = "Active"

	reason = "Lost"
		
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		wait 4
		Call customerSearch1(customerERP_id)
		Call navigate_cardMaintenance()
		Call precondition_blockCard(strStatus,reason)
	
	End If
next