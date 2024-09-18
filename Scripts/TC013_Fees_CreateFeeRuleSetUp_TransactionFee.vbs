Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	feeType = Get_Dictionary(ParamValDict,"feeType" & "_" & iRowCount)
	customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	waiveOnPastDue = Get_Dictionary(ParamValDict,"waiveOnPastDue" & "_" & iRowCount)
	waiveIfCancelled = Get_Dictionary(ParamValDict,"waiveIfCancelled" & "_" & iRowCount)

	
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call addNewFeeRule(feeType,waiveOnPastDue,waiveIfCancelled)
		Call verifyTblElement(".*SearchControl_grdResults","Fee Rule Details",strDesc)
	End If
next