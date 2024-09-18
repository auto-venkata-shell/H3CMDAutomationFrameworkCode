Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	useFleetPIN = Get_Dictionary(ParamValDict,"useFleetPIN" & "_" & iRowCount)
	selectedPIN = Get_Dictionary(ParamValDict,"selectedPIN" & "_" & iRowCount)
	SysVarID = Get_Dictionary(ParamValDict,"SysVarID" & "_" & iRowCount)
	HideSSPSDetails = Get_Dictionary(ParamValDict,"HideSSPSDetails" & "_" & iRowCount)
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
		
	If strExecute = "Yes_"& countryName Then
		Call sspsChanges_DB_queries(hideSSPSDetails,SysVarID)
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call sspsChanges_fieldEditable_cardParameter(useFleetPIN,selectedPIN)
	End If
next