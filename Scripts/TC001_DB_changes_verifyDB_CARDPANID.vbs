Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)

	If strExecute = "Yes_"& countryName Then
		Call dbChanges_replication_DB_CardPANtable(cardPANNum,"GFN_SHELL_SPRINTQA_PH_OLTP")
		Call dbChanges_replication_DB_CardPANtable(cardPANNum,"GFN_SHELL_SPRINTQA_PH_BATCH")
		Call dbChanges_replication_DB_CardPANtable(cardPANNum,"GFN_SHELL_SPRINTQA_PH_REPORTS")
		Call dbChanges_replication_DB_CardPANtable(cardPANNum,"GFN_SHELL_SPRINTQA_PH_WWW")
		
	End If
next