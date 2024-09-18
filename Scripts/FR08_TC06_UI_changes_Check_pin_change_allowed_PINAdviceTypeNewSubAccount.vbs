Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	'sub_customerERP_id = Get_Dictionary(ParamValDict,"sub_customerERP_id" & "_" & iRowCount)
	'subCustName = randomNumber(100,999)
	fullName = "FullName"
	shortName = "ShortName"
	tradingName = "TradingName"
	sub_customerERP_id =  customerERP_sub_id
	If strExecute = "Yes_"& countryName Then
		Call precondition_subLevelCustomer(customerERP_id,sub_customerERP_id)
		Call launchApplicationnewGFN()
		Call customerSearch1(sub_customerERP_id)
		Call ui_changes_CreateSubCust(customerERP_id,sub_customerERP_id,fullName,shortName,tradingName)
		Call checkCustomer_newSubAccountDBTable(customerERP_id,sub_customerERP_id,cust_erp_sub1)
		Call check_newSubAccount_DBValidation(cust_erp_sub1)
		
	End If

next