Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		cust_data.add "custERP", customerERP_id
		cust_data.add "bankType", Get_Dictionary(ParamValDict,"Bank_Type" & "_" & iRowCount)
		cust_data.add "bankName", Get_Dictionary(ParamValDict,"Bank_Name" & "_" & iRowCount)
		cust_data.add "address", Get_Dictionary(ParamValDict,"Bank_Address" & "_" & iRowCount)
		cust_data.add "city", Get_Dictionary(ParamValDict,"City" & "_" & iRowCount)
		cust_data.add "region", Get_Dictionary(ParamValDict,"Region" & "_" & iRowCount)
		cust_data.add "postCode", Get_Dictionary(ParamValDict,"Post_Code" & "_" & iRowCount)
		cust_data.add "sortCode", Get_Dictionary(ParamValDict,"Sort_Code" & "_" & iRowCount)
		cust_data.add "accNumber", Get_Dictionary(ParamValDict,"Account_Number" & "_" & iRowCount)
		cust_data.add "ddAuthName", Get_Dictionary(ParamValDict,"DD_Authorizer_Name" & "_" & iRowCount)
		cust_data.add "branchName", Get_Dictionary(ParamValDict,"BranchName" & "_" & iRowCount)
		
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************

		Call AddCustomerBankDetails(cust_data)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next