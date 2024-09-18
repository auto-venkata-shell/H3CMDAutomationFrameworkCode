Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		cust_data.add "custERP", customerERP_id
		cust_data.add "foreName", Get_Dictionary(ParamValDict,"Forename" & "_" & iRowCount)
		cust_data.add "surName", Get_Dictionary(ParamValDict,"Surname" & "_" & iRowCount)
		cust_data.add "email", Get_Dictionary(ParamValDict,"Email" & "_" & iRowCount)
		cust_data.add "Reg", Get_Dictionary(ParamValDict,"Register" & "_" & iRowCount)
		cust_data.add "MobileNo", Get_Dictionary(ParamValDict,"MobileNmber" & "_" & iRowCount)
		
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************

		Call AddCustomerContactERP(cust_data)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next