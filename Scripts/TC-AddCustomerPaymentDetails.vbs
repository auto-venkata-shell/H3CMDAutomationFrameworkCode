Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		cust_data.add "custERP", customerERP_id
		cust_data.add "insuredLimit", Get_Dictionary(ParamValDict,"Insured_Limit" & "_" & iRowCount)
		cust_data.add "riskType", Get_Dictionary(ParamValDict,"Risk_Type" & "_" & iRowCount)
		cust_data.add "paymentmethod", Get_Dictionary(ParamValDict,"Paymentmethod" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************

		Call AddCustomerPaymentDetails(cust_data)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next