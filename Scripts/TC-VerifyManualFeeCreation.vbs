Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		cust_data.add "custERP", customerERP_id
		cust_data.add "feeType", Get_Dictionary(ParamValDict,"Fee_Type" & "_" & iRowCount)
		cust_data.add "feeRule", Get_Dictionary(ParamValDict,"Fee_Rule" & "_" & iRowCount)
		cust_data.add "quantity", Get_Dictionary(ParamValDict,"Quantity" & "_" & iRowCount)
		cust_data.add "unitPrice", Get_Dictionary(ParamValDict,"Unit_Price" & "_" & iRowCount)
		cust_data.add "manualFeeText", Get_Dictionary(ParamValDict,"Manual_Fee_Text" & "_" & iRowCount)
		cust_data.add "backDate", Get_Dictionary(ParamValDict,"Back_Date" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		Call VerifyManualFeeCreation(cust_data)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next