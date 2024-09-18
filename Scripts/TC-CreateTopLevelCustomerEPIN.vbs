Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		set cust_data = CreateObject("Scripting.Dictionary")
		cust_data.add "lob", Get_Dictionary(ParamValDict,"LOB" & "_" & iRowCount)
		cust_data.add "fullName", Get_Dictionary(ParamValDict,"Full_Name" & "_" & iRowCount)
		cust_data.add "shortName", Get_Dictionary(ParamValDict,"Short_Name" & "_" & iRowCount)
		cust_data.add "tradingName", Get_Dictionary(ParamValDict,"Trading_Name" & "_" & iRowCount)
		cust_data.add "langauge", Get_Dictionary(ParamValDict,"Language" & "_" & iRowCount)
		cust_data.add "legalEnity", Get_Dictionary(ParamValDict,"LegalEntity" & "_" & iRowCount)
		cust_data.add "regNum", Get_Dictionary(ParamValDict,"Registration_Number" & "_" & iRowCount)
		cust_data.add "vat", Get_Dictionary(ParamValDict,"VAT" & "_" & iRowCount)
		cust_data.add "band", Get_Dictionary(ParamValDict,"Band" & "_" & iRowCount)
		cust_data.add "marketingSeg", Get_Dictionary(ParamValDict,"Marketing_Segmentation" & "_" & iRowCount)
		cust_data.add "priceProfile", Get_Dictionary(ParamValDict,"Price_Profile" & "_" & iRowCount)
		cust_data.add "feeGroupName", Get_Dictionary(ParamValDict,"Fee_Group_Name" & "_" & iRowCount)
		cust_data.add "acqChannel", Get_Dictionary(ParamValDict,"Acquisition_Channel" & "_" & iRowCount)
		cust_data.add "custClassification", Get_Dictionary(ParamValDict,"Customer_Classification" & "_" & iRowCount)
		cust_data.add "creditLimit", Get_Dictionary(ParamValDict,"Credit_Limit" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		Call OpenApplication(url)
		Call CreateTopLevelCustomerEPIN(cust_data)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next