Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		cust_data.add "custERP", customerERP_id
		cust_data.add "oGuarantType", Get_Dictionary(ParamValDict,"GuarantType" & "_" & iRowCount)
		cust_data.add "oGCurrency", Get_Dictionary(ParamValDict,"GCurrency" & "_" & iRowCount)
		cust_data.add "oGEDate", Get_Dictionary(ParamValDict,"GEDate" & "_" & iRowCount)
		cust_data.add "oGDetails", Get_Dictionary(ParamValDict,"GDetails" & "_" & iRowCount)
		cust_data.add "oGAdjustmentval", Get_Dictionary(ParamValDict,"GAdjustmentval" & "_" & iRowCount)
		
		
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		cgdetails = Get_Dictionary(ParamValDict,"GDetails" & "_" & iRowCount)
		
		cgadjustment = Get_Dictionary(ParamValDict,"GAdjustmentval" & "_" & iRowCount)
		
		ftype = Get_Dictionary(ParamValDict,"FileType" & "_" & iRowCount)
		
				'valdata = custERP & ";" & custBankNamedesc
		valdata = custERP & ";" & cgdetails & ";" & cgadjustment
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		fPath = appoutbPath & "\DX300B"
		
		Call VerifyCustomerGuaranteeCreation(cust_data)
		
		Call VerifyDX300B()
		
		wait 20
		
		fLocation = verifyFileExistance(fPath,ftype)
		if fLocation<> "" Then
			Call validateFiledata(fLocation,valdata,ftype)
		Else
			Append_TestHTML StepCounter, "DX300B File Identification", "File not generated in the source path:- "&fPath , "FAILED"
			bRunFlag = False
			bFlag = False
		End If
		'********************** Scenario specific variables and business functions *********************
	
	End If
next