Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		
		set cust_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		cust_data.add "custERP", customerERP_id
		
		cust_data.add "VRN", Get_Dictionary(ParamValDict,"CustVRN" & "_" & iRowCount)
		
		cust_data.add "Ctype", Get_Dictionary(ParamValDict,"CardType" & "_" & iRowCount)
		
		cust_data.add "Emtype", Get_Dictionary(ParamValDict,"EmbossType" & "_" & iRowCount)
		cust_data.add "address", Get_Dictionary(ParamValDict,"Address" & "_" & iRowCount)
		cust_data.add "email", Get_Dictionary(ParamValDict,"Email" & "_" & iRowCount)
		
		cust_data.add "cdmethod", Get_Dictionary(ParamValDict,"CarddisMethod" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		card_val =  CreateCardforMainAddressEPN(cust_data)
		
		'********************** Scenario specific variables and business functions *********************
		if card_val <> ""
			Call RunJob196()
			filePath = appDataExchange
			filetype ="CRA"

			fileCRA = verifyFileExistance(filePath,filetype)
			'Call ValidateFile(fileCRA,cNum)
			if fileCRA <> "" Then
				Call validateCRAAddressvalue(fileCRA)
				
			Else
				Append_TestHTML StepCounter,"Validate Address in CRA File","CRA file not generated", "FAILED"
				bRunFlag = False
			End if
		else
		
		End If
	End If
next