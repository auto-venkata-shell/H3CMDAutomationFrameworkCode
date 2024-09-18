Dim DictTbl	
'msgbox "in"
iTotalRows = ParamValDict.Item("DATAROWS")
'msgbox iTotalRows
For iRowCount=1 to iTotalRows 
	'msgbox "DX300A"
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		bName = Get_Dictionary(ParamValDict,"BankName" & "_" & iRowCount)
		
		'fPath = Get_Dictionary(ParamValDict,"FilePath" & "_" & iRowCount)
		
		ftype = Get_Dictionary(ParamValDict,"FileType" & "_" & iRowCount)
		
		valdata = custERP & ";" & custBankNamedesc
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		fPath = appoutbPath & "\DX300A"
		Call VerifyDX300A()
		
		wait 20
		
		fLocation = verifyFileExistance(fPath,ftype)
		if fLocation<> "" Then
			Call validateFiledata(fLocation,valdata,ftype)
		Else
			Append_TestHTML StepCounter, "DX300A File Identification", "File not generated in the source path:- "&fPath , "FAILED"
			bRunFlag = False
			bFlag = False
		End If
		'********************** Scenario specific variables and business functions *********************
	
	End If
next