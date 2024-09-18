Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'stlReportDate = Get_Dictionary(ParamValDict,"BillReportDate" & "_" & iRowCount)
		stlReportDate = settllReportDate
		'PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		'fMFlocation = "\\AEWNW00235iis1.europe.shell.com\outbound\SFNSprint\85\MerchantFulfilment"
		fMFlocation = appoutbPath & "\MerchantFulfilment"
		Call setGlobaldataforFilesValidationforSettlement(PANNum)
		jstatus = validateJobStatus("289","MerchantFulfilment")
	if jstatus = True Then
		'fclocation =  verifyFileExistance(fMFlocation,"MerchantFulfilment_GFN_INV_085")
		fclocation =  verifyFileExistance(fMFlocation,"XML")
		if fclocation <> "" Then
			'vadata = db_settlement_doc_no & ";" & db_settlement_sum_doc_no
			vadata =  db_settlement_sum_doc_no
			Call	validateFiledata(fclocation,vadata,"DAT")
		End If
		'********************** Scenario specific variables and business functions *********************
		'fclocation =  verifyFileExistance(fMFlocation,"XML")
		fclocation =  verifyFileExistance(fMFlocation,"Comma Separated Values File")
		End If
	End If
next