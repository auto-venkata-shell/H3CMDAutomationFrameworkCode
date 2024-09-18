Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		billReportDate = Get_Dictionary(ParamValDict,"BillReportDate" & "_" & iRowCount)
		
		'PANNum = Get_Dictionary(ParamValDict,"PANNum" & "_" & iRowCount)
		PANNum = cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		Call setGlobaldataforFilesValidation(PANNum)
		'jstatus =  validateJobStatus("211","DX301")
		jstatus =  VerifyH3JobStatus("263",billReportDate, PANNum)
	if jstatus = True Then
		filedir = appoutbPath & "\DX362"
		fclocation =  verifyFileExistance(filedir,"DX362_GFN_AR_086_")
		'msgbox fclocation
		DX301_FileLoc = fclocation
		DX301_FName = findFileNamefromPath(fclocation)
		if fclocation <> "" Then
			vadata = db_bill_doc_no & ";" & db_bill_sum_doc_no
			
			Call	validateFiledata(fclocation,vadata,"DAT")
		
		'********************** Scenario specific variables and business functions *********************
			'Call validatenormalJobStatus("208",fclocation)
		Else
		
			Append_TestHTML StepCounter, "DX362 File Identification", "File not generated in the source path:- "&filedir , "FAILED"
			bRunFlag = False
			bFlag = False
		End If
	End If
	End If
next