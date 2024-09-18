Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'folderPath = Get_Dictionary(ParamValDict,"FolderPath" & "_" & iRowCount)
		filePrefix = Get_Dictionary(ParamValDict,"FileNamePrefix" & "_" & iRowCount)
		
		filedir = appoutbPath & "\IFXCore"
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call  validateJobStatus("305",filePrefix)
		'Call  validateJobStatus("306",fileNamePrefix)
		
		jstatus =  VerifyH3JobStatus("305",billReportDate, PANNum)
		jstatus =  VerifyH3JobStatus("306",billReportDate, PANNum)
		fclocation =  verifyFileExistance(filedir,filePrefix)
		if fclocation <> "" Then
			vadata = db_bill_doc_no & ";" & db_bill_sum_doc_no
		
		'Call	validateFiledata(fclocation,vadata,"DAT")
		Else
		
			Append_TestHTML StepCounter, "IFXCore File Identification", "File not generated in the source path:- "&filedir , "FAILED"
			bRunFlag = False
			bFlag = False
		End IF
		'********************** Scenario specific variables and business functions *********************
		query = "Select * from FileSeq where FileSeqID like '%DX053%' ;"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_LastSequence=  dbRecordSet("LastSequence")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		myprefixval = Replace(db_LastSequence,"DX053","IFXCore")
		Set dbRecordSet = execute_db_query("Select * from Customer where CustomerERP = '" & cust_erp & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		db_cust_id = dbRecordSet("CustomerID")
		Set dbRecordSet = Nothing
		myprefixval = myprefixval & "_" & db_cust_id
		
		fclocation =  verifyFileExistance(filedir,myprefixval)
	End If
next