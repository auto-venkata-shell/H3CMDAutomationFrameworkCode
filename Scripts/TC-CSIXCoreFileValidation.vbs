Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		'foldpath = Get_Dictionary(ParamValDict,"FolderPath" & "_" & iRowCount)
		fileNamePrefix = Get_Dictionary(ParamValDict,"FileNamePrefix" & "_" & iRowCount)
		
		filedir = appoutbPath & "\CSIXCore"
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call  validateJobStatus("264",fileNamePrefix)
		jstatus =  VerifyH3JobStatus("305",billReportDate, PANNum)
		jstatus =  VerifyH3JobStatus("306",billReportDate, PANNum)
		fclocation =  verifyFileExistance(filedir,fileNamePrefix)
		if fclocation <> "" Then 
			vadata = db_bill_doc_no & ";" & db_bill_sum_doc_no
		
		'Call	validateFiledata(fclocation,vadata,"Excel")
		Else
		
			Append_TestHTML StepCounter, "CSIXCore File Identification", "File not generated in the source path:- "&filedir , "FAILED"
			bRunFlag = False
			bFlag = False
		End If
		'********************** Scenario specific variables and business functions *********************
	
	End If
next