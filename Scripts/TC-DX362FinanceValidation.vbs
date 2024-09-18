Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		folderPath = Get_Dictionary(ParamValDict,"FolderPath" & "_" & iRowCount)
		fileNamePrefix = Get_Dictionary(ParamValDict,"FileNamePrefix" & "_" & iRowCount)
		
		filedir = appoutbPath &"\DX362"
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		jstatus =  VerifyH3JobStatus("263",billReportDate, PANNum)
		if jstatus = True Then
			fclocation =  verifyFileExistance(filedir,fileNamePrefix)
		
			vadata = db_bill_doc_no & ";" & db_bill_sum_doc_no
		
			Call validateFiledata(fclocation,vadata,"DAT")
		Else
		
			Append_TestHTML StepCounter, "DX362 File Identification", "File not generated in the source path:- "&filedir , "FAILED"
			bRunFlag = False
			bFlag = False
		'********************** Scenario specific variables and business functions *********************
		End If
	End If
next