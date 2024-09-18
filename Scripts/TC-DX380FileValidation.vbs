Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		'folderPath = Get_Dictionary(ParamValDict,"FolderPath" & "_" & iRowCount)
		inputxml = Get_Dictionary(ParamValDict,"Inputfilename" & "_" & iRowCount)
		fileNamePrefix = Get_Dictionary(ParamValDict,"FileNamePrefix" & "_" & iRowCount)
		
		fMFlocation = appoutbPath & "\DX380"
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		jstatus =  validateJobStatus("380",inputxml)
		if jstatus = True Then
			fclocation =  verifyFileExistance(fMFlocation,fileNamePrefix)
			if fclocation <> "" Then
				'vadata = db_settlement_doc_no & ";" & db_settlement_sum_doc_no
				vadata =  db_settlement_sum_doc_no
				Call	validateFiledata(fclocation,vadata,"Txt")
			Else
		
				Append_TestHTML StepCounter, "DX380 File Identification", "File not generated in the source path:- "&fMFlocation , "FAILED"
				bRunFlag = False
				bFlag = False
			End IF
		'********************** Scenario specific variables and business functions *********************
		End If
	End If
next