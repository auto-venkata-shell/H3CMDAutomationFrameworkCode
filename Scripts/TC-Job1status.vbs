Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
		
		jobno = Get_Dictionary(ParamValDict,"jobid" & "_" & iRowCount)
		'fileNamePrefix = Get_Dictionary(ParamValDict,"FileNamePrefix" & "_" & iRowCount)
		
		
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		''Call  validateJobStatus("264",fileNamePrefix)
		'fclocation =  verifyFileExistance(folderPath,fileNamePrefix)
		'msgbox fclocation
		'vadata = db_bill_doc_no & ";" & db_bill_sum_doc_no
		
		Call	validateJobonestatus()
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next