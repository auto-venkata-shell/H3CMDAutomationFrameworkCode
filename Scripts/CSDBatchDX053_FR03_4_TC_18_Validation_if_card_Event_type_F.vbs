Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
				
		'********************** Scenario specific variables *********************
		inputfilepathval = Get_Dictionary(ParamValDict,"TempFileName" & "_" & iRowCount)
		recCount = Get_Dictionary(ParamValDict,"RecordCount" & "_" & iRowCount)
		etType = Get_Dictionary(ParamValDict,"EventType" & "_" & iRowCount)
		
		inboundfolderpath = appinbPath
		dpath= inboundfolderpath & "\"
		
		fPath = sCurrentDirectory & "Test Data\" & inputfilepathval
					
		'********************** Scenario business functions *********************
		
		DX053newFilename = csdImportPreConfigDetailsandcreateDX053File(dPath, fPath, recCount, etType)
		
		if DX053newFilename <> "" Then
			Call verifyDX026FileStatus(inboundfolderpath,DX053newFilename)
			Call verifyJob230and2011Entries()
			Call verifyCardRFIDTableDetails()
		End IF
		
	End If
next