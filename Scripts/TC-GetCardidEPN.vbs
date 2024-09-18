Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
'msgbox iTotalRows
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'cNum = Get_Dictionary(ParamValDict,"CardNum" & "_" & iRowCount)
		'msgbox cNum
		cNum = cardPANNum
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		'Call GetCardID(cNum,10)
		if cNum <>"" Then
			Call validateCardStatus(cNum,"10")
			'msgbox "Test"
			
			Call RunJob196()
			wait 10
			
			'filePath ="\\AEWNW00235iis1.europe.shell.com\DataExchange\SFNSprint\CardManufacturing"
			filePath = appDataExchange
			filetype ="CRA"

			fileCRA = verifyFileExistance(filePath,filetype)
			'Call ValidateFile(fileCRA,cNum)
			if fileCRA <> "" Then
				Call validateFiledata(fileCRA,cNum,"txt")
				Call validateCRADispatchMethodvalue(fileCRA,cNum,"225","C")
			End if
			filetype ="PMA"
			filePMA = verifyFileExistance(filePath,filetype)
			'Call ValidateFile(filePMA,cNum)
			if filePMA<> "" Then
				Call validateFiledata(filePMA,cNum,"txt")
				Call validatePMADispatchMethodvalue(filePMA,cNum,"216","EP")
			End If
			'Call GetCardID(cNum,1)
			Call validateCardStatus(cNum,"1")		'return card id
		End If 
		'********************** Scenario specific variables and business functions *********************
	
	End If
next


	