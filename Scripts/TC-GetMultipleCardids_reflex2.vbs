Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
'msgbox iTotalRows
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		'cNum = Get_Dictionary(ParamValDict,"CardNum" & "_" & iRowCount)
		'msgbox cNum
		cNum = Split(MCarddetails,";")
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		
		'Call GetCardID(cNum,10)
	for itr=0 to ubound(cNum)
	
		if cNum(itr) <>"" Then
			Call validateCardStatus(cNum(itr),"10")
			
		End If
	Next
			Call RunJob196()
			wait 10
			environment.Value("StatusID") = getStatusID(environment.Value("cardPANNum"))
			
			'filePath ="\\AEWNW00235iis1.europe.shell.com\DataExchange\SFNSprint\CardManufacturing"
			filePath = appDataExchange
			filetype ="CRA"

			fileCRA = verifyFileExistance(filePath,filetype)
			'Call ValidateFile(fileCRA,MCarddetails)
			if fileCRA <> "" Then
				Call validateFiledata(fileCRA,MCarddetails,"txt")
			End if
			filetype ="PMA"
			filePMA = verifyFileExistance(filePath,filetype)
			'Call ValidateFile(filePMA,MCarddetails)
			if filePMA<> "" Then
				Call validateFiledata(filePMA,MCarddetails,"txt")
				for itr=0 to ubound(cNum)
			
				if cNum(itr) <>"" Then
					'Call GetCardID(cNum,1)
					Call validateCardStatus(cNum(itr),"1")		'return card id
				End If 
				
				Next
			End If
		
	End If
next