Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	
	If ucase(strExecute) = ucase(strExecutevalue) Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		inputfilepathval = Get_Dictionary(ParamValDict,"TempFileName" & "_" & iRowCount)
		inboundcountryval = Get_Dictionary(ParamValDict,"InboundCountry" & "_" & iRowCount)
		
		FilePrefixNameval = Get_Dictionary(ParamValDict,"FilePrefixName" & "_" & iRowCount)
		'cardnoval = Get_Dictionary(ParamValDict,"CardNumber" & "_" & iRowCount)
		'cardexpdval = Get_Dictionary(ParamValDict,"CardExpirydate" & "_" & iRowCount)
		cardnoval = cardPANNum
		cardexpdval = cardExpiry_date
		fPath = sCurrentDirectory & "Test Data\" & inputfilepathval
		'inboundfolderpath = "\\AEWNW00235iis1.europe.shell.com\inbound\SFNSprint\"& inboundcountryval
		inboundfolderpath = appinbPath
		
		dpath= inboundfolderpath & "\"
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		'msgbox fPath
		'msgbox dPath
		'msgbox FilePrefixNameval
		'msgbox cardnoval
		'msgbox cardexpdval
		DX26newFilename = createandmoveDX026File(fPath,dPath,FilePrefixNameval,cardnoval,cardexpdval)
		
		if DX26newFilename <> "" Then
			Call verifyDX026FileStatus(inboundfolderpath,DX26newFilename)
		End IF
		'********************** Scenario specific variables and business functions *********************
	
	End If
next