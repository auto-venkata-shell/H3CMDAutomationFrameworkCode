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
		
		fPath = sCurrentDirectory & "Test Data\" & inputfilepathval
		if appName = "GFN"  and appEnvName = "SPRINTQA" Then
			inboundfolderpath = appSName & "\inbound\H3Sprint\GLOBAL"
		elseif appName = "GFN"  and appEnvName = "RELEASEQA" Then
			inboundfolderpath = appSName & "\Interfaces\Inbound\H3\GLOBAL" 
		elseif appName = "GFN"  and appEnvName = "RD" Then
			inboundfolderpath = appSName & "\inbound\RD330\GLOBAL"
		elseif appName = "SFN"  and appEnvName = "SPRINTQA" Then
			inboundfolderpath = appSName & "\inbound\SFNSprint\GLOBAL" 
		elseif appName = "SFN"  and appEnvName = "RELEASEQA" Then
			inboundfolderpath = appSName & "\inbound\SFN\GLOBAL" 
		End if
		
		dpath= inboundfolderpath & "\"
		DX451_FilePrefixNameval = FilePrefixNameval
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		'msgbox fPath
		'msgbox dPath
		'msgbox FilePrefixNameval
		'msgbox cardnoval
		'msgbox cardexpdval
		DX451newFilename = createandmoveDX451File(fPath,dPath)
		'msgbox DX350newFilename
		if DX451newFilename <> "" Then
			Call verifyDX026FileStatus(inboundfolderpath,DX451newFilename)
			Call VerifyDX451Jobs(DX451newFilename)
			Call VerifyDX451data()
		End IF
		'********************** Scenario specific variables and business functions *********************
	
	End If
next