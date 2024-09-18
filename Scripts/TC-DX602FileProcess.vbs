Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	
	If ucase(strExecute) = ucase(strExecutevalue) Then
	
		'custERP = Get_Dictionary(ParamValDict,"CustERP" & "_" & iRowCount)
		custERP = customerERP_id
		'msgbox custERP
		inputfilepathval = Get_Dictionary(ParamValDict,"TempFileName" & "_" & iRowCount)
		inboundcountryval = Get_Dictionary(ParamValDict,"InboundCountry" & "_" & iRowCount)
		
		FilePrefixNameval = Get_Dictionary(ParamValDict,"FilePrefixName" & "_" & iRowCount)
		'cardnoval = Get_Dictionary(ParamValDict,"CardNumber" & "_" & iRowCount)
		'cardexpdval = Get_Dictionary(ParamValDict,"CardExpirydate" & "_" & iRowCount)
		
		fPath = sCurrentDirectory & "Test Data\" & inputfilepathval
		'inboundfolderpath = appinbPath
		if appName = "GFN"  and appEnvName = "SPRINTQA" Then
			inboundfolderpath = appSName & "\inbound\H3Sprint\"& countryCode
		elseif appName = "GFN"  and appEnvName = "RELEASEQA" Then
			inboundfolderpath = appSName & "\Interfaces\Inbound\H3\" & countryCode
		elseif appName = "GFN"  and appEnvName = "RD" Then
			inboundfolderpath = appSName & "\inbound\RD330\"& countryCode
		elseif appName = "SFN"  and appEnvName = "SPRINTQA" Then
			inboundfolderpath = appSName & "\inbound\SFNSprint\" & countryCode
		elseif appName = "SFN"  and appEnvName = "RELEASEQA" Then
			inboundfolderpath = appSName & "\inbound\SFN\" & countryCode
		End if
		dpath= inboundfolderpath & "\"
		
		DX602_FilePrefixNameval = FilePrefixNameval
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
		'Call OpenApplication(url)
		'msgbox fPath
		'msgbox dPath
		'msgbox FilePrefixNameval
		'msgbox cardnoval
		'msgbox cardexpdval
		'msgbox custERP
		DX620newFilename = createandmoveDX602File(fPath,dPath,custERP)
		'msgbox DX620newFilename
		if DX620newFilename <> "" Then
			Call verifyDX026FileStatus(inboundfolderpath,DX620newFilename)
		End IF
		jstatus =  VerifyH3JobStatus("208",billReportDate, PANNum)
		
	
	End If
next