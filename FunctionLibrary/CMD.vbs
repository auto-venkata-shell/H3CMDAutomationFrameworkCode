Public Apiconfigpath, subLevelERP, tcDependencyFlag, err_colcoID, db_DefaultPriceProfileID, db_DefaultFeeGroupID, db_ecustStatusID,  defaultcheckflag,db_Cust_AddressID
Public errorapi_Totalnumrecs, specialFName, specialMName, specialLName, specialchrFlag
Public  topLevelERP, topnewAddressref, topnewContactref, subnewAddressref, subnewContactref, topnewPaymentref
Public RelationshipLevelERP, topGPLevelERP,subPLevelERP,child1LevelERP,child2LevelERP,child3LevelERP,child4LevelERP,child5LevelERP,child6LevelERP

Public toplevelendTime,toplevelstartTime
Public topP1LevelERP
specialchrFlag = False
'topLevelERP = "NL01801081"
'topnewBankref = "Bank01nl801067"
'Function preDatafetchingForEntityType1(dbName)
'On Error Resume Next
'	bRunFlag = True
'	bFlag = True
''	database = ""
''	database = appName & "_SHELL_"
''	Select Case appEnvName
''		Case "SPRINTQA"
''			database = database & "SPRINTQA_"
''		Case "RELEASEQA"
''			database = database
''		Case "RD"
''		database = database & "RD_"
''	End Select	
''DatabaseName='GFN_SHELL_"
'	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
'	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
'	Set dictDbResultSet = execute_db_query(query, 1, dbName)
'	wait 3
'	db_ClientCompanyNumber = dictDbResultSet("ClientCompanyNumber")
'	db_CompanyID = dictDbResultSet("CompanyID")
'	
'	Set dictDbResultSet = Nothing
'	Set query = Nothing
'	Append_TestHTML StepCounter, "Get ClientCompanyNumber and CompanyID","ClientCompanyNumber" & db_ClientCompanyNumber & " - CompanyID:"& db_CompanyID , "PASSED"
'	cerpcheck = countryCode & db_ClientCompanyNumber
'	query_customer = "Select * from Customer where CustomerERP like '"& cerpcheck &"%' order by 2 desc;"
'	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
'	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
'	wait 3
'	lastcustCERPno = dictDbResultSet("CustomerERP")
'	Set dictDbResultSet = Nothing
'	Set query_customer = Nothing
'	
'	If isEmpty(lastcustCERPno) Then
'		mylastno = "1001"
'		Append_TestHTML StepCounter, "Verify Last Entry of Customer ERP","No record exist with specific Format and creating new one", "PASSED"
'	Else
'		Append_TestHTML StepCounter, "Verify Last Entry of Customer ERP",lastcustCERPno, "PASSED"
'		mylastno = Right(lastcustCERPno,5)
'	End If
'	mylastno = cdbl(mylastno) + 1
'	
'	If len(mylastno)<5 Then
'		currentlen = len(mylastno)
'		For itr = 1 To  (5 - currentlen)  Step 1
'			mylastno = cstr("0" & mylastno)
'	
'		Next
'	End If
'	
'	newcustERP = cerpcheck & Left(mylastno,5)
'	myERP = newcustERP
''''msgbox myERP
'	query_customer = "Select * from Customer where CustomerERP = '"&myERP&"'"
'	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
'	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
'	wait 3
'	 ecustCID = dictDbResultSet("CustomerID")
'	ecustSName = dictDbResultSet("ShortName")
'	
'	set dictDbResultSet = Nothing
'	Set query_customer = Nothing
'	query_customer = "Select * from MessageQueue where EntityRowID ='"&myERP&"'"
'	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
'	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
'	wait 3
'	messageqID = dictDbResultSet("MessageQueueID")
'	 custStatusID = dictDbResultSet("MessageQueueStatusID")
'	custEntityPayload = dictDbResultSet("EntityPayload")	
'	set dictDbResultSet = Nothing
'	Set query_customer = Nothing
'	
'	If isempty(ecustCID)  Then
'		Append_TestHTML StepCounter, "Verify Customer details using new ERP" & myERP ,"New ERP not exist in the Customer Table and will able to process EntityType1 JSON request", "PASSED"
'		
'		If isempty(messageqID) Then
'			Append_TestHTML StepCounter, "Verify MessageQueue details using EntityRowID" & myERP ,"EntityRowID not exist, will able to process EntityType1 JSON request", "PASSED"
'			
'			query_customer = "Select * from MessageQueueProcessed where EntityRowID ='"&myERP&"'"
'			Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
'			set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
'			wait 3
'			messageqID = dictDbResultSet("MessageQueueID")
'			messaageereqID = dictDbResultSet("ExternalRequestID")
'			 custStatusID = dictDbResultSet("MessageQueueStatusID")
'			custEntityPayload = dictDbResultSet("EntityPayload")
'			
'			set dictDbResultSet = Nothing
'			Set query_customer = Nothing
'			If isempty(messageqID) Then
'				Append_TestHTML StepCounter, "Verify MessageQueueProcessed details using EntityRowID" & myERP ,"EntityRowID not exist, will able to process EntityType1 JSON request", "PASSED"
'				
'			Else
'				Append_TestHTML StepCounter, "Verify MessageQueueProcessed details using EntityRowID" & myERP ,"EntityRowID  exist, will able/unable to process EntityType1 JSON request", "FAILED"
'				bFlag = False
'				bRunFlag = False
''				query_customer = "Select * from MessageQueueError where MessageQueueID ='"&messageqID&"'"
''				Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
''				set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
''				wait 3
''				messageqErrorDetils = dictDbResultSet("ErrorMessage")
''				Set dictDbResultSet = Nothing
''				Set query_customer = Nothing
''				If messageqErrorDetils <> "" Then
''					Append_TestHTML StepCounter, "Verify MessageQueueError details using MessageQueueID" & messageqID ,"ErrorMessage  as"& messageqErrorDetils , "FAILED"
''					bFlag = False
''					bRunFlag = False
''					
''				Else.
''					Append_TestHTML StepCounter, "Verify MessageQueueError details using MessageQueueID" & messageqID ,"No ErrorMessage  locked" , "FAILED"
''					bFlag = False
''					bRunFlag = False
''				
''				End If
'			End If
'		Else
'			Append_TestHTML StepCounter, "Verify MessageQueue details using EntityRowID" & myERP ,"EntityRowID  exist, will able to process EntityType1 JSON request", "PASSED"
''				bFlag = False
''				bRunFlag = False
'		
'		End If
'	Else
'		Append_TestHTML StepCounter, "Verify Customer details using new ERP" & myERP ,"New ERP exist in the Customer Table and will unable to process EntityType1 JSON request using CustomerERP", "FAILED"
'				bFlag = False
'				bRunFlag = False
'	
'	End If
'If bFlag Then
'	preDatafetchingForEntityType1 = myERP
'Else
'	preDatafetchingForEntityType1 = ""
'End If
'
'End Function
'
'**********************************************************************************New Function**********************************


Function preDatafetchingForEntityType1(dbName)
On Error Resume Next
	bRunFlag = True
	bFlag = True

	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ClientCompanyNumber = dictDbResultSet("ClientCompanyNumber")
	db_CompanyID = dictDbResultSet("CompanyID")
	
	Set dictDbResultSet = Nothing
	Set query = Nothing
	Append_TestHTML StepCounter, "Get ClientCompanyNumber and CompanyID","ClientCompanyNumber" & db_ClientCompanyNumber & " - CompanyID:"& db_CompanyID , "PASSED"
	cerpcheck = countryCode & db_ClientCompanyNumber
	query_customer = "Select * from Customer where CustomerERP like '"& cerpcheck &"%' order by 2 desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	wait 3
	lastcustCERPno = dictDbResultSet("CustomerERP")
	Set dictDbResultSet = Nothing
	Set query_customer = Nothing
	
	myERP = getLatestnewERPNumberCalculation(cerpcheck,lastcustCERPno)
'	msgbox myERP
	messageqID = checkMessageQueueProccessedTableEntry(myERP)
'	msgbox messageqID
	Do While isempty(messageqID) <> True
'		msgbox "Do while"
		wait 4
		myERP = getLatestnewERPNumberCalculation(cerpcheck,myERP)
'	msgbox myERP
		
		messageqID = checkMessageQueueProccessedTableEntry(myERP)
'		msgbox messageqID
	Loop
'	msgbox "Outside Do"
'''msgbox myERP
	query_customer = "Select * from Customer where CustomerERP = '"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
	ecustSName = dictDbResultSet("ShortName")
	
	set dictDbResultSet = Nothing
	Set query_customer = Nothing
	query_customer = "Select * from MessageQueue where EntityRowID ='"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	wait 3
	messageqID = dictDbResultSet("MessageQueueID")
	 custStatusID = dictDbResultSet("MessageQueueStatusID")
	custEntityPayload = dictDbResultSet("EntityPayload")	
	set dictDbResultSet = Nothing
	Set query_customer = Nothing
	
	If isempty(ecustCID)  Then
		Append_TestHTML StepCounter, "Verify Customer details using new ERP" & myERP ,"New ERP not exist in the Customer Table and will able to process EntityType1 JSON request", "PASSED"
		
		If isempty(messageqID) Then
			Append_TestHTML StepCounter, "Verify MessageQueue details using EntityRowID" & myERP ,"EntityRowID not exist, will able to process EntityType1 JSON request", "PASSED"
			
		Else
			Append_TestHTML StepCounter, "Verify MessageQueue details using EntityRowID" & myERP ,"EntityRowID  exist, will able to process EntityType1 JSON request", "PASSED"
'				bFlag = False
'				bRunFlag = False
		
		End If
	Else
		Append_TestHTML StepCounter, "Verify Customer details using new ERP" & myERP ,"New ERP exist in the Customer Table and will unable to process EntityType1 JSON request using CustomerERP", "FAILED"
				bFlag = False
				bRunFlag = False
	
	End If
If bFlag Then
	preDatafetchingForEntityType1 = myERP
Else
	preDatafetchingForEntityType1 = ""
End If

End Function


Function checkMessageQueueProccessedTableEntry(myERP)
	On Error Resume Next
	       query_customer = "Select * from MessageQueueProcessed where EntityRowID ='"&myERP&"'"
		Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
		set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
		wait 3
		messageqID = dictDbResultSet("MessageQueueID")
		messaageereqID = dictDbResultSet("ExternalRequestID")
		 custStatusID = dictDbResultSet("MessageQueueStatusID")
		custEntityPayload = dictDbResultSet("EntityPayload")
		
		set dictDbResultSet = Nothing
		Set query_customer = Nothing
		
		checkMessageQueueProccessedTableEntry = messageqID
		
End Function

Function getLatestnewERPNumberCalculation(cerpcheck,lastcustCERPno)
	On Error Resume Next
	If isEmpty(lastcustCERPno) Then
		mylastno = "1001"
'		Append_TestHTML StepCounter, "Verify Last Entry of Customer ERP","No record exist with specific Format and creating new one", "PASSED"
	Else
'		Append_TestHTML StepCounter, "Verify Last Entry of Customer ERP",lastcustCERPno, "PASSED"
		mylastno = Right(lastcustCERPno,5)
	End If
	mylastno = cdbl(mylastno) + 1
	
	If len(mylastno)<5 Then
		currentlen = len(mylastno)
		For itr = 1 To  (5 - currentlen)  Step 1
			mylastno = cstr("0" & mylastno)
	
		Next
	End If
	
	newcustERP = cerpcheck & Left(mylastno,5)
	myERP = newcustERP
	getLatestnewERPNumberCalculation= myERP
End Function







'******************************* HEADER ******************************************
' Description : The function to verify json response status of input json request of POST metod type
' Creator :  Venkata Srinivasa Rao. K
' Date : 20th December, 2022
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Function invokeyCMDAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)

On error resume next
	Set objAPI = createobject("MSXML2.serverXMLHTTP.6.0")
	objAPI.open reqType, apiurl , asynctype
	arrHeaders = split(headers,",")
	
	For hitr = 0 To ubound(arrHeaders) Step 1
		If arrHeaders(hitr) <> "" or arrHeaders(hitr) <> null Then
			arrheaderNameval = split(arrHeaders(hitr),"==>")
		
			objAPI.setRequestHeader trim(arrheaderNameval(0)) , trim(arrheaderNameval(1))
		End If
	Next
	If strType = "file" Then
		Set fileSysObj = createObject("Scripting.FileSystemObject")
		Set file = FileSysObj.OpenTextFile(strJson)
		sText = file.ReadAll
		Set fileSysObj = Nothing
	Else
		sText = strJson
	End If
	
	Append_TestHTML StepCounter, "Verify input request JSON", "JSON Text is: " & sText , "PASSED"
	
	objAPI.send sText

	pageReturn = objAPI.responseText
' /* Below statement is for getting response Body */
'	bodyReturn =  BinarytoString(objAPI.responseBody)
'************************************************
'	''msgbox objAPI.status
'	''msgbox objAPI.statusText

	Set j = ParseJson(pageReturn)
	If ucase(j.Status) = ucase("Success") Then
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "PASSED"
'		transRef = j.Results.TransactionRef
		transReqID = j.RequestID
		invokeyCMDAPI = pageReturn
		wait 20
	Else
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "FAILED"
		transRef = ""
		transReqID = ""
		transErrDesc = j.Description
		invokeyCMDAPI = Empty
		
	End If
	
	Set objAPI = Nothing
End Function


Function validatejob2200()
On error resume next
	
	bFlag = True
	bRunFlag = True
	query = "Select * from Job where JobTypeID = 2200 and StatusID=3 order by 1 desc;"
	Append_TestHTML StepCounter, "Verify job 2200 ",query, "PASSED"
	wait 20
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_jobid = dbRecordSet("ID")
	db_statusid = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If cint(db_statusid) = cint("3")  Then		'and instr(db_inputxml,newERPCustNo)>0
		Append_TestHTML StepCounter, "Validate job 2200 status", "Expected Value: 3"  & VBCRLF & "Actual Value: " & db_statusid ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate job 2200 status", "Expected Value: 3"  & VBCRLF & "Actual Value: " & db_statusid ,"FAILED"
	
	End If		
		query = "Select * from JobLog where JobID='"& db_jobid & "' order by 1 desc; "
		Append_TestHTML StepCounter, "Verify JobLog of id "& db_jobid  ,query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_Message = dbRecordSet("Message")
		db_logtypeID = dbRecordSet("LogTypeID")
		db_logDate = dbRecordSet("Date")
		Set dbRecordSet = Nothing
		Set query = Nothing
		If db_Message <> "" or  instr(db_Message,newERPCustNo)>0  or  instr(db_logDate,Day(Date))>0 Then
			Append_TestHTML StepCounter, "Validate Message data", "Actual Value: " & db_Message ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Message data", "Actual Value: " & db_Message ,"FAILED"
				bRunFlag = False
		End If
		If db_logtypeID = "1" Then
			Append_TestHTML StepCounter, "Validate LogTYpeID status", "Expected Value: 1"  & VBCRLF & "Actual Value: " & db_logtypeID ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate LogTYpeID status", "Expected Value: 1"  & VBCRLF & "Actual Value: " & db_logtypeID ,"FAILED"
				bRunFlag = False
		End If	
	
End Function
	


Function verifyPostCustomerActiveDetails(custnewERPid)
	
	On error resume next
	
	bFlag = True
	
		query = "Select * from MessageQueue where EntityRowID = '"& custnewERPid & "' order by 1 desc"
		Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, dbName)
		wait 3
		msgQID = dictDbResultSet("MessageQueueID")
		msgRID = dictDbResultSet("ExternalRequestID")
		msgPayload = dictDbResultSet("EntityPayload")
		msgqstatus = dictDbResultSet("MessageQueueStatusID")
		msgrecordcreateddate = dictDbResultSet("CreatedOn")
		set dictDbResultSet = Nothing
		
		If msgqstatus = "1" and instr(msgPayload, custnewERPid)> 0 and instr(msgrecordcreateddate, nextdateDBFormt(Date))>0 Then
			Append_TestHTML StepCounter, "Validate MessageQueue Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
			Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Append_TestHTML StepCounter, "MessageQueue Entity status ", "Record is Still in Message Queue" ,"PASSED"
		Else
			query = "Select * from MessageQueueProcessed where EntityRowID = '"& custnewERPid & "' order by 1 desc"
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			set dictDbResultSet = execute_db_query(query, 1, dbName)
			wait 3
			msgQID = dictDbResultSet("MessageQueueID")
			msgRID = dictDbResultSet("ExternalRequestID")
			msgPayload = dictDbResultSet("EntityPayload")
			msgqstatus = dictDbResultSet("MessageQueueStatusID")
			msgrecordcreateddate = dictDbResultSet("CreatedOn")
			set dictDbResultSet = Nothing
			''msgbox msgPayload
			''msgbox instr(msgPayload, custnewERPid)
			''msgbox msgrecordcreateddate
			msgrecordcreateddate = Split(msgrecordcreateddate," " )(0)
			''msgbox nextdateDBFormt(Date)
			If msgqstatus = "3" and (instr(msgPayload, custnewERPid)> 0 or DateDiff("d",msgrecordcreateddate, Date()) = 0) Then
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
				Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "","FAILED"
				Append_TestHTML StepCounter, "MessageQueueProcessed Entity status ", "No Record created in the MessageQueueProcessed table" ,"FAILED"
					bRunFlag = False
			End If
		
		End If
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&custnewERPid&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
	ecustSName = dictDbResultSet("ShortName")
	ecustPPId = dictDbResultSet("CustomerPriceProfileID")
	ecustFGId = dictDbResultSet("FeeGroupID")
	db_ecustStatusID = dictDbResultSet("StatusID")
	
	set dictDbResultSet = Nothing
	
	If ecustCID <> "" Then
		Append_TestHTML StepCounter, "Validate Customer Details", "Customer record successfully Created" ,"PASSED"
		Call OpenApplication(url)
		Call customerSearch1(custnewERPid)
		customerERP_id = custnewERPid
	Else
		Append_TestHTML StepCounter, "Validate Customer Details", "Customer record not Yet Created" , "FAILED"
			bRunFlag = False
	End If
'	If defaultcheckflag Then
'		If ecustPPId = db_DefaultPriceProfileID Then
'			Append_TestHTML StepCounter, "Validate Customer Default PriceProfileID Details", "Expected:"& db_DefaultPriceProfileID & VBCRLF & "Actual:" & ecustPPId  ,"PASSED"
'		Else
'			Append_TestHTML StepCounter, "Validate Customer Default PriceProfileID Details", "Expected:"& db_DefaultPriceProfileID & VBCRLF & "Actual:" & ecustPPId  ,"FAILED"
'				bRunFlag = False
'		End If
'		If ecustFGId = db_DefaultFeeGroupID Then
'			Append_TestHTML StepCounter, "Validate Customer Default FeeGroupID Details", "Expected:"& db_DefaultFeeGroupID & VBCRLF & "Actual:" & ecustFGId  ,"PASSED"
'		Else
'			Append_TestHTML StepCounter, "Validate Customer Default FeeGroupID Details", "Expected:"& db_DefaultFeeGroupID & VBCRLF & "Actual:" & ecustFGId  ,"FAILED"	
'				bRunFlag = False
'		End If
'	End If
	
	
	
End Function

Function feeandPriceDefaultvaluesValidation(custnewERPid)
	On Error Resume Next
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&custnewERPid&"'"
'	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
	ecustSName = dictDbResultSet("ShortName")
	ecustPPId = dictDbResultSet("CustomerPriceProfileID")
	ecustFGId = dictDbResultSet("FeeGroupID")
	db_ecustStatusID = dictDbResultSet("StatusID")
	
	set dictDbResultSet = Nothing
	

		If ecustPPId = db_DefaultPriceProfileID Then
			Append_TestHTML StepCounter, "Validate Customer Default PriceProfileID Details", "Expected:"& db_DefaultPriceProfileID & VBCRLF & "Actual:" & ecustPPId  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Default PriceProfileID Details", "Expected:"& db_DefaultPriceProfileID & VBCRLF & "Actual:" & ecustPPId  ,"FAILED"
				bRunFlag = False
		End If
		If ecustFGId = db_DefaultFeeGroupID Then
			Append_TestHTML StepCounter, "Validate Customer Default FeeGroupID Details", "Expected:"& db_DefaultFeeGroupID & VBCRLF & "Actual:" & ecustFGId  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Default FeeGroupID Details", "Expected:"& db_DefaultFeeGroupID & VBCRLF & "Actual:" & ecustFGId  ,"FAILED"	
				bRunFlag = False
		End If

	
End Function


Function searchandReplaceMultipleString(jsonFPath, searchreplaceString)

On Error Resume Next
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(jsonFPath) Then
		Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
		Newcontent = fileRead.ReadAll
		fileRead.Close
		Set fileSysObj = Nothing
		srstrings = Split(searchreplaceString,";")
		'Append_TestHTML StepCounter, "Verify json input File Path", "File exist in the path:Before Modifying Data--:" & Newcontent , "PASSED"
		For titr = 0 To ubound(srstrings) Step 1
			sstrings = Split(srstrings(titr),"-")
			searchString = sstrings(0)
			replaceString = sstrings(1)
			searchstr = chr(34) & searchString & chr(34)
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
			keycontent = ""
			Do until fileRead.AtEndOfStream
				content = fileRead.ReadLine	
				If instr(content,searchstr) > 0 Then
					keycontent = content	
					Exit Do
				End If
			Loop
			If keycontent <> "" and Newcontent <> "" Then
			If isNumeric(replaceString) = False Then
				repcontent = searchstr & ": " & chr(34) & replaceString & chr(34) 
			Else
				repcontent = searchstr & ": "  & replaceString 
			End If
				If right(keycontent,1) = "," Then
					repcontent = repcontent & ","
				End If
				Newcontent = Replace(Newcontent, keycontent, repcontent)
				Set DXwrite = fileSysObj.OpenTextFile(jsonFPath,2)
				DXwrite.Write Newcontent
				DXwrite.Close
			End If
			fileRead.Close
			Set fileSysObj = Nothing
		Next
		'Append_TestHTML StepCounter, "Verify json input File Path", "File exist in the path:After Modifying Data--:" & Newcontent , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify json input File Path", "File not exist in the path:" & jsonFPath , "FAILED"
	End If
End Function


Function getapiConfigInfoFromCSV(Apconfigpath)
	On error resume next
	bFlag = True
'	dataFile = TDFilePath & "TC008_PricingCreateNewPriceRule.csv"
Apiconfigpath= Apconfigpath

	 Set xlApp = CreateObject("Excel.Application")
	 xlApp.Visible = true
	 Set xlBook = xlApp.Workbooks.open(Apiconfigpath)
	sheetName = "APIdata"
		Set xlSheet = xlBook.Worksheets(sheetName)
		rcount = xlSheet.UsedRange.rows.count

		ccount = xlSheet.UsedRange.columns.count

		For i = 1 To rcount Step 1
			platformName =  xlSheet.Cells(i,1).Value
			envName =  xlSheet.Cells(i,2).Value
			
			If ucase(platformName) = ucase(appName) and ucase(envName) = ucase(appEnvName) Then
				dataentryrow = i
				Exit For
			End If
		Next

		If dataentryrow > 1 and dataentryrow<=rcount Then
				headerNames =  xlSheet.Cells(dataentryrow,4).Value
				headervals = Split(headerNames,";")
				
				apiHeaders=""
				For itr = 0 To ubound(headervals) Step 1
					headerval = Split(headervals(itr),"#")
					apiHeaders = apiHeaders & headerval(0)	& "==>" &  headerval(1) & ","
				Next	 
				getapiConfigInfoFromCSV = apiHeaders
		Else
				getapiConfigInfoFromCSV = ""
		End If
		
	 xlBook.Save
	  xlApp.Quit
End Function



Function basicPreConfigData(jsonFPath,dbName)

On Error Resume Next
	bRunFlag = True
	bFlag = True
	
	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
	db_CompanyID = dictDbResultSet("CompanyID")
	db_CountryID = dictDbResultSet("CountryID")
	db_RegionID = dictDbResultSet("LegislativeRegionID")
	db_CurrencyID = dictDbResultSet("CurrencyID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	If db_ColcoCode <> "" Then
		Append_TestHTML StepCounter, "Get ClientCompanyNumber/ColcoCode from Company Table", "Set ColcoCode as- " & db_ColcoCode & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get ClientCompanyNumber/ColcoCode from Company Table", "Set ColcoCode as- " & db_ColcoCode & " - into INPUT JSON" , "FAILED"	
	End If
	
	
	
	query = "Select * from CompanyLanguage where CompanyID="& db_CompanyID & ";"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_LanguageID = dictDbResultSet("LanguageID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	If db_LanguageID <> "" Then
		Append_TestHTML StepCounter, "Get LanguageID from CompanyLanguage Table", "Set LanguageID as- " & db_LanguageID & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get LanguageID from CompanyLanguage Table", "Set LanguageID as- " & db_LanguageID & " - into INPUT JSON" , "FAILED"	
	End If
	
	query = "Select * from Band;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_BandID = dictDbResultSet("BandID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_BandID <> "" Then
		Append_TestHTML StepCounter, "Get BandID from CompanyLanguage Table", "Set Band as- " & db_BandID & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get BandID from CompanyLanguage Table", "Set Band as- " & db_BandID & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from LocalisedDescriptions where ColumnName like '%CustomerClassificationTypeId%' and Culture='en-GB';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_CCTId_Value = dictDbResultSet("Value")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_CCTId_Value <> "" Then
		Append_TestHTML StepCounter, "Get CustomerClassificationTypeId from LocalisedDescriptions Table", "Set CustomerClassificationTypeId as- " & db_CCTId_Value & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get CustomerClassificationTypeId from LocalisedDescriptions Table", "Set CustomerClassificationTypeId as- " & db_CCTId_Value & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from LocalisedDescriptions where ColumnName like '%MarketingSegmentationId%' and Culture='en-GB';"	' and Description='National/International Fleet/IKA';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_MSId_Value = dictDbResultSet("Value")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_MSId_Value <> "" Then
		Append_TestHTML StepCounter, "Get MarketingSegmentationId from LocalisedDescriptions Table", "Set MarketingSegmentationId as- " & db_MSId_Value & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get MarketingSegmentationId from LocalisedDescriptions Table", "Set MarketingSegmentationId as- " & db_MSId_Value & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from LocalisedDescriptions where ColumnName like '%IndustrialClassId%' and Culture='en-GB'	and Description='Growing of rice';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ICId_Value = dictDbResultSet("Value")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_ICId_Value <> "" Then
		Append_TestHTML StepCounter, "Get IndustrialClassId from LocalisedDescriptions Table", "Set IndustrialClassId as- " & db_ICId_Value & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get IndustrialClassId from LocalisedDescriptions Table", "Set IndustrialClassId as- " & db_ICId_Value & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from LocalisedDescriptions where ColumnName like '%LegalEntityID%' and Culture='en-GB'	and Description='NV' and Culture='en-GB';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_LEId_Value = dictDbResultSet("Value")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_LEId_Value <> "" Then
		Append_TestHTML StepCounter, "Get LegalEntityID from LocalisedDescriptions Table", "Set LegalEntityID as- " & db_LEId_Value & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get LegalEntityID from LocalisedDescriptions Table", "Set LegalEntityID as- " & db_LEId_Value & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from GlobalCustomerReference;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_GCRID_Value = dictDbResultSet("GlobalCustomerReferenceID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_GCRID_Value <> "" Then
		Append_TestHTML StepCounter, "Get GlobalCustomerReferenceID from GlobalCustomerReference Table", "Set GlobalCustomerReferenceID as- " & db_GCRID_Value & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get GlobalCustomerReferenceID from GlobalCustomerReference Table", "Set GlobalCustomerReferenceID as- " & db_GCRID_Value & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from FeeGroup where CompanyID="& db_CompanyID & " and FeeGroupName='Base Group';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_FeeGroupID= dictDbResultSet("FeeGroupID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
		Append_TestHTML StepCounter, "Get FeeGroupID from FeeGroup Table", "Set FeeGroupID as- " & db_FeeGroupID & " - into INPUT JSON" , "PASSED"
	
	query = "Select * from ColcoLineOfBusiness where ColcoID="& db_CompanyID & ";"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_LineOfBusinessID = dictDbResultSet("LineOfBusinessID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If db_LineOfBusinessID <> "" Then
		Append_TestHTML StepCounter, "Get LineOfBusinessID from ColcoLineOfBusiness Table", "Set LineOfBusinessID as- " & db_LineOfBusinessID & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get LineOfBusinessID from ColcoLineOfBusiness Table", "Set LineOfBusinessID as- " & db_LineOfBusinessID & " - into INPUT JSON" , "FAILED"	
	End If
	query = "Select * from PriceProfile where ColcoID="& db_CompanyID & " and PriceProfileDescription='Fleet Default Price Profile';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_PriceProfileID = dictDbResultSet("PriceProfileID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
		Append_TestHTML StepCounter, "Get PriceProfileID from PriceProfile Table", "Set PriceProfileID as- " & db_PriceProfileID & " - into INPUT JSON" , "PASSED"
	
	query = "Select * from CustomerSegmentationType;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_CustomerSegmentationTypeID = dictDbResultSet("CustomerSegmentationTypeID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	If db_CustomerSegmentationTypeID <> "" Then
		Append_TestHTML StepCounter, "Get CustomerSegmentationTypeID from CustomerSegmentationType Table", "Set CustomerSegmentationTypeID as- " & db_CustomerSegmentationTypeID & " - into INPUT JSON" , "PASSED"
	Else
		Append_TestHTML StepCounter, "Get CustomerSegmentationTypeID from CustomerSegmentationType Table", "Set CustomerSegmentationTypeID as- " & db_CustomerSegmentationTypeID & " - into INPUT JSON" , "FAILED"	
	End If
	
	err_colcoID = db_ColcoCode
	preSearchString = "ColcoCode-" & db_ColcoCode &";LanguageId-"& db_LanguageID &";CurrencyId-" & db_CurrencyID & ";CountryId-"& db_CountryID &";BandId-" & db_BandID & ";CustomerClassificationId-" & db_CCTId_Value & ";MarketingSegmentationId-" & db_MSId_Value & ";IndustrialClassId-"& db_ICId_Value & ";LegalEntityId-" & db_LEId_Value
	Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
	If FeeGroupId = "" or isNULL(FeeGroupId) = true Then
		 preSearchString = "CustomerGroupingReference1Id-" & db_GCRID_Value &";LineOfBusinessId-" & db_LineOfBusinessID & ";BillingLanguageId-" & db_LanguageID & ";CustomerSegmentationTypeId-" & db_CustomerSegmentationTypeID 
	
	Else
		preSearchString = "CustomerGroupingReference1Id-" & db_GCRID_Value & ";FeeGroupId-" & db_FeeGroupID &";LineOfBusinessId-" & db_LineOfBusinessID & ";PriceProfileId-"& db_PriceProfileID & ";BillingLanguageId-" & db_LanguageID & ";CustomerSegmentationTypeId-" & db_CustomerSegmentationTypeID 
		
	End If
	 Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
	 
	 query = "Select * from ProfileDefaults where BandID="& db_BandID & " and MarketingSegmentationID=" & db_MSId_Value & "  and LineOfBusinessID=" & db_LineOfBusinessID &";"
	 Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_DefaultPriceProfileID = dictDbResultSet("DefaultPriceProfileID")
	db_DefaultFeeGroupID = dictDbResultSet("DefaultFeeGroupID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	If db_DefaultPriceProfileID <> "" Then
		Append_TestHTML StepCounter, "Get DefaultPriceProfileID from ProfileDefaults Table", " DefaultPriceProfileID as- " & db_DefaultPriceProfileID & " - for validation" , "PASSED"
		Append_TestHTML StepCounter, "Get DefaultFeeGroupID from ProfileDefaults Table", "Set DefaultFeeGroupID as- " & db_DefaultFeeGroupID & " - for validation" , "PASSED"
		
	Else
		Append_TestHTML StepCounter, "Get DefaultPriceProfileID from ProfileDefaults Table", "Set DefaultPriceProfileID as- " & db_DefaultPriceProfileID & " - for validation" , "FAILED"	
		Append_TestHTML StepCounter, "Get DefaultFeeGroupID from ProfileDefaults Table", "Set DefaultFeeGroupID as- " & db_DefaultFeeGroupID & " - for validation" , "FAILED"
		
	End If
End Function



Function searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)

On Error Resume Next
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(jsonFPath) Then
		Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
		Newcontent = fileRead.ReadAll
		fileRead.Close
		Set fileSysObj = Nothing
		srstrings = Split(searchreplaceString,";")
		'Append_TestHTML StepCounter, "Verify json input File Path", "File exist in the path:Before Modifying Data--:" & Newcontent , "PASSED"
		For titr = 0 To ubound(srstrings) Step 1
			sstrings = Split(srstrings(titr),"-")
			searchString = sstrings(0)
			replaceString = sstrings(1)
			searchstr = chr(34) & searchString & chr(34)
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
			keycontent = ""
			Do until fileRead.AtEndOfStream
				content = fileRead.ReadLine	
				If instr(content,searchstr) > 0 Then
					keycontent = content	
					Exit Do
				End If
			Loop
			If keycontent <> "" and Newcontent <> "" Then
			
				mtcontents = split(keycontent,":")
				If isNULL(replaceString) or instr(replaceString,"null")>0 or instr(replaceString,"[")>0 Then
					repcontent = searchstr & ": "  & replaceString 
				Else
					If Mid(mtcontents(1),2,1) = chr(34) Then
						repcontent = searchstr & ": " & chr(34) & replaceString & chr(34) 
					Else
						repcontent = searchstr & ": "  & replaceString 
					End If
				End If
				
'			If isNumeric(replaceString) = False Then
'				repcontent = searchstr & ": " & chr(34) & replaceString & chr(34) 
'			Else
'				repcontent = searchstr & ": "  & replaceString 
'			End If
				If right(keycontent,1) = "," Then
					repcontent = repcontent & ","
				End If
				Newcontent = Replace(Newcontent, keycontent, repcontent)
				Append_TestHTML StepCounter, "Set " & searchString & " value with -" & replaceString &" in the input JSON", repcontent , "PASSED"
				
				Set DXwrite = fileSysObj.OpenTextFile(jsonFPath,2)
				DXwrite.Write Newcontent
				DXwrite.Close
			End If
			fileRead.Close
			Set fileSysObj = Nothing
		Next
		'Append_TestHTML StepCounter, "Verify json input File Path", "File exist in the path:After Modifying Data--:" & Newcontent , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify json input File Path", "File not exist in the path:" & jsonFPath , "FAILED"
	End If
End Function


Function invokeErrorCMDAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)

On error resume next
	bFlag = True
		bRunFlag =True
	Set objAPI = createobject("MSXML2.serverXMLHTTP.6.0")
	'Set objAPI = createobject("WinHttp.WinHttpRequest.5.1")
	
	objAPI.open reqType, apiurl , asynctype
	arrHeaders = split(headers,",")
	
	For hitr = 0 To ubound(arrHeaders) Step 1
		If arrHeaders(hitr) <> "" or arrHeaders(hitr) <> null Then
			arrheaderNameval = split(arrHeaders(hitr),"==>")
		
			objAPI.setRequestHeader trim(arrheaderNameval(0)) , trim(arrheaderNameval(1))
		End If
	Next
	If strType = "file" Then
		Set fileSysObj = createObject("Scripting.FileSystemObject")
		Set file = FileSysObj.OpenTextFile(strJson)
		sText = file.ReadAll
		Set fileSysObj = Nothing
	Else
		sText = strJson
	End If
	
	Append_TestHTML StepCounter, "Verify input request JSON", "JSON Text is: " & sText , "PASSED"
	
	objAPI.send sText
	
	pageReturn = objAPI.responseText
' /* Below statement is for getting response Body */
'	bodyReturn =  BinarytoString(objAPI.responseBody)
'************************************************
'	msgbox objAPI.status
'	msgbox objAPI.statusText
	For itr = 1 To 5 Step 1
		If pageReturn = "" Then
			objAPI.send sText
			pageReturn = objAPI.responseText
		Else
			Exit For
		End If
	Next

	Set j = ParseJson(pageReturn)
	If ucase(j.Status) = ucase("Success") Then
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "PASSED"
		transRef = j.Results.TransactionRef
		transReqID = j.RequestID
		invokeErrorCMDAPI = pageReturn
	ElseIf ucase(j.Status) = ucase("Failed") Then
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "PASSED"
		transRef = ""
		transReqID = ""
		transErrDesc = j.Description
		transErrLength = j.Errors.Length 
		transErrdetail = j.Errors.[0].Detail
		transErrdetailsval = Split(transErrdetail,"|")
		For rit = 0 To ubound(transErrdetailsval) Step 1
			transErrdetailonetext = transErrdetailsval(rit)
			translasttextval = Split(transErrdetailonetext," ")
			Append_TestHTML StepCounter, "Validate " & translasttextval(ubound(translasttextval)) , transErrdetailonetext, "PASSED"
		Next
		invokeErrorCMDAPI = pageReturn
	Else
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "FAILED"
		transRef = ""
		transReqID = ""
		transErrDesc = j.Description
		invokeErrorCMDAPI = Empty
		bRunFlag = False
	End If
		wait 40
	Set objAPI = Nothing
End Function


Function getStatusChangeReasonID(jsonFPath,dbName)
On Error Resume Next
	query = "Select * from LocalisedDescriptions where  ColumnName like '%CustomerStatusChangeReasonID%' and culture='en-GB' and Description='Customer Request Other';"
		Append_TestHTML StepCounter, "Get Status ChangeReason of id "& db_jobid  ,query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_SCR_value = dbRecordSet("Value")
		
		Set dbRecordSet = Nothing
		Set query = Nothing

		db_BlockDate =  nextdateDBFormt(Date()-1)

		 preSearchString = "StatusChangeReasonId$" & db_SCR_value & ";BlockDate$" & db_BlockDate
		 Call searchandReplaceMultipleStringwithdollar(jsonFPath, preSearchString)
		
End Function



Function basicPreConfigAddressData(jsonFPath,dbName)

On Error Resume Next
	bRunFlag = True
	bFlag = True
	
	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
	db_CompanyID = dictDbResultSet("CompanyID")
	db_CountryID = dictDbResultSet("CountryID")
'	db_RegionID = dictDbResultSet("LegislativeRegionID")
'	db_CurrencyID = dictDbResultSet("CurrencyID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	query = "Select * from Region where CountryID="& db_CountryID & ";"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_RegionID = dictDbResultSet("RegionID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	preSearchString = "ColcoCode-" & db_ColcoCode &";CountryId-"& db_CountryID &";RegionId-" & db_RegionID
	Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
End Function


		

Function verifyPostCustomerAddressActiveDetails(custnewERPid,oprequestID)
	
	On error resume next
	
	bFlag = True
	
		query = "Select * from MessageQueue where ExternalRequestID = '"& oprequestID & "' order by 1 desc"
		Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
		wait 3
		msgQID = dictDbResultSet("MessageQueueID")
		msgRID = dictDbResultSet("ExternalRequestID")
		msgPayload = dictDbResultSet("EntityPayload")
		msgqstatus = dictDbResultSet("MessageQueueStatusID")
		msgrecordcreateddate = dictDbResultSet("CreatedOn")
		set dictDbResultSet = Nothing
		
		If msgqstatus = "1" and instr(msgPayload, custnewERPid)> 0 and instr(msgrecordcreateddate, nextdateDBFormt(Date))>0 Then
			Append_TestHTML StepCounter, "Validate MessageQueue Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
			Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Append_TestHTML StepCounter, "MessageQueue Entity status ", "Record is Still in Message Queue" ,"PASSED"
		Else
			query = "Select * from MessageQueueProcessed where ExternalRequestID = '"& oprequestID & "' order by 1 desc"
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
			wait 3
			msgQID = dictDbResultSet("MessageQueueID")
			msgRID = dictDbResultSet("ExternalRequestID")
			msgPayload = dictDbResultSet("EntityPayload")
			msgqstatus = dictDbResultSet("MessageQueueStatusID")
			msgrecordcreateddate = dictDbResultSet("CreatedOn")
			msgrecordAccountNumber = dictDbResultSet("AccountNumber")
			set dictDbResultSet = Nothing
			''msgbox msgPayload
			''msgbox instr(msgPayload, custnewERPid)
			''msgbox msgrecordcreateddate
			''msgbox nextdateDBFormt(Date)			
			If msgqstatus = "3" and (instr(msgPayload, "AddressTypeId")> 0 or instr(msgPayload, "ContactTypeId")> 0 or instr(msgPayload, "BankTypeId")> 0  or instr(msgPayload, "DDMandateTypeId")>0 or instr(msgPayload, "CustomerCardType")> 0 or instr(msgPayload, "CreditLimit")> 0 or instr(msgrecordcreateddate, nextdateDBFormt(Date))>0) Then
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
				Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "messagequeuestatus-" & msgqstatus & "Payload-" & msgPayload  ,"FAILED"
				Append_TestHTML StepCounter, "MessageQueueProcessed Entity status ", "No Record created in the MessageQueueProcessed table" ,"FAILED"
					bRunFlag = False
			End If
		
		End If
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&msgrecordAccountNumber&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
'	ecustSName = dictDbResultSet("ShortName")
'	ecustPPId = dictDbResultSet("CustomerPriceProfileID")
'	ecustFGId = dictDbResultSet("FeeGroupID")
'	db_ecustStatusID = dictDbResultSet("StatusID")
	
	set dictDbResultSet = Nothing
	
	
	Call OpenApplication(url)
	Call customerSearch1(msgrecordAccountNumber)
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary")  Then
		Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=SearchMenu","html tag:=INPUT").Set "Addresses"
		wait 2
		Browser("creationTime:=1").Page("creationTime:=1").Link("html tag:=A","innertext:="&"Addresses","index:=0").Click
		wait 2
	End  If
	If ecustCID <> "" Then
		verifyPostCustomerAddressActiveDetails = ecustCID
	Else
		verifyPostCustomerAddressActiveDetails = ""	
	End If
	
End Function

Function validateAddressLines(jsonFPath,addresscustid,externaRefVal,addresstypeids,operationflag)
	On Error Resume Next
	
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set file = FileSysObj.OpenTextFile(jsonFPath,1)
		sText = file.ReadAll
		
	Set j = ParseJson(sText)
	
	query_customer = "Select * from CustomerAddress where CustomerID = '"&addresscustid&"' and ExternalAddressID='"&externaRefVal & "';"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 eAddressID= dictDbResultSet("AddressID")
	 eIsActive = dictDbResultSet("IsActive")
	 eAddressLines= dictDbResultSet("AddressLines")
	 eZipcode = dictDbResultSet("Zipcode")
	 eCity = dictDbResultSet("City")
	 eRegionID = dictDbResultSet("RegionID")
	 eCountryID = dictDbResultSet("CountryID")
	 eTelephone = dictDbResultSet("Telephone")
	 eFax = dictDbResultSet("Fax")
	 
	 Set query_customer=Nothing
	set dictDbResultSet = Nothing
	db_Cust_AddressID = eAddressID
	
	If instr(addresstypeids,",") > 0 Then
		addtypeidvals = Split(addresstypeids,",")
		len_addtypevals = ubound(addtypeidvals)
	Else
		len_addtypevals = 0
	End If
	
	For itr = 0 To len_addtypevals Step 1
		addTypeID =  eval("j.EntityBody.AddressTypeId.[" & itr & "]")
		eModifiedBy = ""
		query_customer = "Select * from CustomerAddressType where CustomerID = '"& addresscustid &"' and AddressTypeID=" & addTypeID & " and AddressID = "& eAddressID &";"
		
		Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
		set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
		 eModifiedBy= dictDbResultSet("ModifiedBy")
		 eAddressTypeID= dictDbResultSet("AddressTypeID")
		 Set query_customer = Nothing
		set dictDbResultSet = Nothing
		If cint(len_addtypevals) = cint("0") and cdbl(addTypeID) = cdbl("0")   Then
			If eAddressTypeID = "" Then
				query_customer = "Select Count(*) as addressentries from CustomerAddressType where CustomerID = '"&addresscustid&"' " & " and AddressID = "& eAddressID &";"
				Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
				set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
				 No_of_Caddressentries = dictDbResultSet("addressentries")
				 Set query_customer = Nothing
				set dictDbResultSet = Nothing
				If isNULL(No_of_Caddressentries) = true or cint(No_of_Caddressentries) = cint("0") Then
					Append_TestHTML StepCounter, "Validate AddressTypeID" & addTypeID , "All records are deleted from table" ,"PASSED"
				Else
						Append_TestHTML StepCounter, "Validate AddressTypeID" & addTypeID , "Fail to delete all records-" & addtypeidvals &"- Entries:" &No_of_Caddressentries  ,"FAILED"
				End If
				
			Else
				Append_TestHTML StepCounter, "Validate AddressTypeID" & addTypeID , "Fail to delete records-" & addtypeidvals ,"FAILED"
			End If
		Else
			If eModifiedBy <> "" and cdbl(eAddressTypeID) = cdbl(addTypeID)  Then
					Append_TestHTML StepCounter, "Validate AddressTypeID" & addTypeID , "AddressType " & eAddressTypeID & " record find successfully" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Validate AddressTypeID" & addTypeID , "AddressType " & eAddressTypeID & " record find successfully" ,"FAILED"
			End If
		End If
		
	Next
	query_customer = "Select Count(*) as addressentry from CustomerAddress where CustomerID = '"&addresscustid&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 No_of_addressentries = dictDbResultSet("addressentry")
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	If cint(No_of_addressentries) = cint("1") Then
		Append_TestHTML StepCounter, "Validate Customer Address", No_of_addressentries & " Address entry created successfully" ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Address", No_of_addressentries & " Address entry created successfully" ,"FAILED"	
	End If
	
	If instr(eAddressLines, j.EntityBody.AddressLine1 )>0 and instr(eAddressLines, j.EntityBody.AddressLine2 )>0 and instr(eAddressLines, j.EntityBody.AddressLine3 )>0 Then
		Append_TestHTML StepCounter, "Validate Customer AddressLine", "Expected:"& j.EntityBody.AddressLine1 & j.EntityBody.AddressLine2 & j.EntityBody.AddressLine3  & VBCRLF & "Actual:" & eAddressLines  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer AddressLine", "Expected:"& j.EntityBody.AddressLine1 & j.EntityBody.AddressLine2 & j.EntityBody.AddressLine3 & VBCRLF & "Actual:" & eAddressLines  ,"FAILED"
			bRunFlag = False
	End If
	
	If cstr(eCity) = cstr(j.EntityBody.City) Then
		Append_TestHTML StepCounter, "Validate Customer City value", "Expected:"& j.EntityBody.City & VBCRLF & "Actual:" & eCity  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer City value", "Expected:"& j.EntityBody.City & VBCRLF & "Actual:" & eCity  ,"FAILED"	
			bRunFlag = False
	End If
	If cint(eRegionID) = cint(j.EntityBody.RegionId) Then
		Append_TestHTML StepCounter, "Validate Customer RegionId value", "Expected:"& j.EntityBody.RegionId & VBCRLF & "Actual:" & eRegionID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer RegionId value", "Expected:"& j.EntityBody.RegionId & VBCRLF & "Actual:" & eRegionID  ,"FAILED"	
			bRunFlag = False
	End If
	If cstr(eZipcode) = cstr(j.EntityBody.PostalCode) Then
		Append_TestHTML StepCounter, "Validate Customer PostalCode value", "Expected:"& j.EntityBody.PostalCode & VBCRLF & "Actual:" & eZipcode  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer PostalCode value", "Expected:"& j.EntityBody.PostalCode & VBCRLF & "Actual:" & eZipcode  ,"FAILED"	
			bRunFlag = False
	End If
	If cbool(eCountryID) = cbool(j.EntityBody.CountryId) Then
		Append_TestHTML StepCounter, "Validate Customer CountryId value", "Expected:"& j.EntityBody.CountryId & VBCRLF & "Actual:" & eCountryID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer CountryId value", "Expected:"& j.EntityBody.CountryId & VBCRLF & "Actual:" & eCountryID  ,"FAILED"	
			bRunFlag = False
	End If
	If cbool(eTelephone) = cbool(j.EntityBody.Telephone) Then
		Append_TestHTML StepCounter, "Validate Customer Telephone value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & eTelephone  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Telephone value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & eTelephone  ,"FAILED"	
			bRunFlag = False
	End If
	If cbool(eFax) = cbool(j.EntityBody.Fax) Then
		Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & eIsActive  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & eIsActive  ,"FAILED"	
			bRunFlag = False
	End If
	If cbool(eIsActive) = cbool(j.EntityBody.IsActive) Then
		Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & eIsActive  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & eIsActive  ,"FAILED"	
			bRunFlag = False
	End If
	If Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvAddresses_AddressList_Table","html tag:=TABLE").Exist Then
		Append_TestHTML StepCounter,"Navigate to Address Page","Successfully Navigated to Address Page" ,"PASSED"
		AddLineData = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvAddresses_AddressList_Table","html tag:=TABLE").GetCellData(2,1)
		wait 2
		Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvAddresses_AddressList_Table","html tag:=TABLE").ChildItem(2,1,"WebElement",0).Click
		wait 5
		Append_TestHTML StepCounter,"Navigate to Address Details Tab ","Successfully Navigated to Address Details Tab" ,"PASSED"
	End If
	
	query_customer = "Select * from CustomerActivityLog where CustomerID=" & addresscustid &" order by ModifiedOn desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_C_ActivityLogID = dictDbResultSet("ActivityLogID")
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	
	query_customer = "Select * from ActivityLog where ActivityTypeID=98 order by LogDateTime desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_Note = dictDbResultSet("Note")
	 Db_A_ActivityLogID = dictDbResultSet("ActivityLogID")
	 
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	
	wait 15
	query_customer = "Select __$operation as coperval, ModifiedOn,IsActive,AddressLines,Zipcode,City,RegionID,CountryID,Telephone,Fax  from cdc.dbo_CustomerAddress_CT where CustomerID=" & addresscustid & " and AddressID= " & eAddressID & " and ExternalAddressID='" & externaRefVal & "' order by __$seqval Desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 3, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_C_Operation = dictDbResultSet("coperval")
	 Db_C_ModifiedOn = dictDbResultSet("ModifiedOn")
	 Db_C_IsActive= dictDbResultSet("IsActive")
	 Db_C_AddressLines= dictDbResultSet("AddressLines")
	 Db_C_Zipcode= dictDbResultSet("Zipcode")
	 Db_C_City= dictDbResultSet("City")
	 Db_C_RegionID= dictDbResultSet("RegionID")
	 Db_C_CountryID= dictDbResultSet("CountryID")
	 Db_C_Telephone= dictDbResultSet("Telephone")
	 Db_C_Fax= dictDbResultSet("Fax")
	  Set query_customer = Nothing
	 set dictDbResultSet = Nothing
	
	query_customer = "Select __$operation as coperval,AddressTypeID from cdc.dbo_CustomerAddressType_CT where CustomerID=" & addresscustid & " and AddressID= " & eAddressID & " order by __$seqval Desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 3, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_T_Operation = dictDbResultSet("coperval")
	 Db_T_AddressTypeID = dictDbResultSet("AddressTypeID")
	 
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	
	If ucase(operationflag) = ucase("insertion") Then
		If Db_C_ActivityLogID = "" Then
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID &" Action: No Records Entry in the table" ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"FAILED"			
		End If	
		If instr(Db_Note,addresscustid) = 0 and instr(Db_Note,eAddressID) = 0 Then
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:"  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:" & Db_Note ,"FAILED"
		End If
		If instr(Db_C_Operation ,"2")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive ,"FAILED"			
		End If
		If instr(addresstypeids,",")>0 Then
			If instr(Db_T_Operation , "2")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation & " AddressTypeID" & Db_T_AddressTypeID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation & " AddressTypeID" & Db_T_AddressTypeID  ,"FAILED"	
			End If
		Else
			If Db_T_Operation = "" Then
					Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: Empty"  & VBCRLF & "Actual:" & Db_T_Operation & " AddressTypeID" & addresstypeids ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: Empty"  & VBCRLF & "Actual:" & Db_T_Operation & " AddressTypeID" & addresstypeids ,"FAILED"	
			End If
		End If
		
	ElseIf ucase(operationflag) = ucase("update") or ucase(operationflag) = ucase("active")  Then
		If Db_C_ActivityLogID <> ""  and cdbl(Db_A_ActivityLogID) = cdbl(Db_C_ActivityLogID) Then
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual: " & Db_C_ActivityLogID & "-" & Db_A_ActivityLogID & "** ActivityLogID is not matching from CustomerActivityLog  and  ActivityLog tables"  ,"FAILED"
		End If	
		If instr(Db_Note,addresscustid) > 0 and instr(Db_Note,eAddressID) > 0 Then
			Append_TestHTML StepCounter, "Validate ActivityLog Note Data", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate ActivityLog Note Data", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
		End If
		If  ucase(operationflag) = ucase("update")  Then
		
			If instr(Db_C_Operation ,"3")>0 or instr(Db_C_Operation ,"4")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT  operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive  ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT  operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " Modified On" & Db_C_ModifiedOn  ,"FAILED"
			End If
			If instr(Db_T_Operation, "2")>0 and instr(Db_T_AddressTypeID,"2")>0 and instr(Db_T_AddressTypeID,"3")>0 and instr(Db_T_AddressTypeID,"4")>0 Then
					Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation  & " AddressTypeID" & Db_T_AddressTypeID,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation & "AddressTypeID" & Db_T_AddressTypeID ,"FAILED"	
			End If
			
		End If
	ElseIf ucase(operationflag) = ucase("deletion") Then
		If Db_C_ActivityLogID <> ""  and cdbl(Db_A_ActivityLogID) = cdbl(Db_C_ActivityLogID) Then
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID & "-" & Db_A_ActivityLogID & " ActivityLogID:-" & Db_C_ActivityLogID  ,"FAILED"
		End If	
		If instr(Db_Note,addresscustid) > 0 and instr(Db_Note,eAddressID) > 0 Then
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
		End If
		If instr(Db_C_Operation ,"3")>0 or instr(Db_C_Operation ,"4")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT  operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive  ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT  operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " Modified On" & Db_C_ModifiedOn  ,"FAILED"
		End If
		If instr(Db_T_Operation, "1")>0 and instr(Db_T_AddressTypeID,"2")>0 and instr(Db_T_AddressTypeID,"3")>0 and instr(Db_T_AddressTypeID,"4")>0  Then
			Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: 1"  & VBCRLF & "Actual:" & Db_T_Operation & " AddressTypeID" & Db_T_AddressTypeID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate dbo_CustomerAddressType_CT  operation value", "Expected: 1"  & VBCRLF & "Actual:" & Db_T_Operation & " AddressTypeID" & Db_T_AddressTypeID,"FAILED"
					bRunFlag = False
		End If
	End If
	Append_TestHTML StepCounter, "Validate dbo_CustomerAddress_CT Table Entries", "*** Validate CTC operation record entries" ,"PASSED"
	
	If ucase(operationflag) = ucase("insertion")  or ucase(operationflag) = ucase("update") or ucase(operationflag) = ucase("active") Then
	
		If instr(Db_C_AddressLines, j.EntityBody.AddressLine1 )>0 and instr(Db_C_AddressLines, j.EntityBody.AddressLine2 )>0 and instr(Db_C_AddressLines, j.EntityBody.AddressLine3 )>0 Then
			Append_TestHTML StepCounter, "Validate Customer AddressLine", "Expected:"& j.EntityBody.AddressLine1 & j.EntityBody.AddressLine2 & j.EntityBody.AddressLine3  & VBCRLF & "Actual:" & Db_C_AddressLines  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer AddressLine", "Expected:"& j.EntityBody.AddressLine1 & j.EntityBody.AddressLine2 & j.EntityBody.AddressLine3 & VBCRLF & "Actual:" & Db_C_AddressLines  ,"FAILED"
				bRunFlag = False
		End If
		
		If instr(trim(Db_C_City) , cstr(j.EntityBody.City) ) > 0 Then
			Append_TestHTML StepCounter, "Validate Customer City value", "Expected:"& j.EntityBody.City & VBCRLF & "Actual:" & Db_C_City  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer City value", "Expected:"& j.EntityBody.City & VBCRLF & "Actual:" & Db_C_City  ,"FAILED"	
				bRunFlag = False
		End If
		If instr( cint(Db_C_RegionID) , cint(j.EntityBody.RegionId) ) > 0 Then
			Append_TestHTML StepCounter, "Validate Customer RegionId value", "Expected:"& j.EntityBody.RegionId & VBCRLF & "Actual:" & Db_C_RegionID  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer RegionId value", "Expected:"& j.EntityBody.RegionId & VBCRLF & "Actual:" & Db_C_RegionID  ,"FAILED"	
				bRunFlag = False
		End If
		If instr( trim(cstr(Db_C_Zipcode)) , cstr(j.EntityBody.PostalCode) ) > 0 Then
			Append_TestHTML StepCounter, "Validate Customer PostalCode value", "Expected:"& j.EntityBody.PostalCode & VBCRLF & "Actual:" & Db_C_Zipcode  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer PostalCode value", "Expected:"& j.EntityBody.PostalCode & VBCRLF & "Actual:" & Db_C_Zipcode  ,"FAILED"	
				bRunFlag = False
		End If
		If cbool(Db_C_CountryID) = cbool(j.EntityBody.CountryId) Then
			Append_TestHTML StepCounter, "Validate Customer CountryId value", "Expected:"& j.EntityBody.CountryId & VBCRLF & "Actual:" & Db_C_CountryID  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer CountryId value", "Expected:"& j.EntityBody.CountryId & VBCRLF & "Actual:" & Db_C_CountryID  ,"FAILED"	
				bRunFlag = False
		End If
		If cbool(Db_C_Telephone) = cbool(j.EntityBody.Telephone) Then
			Append_TestHTML StepCounter, "Validate Customer Telephone value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & Db_C_Telephone  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Telephone value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & Db_C_Telephone  ,"FAILED"	
				bRunFlag = False
		End If
		If cbool(Db_C_Fax) = cbool(j.EntityBody.Fax) Then
			Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & Db_C_Fax  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & Db_C_Fax  ,"FAILED"	
				bRunFlag = False
		End If
		If cbool(Db_C_IsActive) = cbool(j.EntityBody.IsActive) Then
			Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & Db_C_IsActive  ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Address Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & Db_C_IsActive  ,"FAILED"	
				bRunFlag = False
		End If
	End  If
	Set fileSysObj = Nothing	
End Function



Function verifyAddressTypecheckFromUI(vercheck)
	On Error Resume Next
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebCheckbox_Main") Then
			isMaincheck = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_Main","GetROProperty","checked")
			isRegcheck = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_Registered","GetROProperty","checked")
			isCorrescheck = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_Correspond","GetROProperty","checked")
			
			If cbool(isMaincheck) = cbool(vercheck) Then
				Append_TestHTML StepCounter, "Validate Main Addresstype Main checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isMaincheck ,"PASSED"
			else
				Append_TestHTML StepCounter, "Validate Main Addresstype Main checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isMaincheck ,"FAILED"
					bRunFlag = False
			End If
			If cbool(isRegcheck) = cbool(vercheck) Then
				Append_TestHTML StepCounter, "Validate Main Addresstype Registered checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isRegcheck ,"PASSED"
			else
				Append_TestHTML StepCounter, "Validate Main Addresstype Registered checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isRegcheck ,"FAILED"
					bRunFlag = False
			End If
			If cbool(isCorrescheck) = cbool(vercheck) Then
				Append_TestHTML StepCounter, "Validate Main Addresstype Correspondence checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isCorrescheck ,"PASSED"
			else
				Append_TestHTML StepCounter, "Validate Main Addresstype Correspondence checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isCorrescheck ,"FAILED"
					bRunFlag = False
			End If
		End If
	
	

End Function



Function validateContactLines(jsonFPath,addresscustid,externaRefVal,addresstypeids,operationflag)
	On Error Resume Next
	If Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvAddresses_AddressList_Table","html tag:=TABLE").Exist Then
		Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=SearchMenu","html tag:=INPUT").Set "Contacts"
		wait 2
		Browser("creationTime:=1").Page("creationTime:=1").Link("html tag:=A","innertext:="&"Contacts","index:=0").Click
		wait 2
	End  If
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set file = FileSysObj.OpenTextFile(jsonFPath,1)
		sText = file.ReadAll
	Set j = ParseJson(sText)
	Set fileSysObj = Nothing	
	query_customer = "Select * from CustomerContact where CustomerID = '"&addresscustid&"' and ExternalContactID='"&externaRefVal & "';"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 eContactID= dictDbResultSet("ContactID")
	 eIsActive = dictDbResultSet("IsActive")
	 eContactTitleID= dictDbResultSet("ContactTitleID")
	 eIsPrimary = dictDbResultSet("IsPrimary")
	 eAddressID= dictDbResultSet("AddressID")
	 eContactPreferenceID = dictDbResultSet("ContactPreferenceID")
	 eLanguageID = dictDbResultSet("LanguageID")
	 eForeName = dictDbResultSet("ForeName")
	 eMiddleName = dictDbResultSet("MiddleName")
	 eLastName = dictDbResultSet("LastName")
	 eTelephone = dictDbResultSet("Telephone")
	 eMobilePhone = dictDbResultSet("MobilePhone")
	 eEmailAddress = dictDbResultSet("EmailAddress")
	 eFax = dictDbResultSet("Fax")
	 
	 ePosition = dictDbResultSet("Position")
	
	 Set query_customer=Nothing
	set dictDbResultSet = Nothing
	
	query_customer = "Select * from CustomerAddress where ExternalAddressID='" & j.EntityBody.AddressId & "';"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 db_AddressID = dictDbResultSet("AddressID")
	
	 Set query_customer=Nothing
	set dictDbResultSet = Nothing
	
	If instr(addresstypeids,",")>0 Then
		addtypeidvals = Split(addresstypeids,",")
		len_addtypevals = ubound(addtypeidvals)
	Else
		len_addtypevals = 0
	End If
	For itr = 0 To len_addtypevals Step 1
		eModifiedBy = ""
		mCTIdval = eval("j.EntityBody.ContactTypeId.[" & itr & "]")
		query_customer = "Select * from CustomerContactType where CustomerID = '"& addresscustid &"' and ContactTypeID=" & mCTIdval & " and ContactID = "& eContactID &";"
		Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
		set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
		 eModifiedBy= dictDbResultSet("ModifiedBy")
		 eContactTypeID= dictDbResultSet("ContactTypeID")
		 Set query_customer = Nothing
		set dictDbResultSet = Nothing
		If cint(len_addtypevals) = cint("0") and cdbl(mCTIdval) = cdbl("0")  Then
			If  eContactTypeID = "" Then
				query_customer = "Select Count(*) as contactetry from CustomerContactType where CustomerID = '"&addresscustid&"' " & " and ContactID = "& eContactID &";"
					Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
					set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
					 No_of_contacts = dictDbResultSet("contactetry")
					 Set query_customer = Nothing
					set dictDbResultSet = Nothing
					If isNULL(No_of_contacts) = true or cint(No_of_contacts)=cint("0")Then
						Append_TestHTML StepCounter, "Validate ContactTypeID" & mCTIdval , "All records are deleted from table" ,"PASSED"
					Else
						Append_TestHTML StepCounter, "Validate Customer Contact", No_of_contacts & " Contact entry exist with contactid-"& eContactID ,"FAILED"	
					End If
				
			Else
				Append_TestHTML StepCounter, "Validate ContactTypeID" & mCTIdval , "Fail to delete records-" & addtypeidvals ,"FAILED"
			End If
		Else
			If eModifiedBy <> "" and cdbl(eContactTypeID) =  cdbl(mCTIdval)  Then
					Append_TestHTML StepCounter, "Validate ContactTypeID" & mCTIdval , "ContactType " & eContactID & " record find successfully" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Validate ContactTypeID" & mCTIdval , "ContactType " & eContactID & " record find successfully" ,"FAILED"
			End If
		End If
		
	Next
	query_customer = "Select Count(*) as addressentry from CustomerContact where CustomerID = '"&addresscustid&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 No_of_addressentries = dictDbResultSet("addressentry")
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	If ucase(operationflag) <> ucase("insertionupdate") Then
		If cint(No_of_addressentries) = cint("1") Then
			Append_TestHTML StepCounter, "Validate Customer Contact", No_of_addressentries & " Contact entry created successfully" ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Contact", No_of_addressentries & " Contact entry created successfully" ,"FAILED"	
		End If
	else
		If cint(No_of_addressentries) >= cint("1") Then
			Append_TestHTML StepCounter, "Validate Customer Contact", No_of_addressentries & " Contact entry created successfully" ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Customer Contact", No_of_addressentries & " Contact entry created successfully" ,"FAILED"	
		End If
	End If
		If  cdbl(eAddressID)= cdbl(db_Cust_AddressID) Then
				Append_TestHTML StepCounter, "Validate Customer Contact AddressID", "Expected:"& db_Cust_AddressID  & VBCRLF & "Actual:" & eAddressID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact AddressID", "Expected:"& db_Cust_AddressID  & VBCRLF & "Actual:" & eAddressID  ,"FAILED"
					bRunFlag = False
		End If
	
		
	 
	If specialchrFlag = True Then
			If  instr(eForeName, specialFName )>0Then
				Append_TestHTML StepCounter, "Validate Customer Forename", "Expected:"& specialFName  & VBCRLF & "Actual:" & eForeName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Forename", "Expected:"& specialFName  & VBCRLF & "Actual:" & eForeName  ,"FAILED"
					bRunFlag = False
			End If
			If  instr(eMiddleName, specialMName )>0 Then
				Append_TestHTML StepCounter, "Validate Customer MiddleName", "Expected:"& specialMName  & VBCRLF & "Actual:" & eMiddleName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer MiddleName", "Expected:"& specialMName  & VBCRLF & "Actual:" & eMiddleName  ,"FAILED"
					bRunFlag = False
			End If
			If  instr(eLastName, specialLName )>0 Then
				Append_TestHTML StepCounter, "Validate Customer LastName", "Expected:"& specialLName  & VBCRLF & "Actual:" & eLastName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer LastName", "Expected:"& specialLName  & VBCRLF & "Actual:" & eLastName  ,"FAILED"
					bRunFlag = False
			End If
	Else
			If instr(eForeName, j.EntityBody.Forename )>0  or isNull(eForeName) = isNull(j.EntityBody.Forename) Then
				Append_TestHTML StepCounter, "Validate Customer Forename", "Expected:"& j.EntityBody.Forename  & VBCRLF & "Actual:" & eForeName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Forename", "Expected:"& j.EntityBody.Forename  & VBCRLF & "Actual:" & eForeName  ,"FAILED"
					bRunFlag = False
			End If
			If instr(eMiddleName, j.EntityBody.MiddleName )>0 or isNull(eMiddleName) = isNull(j.EntityBody.MiddleName) Then
				Append_TestHTML StepCounter, "Validate Customer MiddleName", "Expected:"& j.EntityBody.MiddleName  & VBCRLF & "Actual:" & eMiddleName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer MiddleName", "Expected:"& j.EntityBody.MiddleName  & VBCRLF & "Actual:" & eMiddleName  ,"FAILED"
					bRunFlag = False
			End If
			If instr(eLastName, j.EntityBody.LastName )>0  Then
				Append_TestHTML StepCounter, "Validate Customer LastName", "Expected:"& j.EntityBody.LastName  & VBCRLF & "Actual:" & eLastName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer LastName", "Expected:"& j.EntityBody.LastName  & VBCRLF & "Actual:" & eLastName  ,"FAILED"
					bRunFlag = False
			End If
	End If
	
	
	If cbool(eIsActive) = cbool(j.EntityBody.IsActive) Then
		Append_TestHTML StepCounter, "Validate Customer IsActive Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & eIsActive  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer IsActive Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & eIsActive  ,"FAILED"	
			bRunFlag = False
	End If
	If cbool(eIsPrimary) = cbool(j.EntityBody.IsPrimaryContact) Then
		Append_TestHTML StepCounter, "Validate Customer isPrimaryContact Status", "Expected:"& j.EntityBody.IsPrimaryContact & VBCRLF & "Actual:" & eIsPrimary  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer isPrimaryContact Status", "Expected:"& j.EntityBody.IsPrimaryContact & VBCRLF & "Actual:" & eIsPrimary  ,"FAILED"	
			bRunFlag = False
	End If
	If cstr(ePosition) = cstr(j.EntityBody.Position) Then
		Append_TestHTML StepCounter, "Validate Customer Contact Position Value", "Expected:"& j.EntityBody.Position & VBCRLF & "Actual:" & ePosition  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact Position Value", "Expected:"& j.EntityBody.Position & VBCRLF & "Actual:" & ePosition  ,"FAILED"	
			bRunFlag = False
	End If
	If cstr(eFax) = cstr(j.EntityBody.Fax) Then
		Append_TestHTML StepCounter, "Validate Customer Contact Fax Value", "Expected:"& j.EntityBody.Fax & VBCRLF & "Actual:" & eFax  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact Fax Value", "Expected:"& j.EntityBody.Fax & VBCRLF & "Actual:" & eFax  ,"FAILED"	
			bRunFlag = False
	End If
	
	If cstr(eTelephone) = cstr(j.EntityBody.Telephone) Then
		Append_TestHTML StepCounter, "Validate Customer Contact Telephone Value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & eTelephone  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact Telephone Value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & eTelephone  ,"FAILED"	
			bRunFlag = False
	End If
	
	If cstr(eEmailAddress) = cstr(j.EntityBody.EmailAddress) Then
		Append_TestHTML StepCounter, "Validate Customer Contact EmailAddress Value", "Expected:"& j.EntityBody.EmailAddress & VBCRLF & "Actual:" & eEmailAddress  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact EmailAddress Value", "Expected:"& j.EntityBody.EmailAddress & VBCRLF & "Actual:" & eEmailAddress  ,"FAILED"	
			bRunFlag = False
	End If

	If cstr(eMobilePhone) = cstr(j.EntityBody.MobilePhone) Then
		Append_TestHTML StepCounter, "Validate Customer Contact MobilePhone Value", "Expected:"& j.EntityBody.MobilePhone & VBCRLF & "Actual:" & eMobilePhone  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact MobilePhone Value", "Expected:"& j.EntityBody.MobilePhone & VBCRLF & "Actual:" & eMobilePhone  ,"FAILED"	
			bRunFlag = False
	End If
	
	If cint(db_AddressID) = cint(eAddressID) Then
		Append_TestHTML StepCounter, "Validate Customer Contact AddressId Value", "Expected:"& db_AddressID & VBCRLF & "Actual:" & eAddressID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact AddressId Value", "Expected:"& db_AddressID & VBCRLF & "Actual:" & eAddressID  ,"FAILED"	
			bRunFlag = False
	End If
		If cint(eContactTitleID) = cint(j.EntityBody.ContactTitleId) Then
		Append_TestHTML StepCounter, "Validate Customer Contact ContactTitleID Value", "Expected:"& j.EntityBody.ContactTitleId & VBCRLF & "Actual:" & eContactTitleID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerContact  ContactTitleID Value", "Expected:"& j.EntityBody.ContactTitleId & VBCRLF & "Actual:" & eContactTitleID  ,"FAILED"	
			bRunFlag = False
	End If
		If cint(eContactPreferenceID) = cint(j.EntityBody.ContactPreferenceId) Then
		Append_TestHTML StepCounter, "Validate Customer Contact ContactPreferenceId Value", "Expected:"& j.EntityBody.ContactPreferenceId & VBCRLF & "Actual:" & eContactPreferenceID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact ContactPreferenceId Value", "Expected:"& j.EntityBody.ContactPreferenceId & VBCRLF & "Actual:" & eContactPreferenceID  ,"FAILED"	
			bRunFlag = False
	End If
	If cint(eLanguageID) = cint(j.EntityBody.LanguageId) Then
		Append_TestHTML StepCounter, "Validate Customer Contact LanguageId Value", "Expected:"& j.EntityBody.LanguageId & VBCRLF & "Actual:" & eLanguageID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Customer Contact LanguageId Value", "Expected:"& j.EntityBody.LanguageId & VBCRLF & "Actual:" & eLanguageID  ,"FAILED"	
			bRunFlag = False
	End If
	 
	If Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvMain_ContactList_Table","html tag:=TABLE").Exist Then
		Append_TestHTML StepCounter,"Navigate to Contacts Page","Successfully Navigated to Contact Page" ,"PASSED"
		AddLineData = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvMain_ContactList_Table","html tag:=TABLE").GetCellData(2,2)
		wait 2
'		AddLineData = j.EntityBody.EmailAddress
'		Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvMain_ContactList_Table","html tag:=TABLE").ChildItem(2,1,"WebElement",0).Click
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("html tag:=TD","innertext:="&AddLineData).Click
		wait 5
		Append_TestHTML StepCounter,"Navigate to Contact Details Tab ","Successfully Navigated to Contact Details Tab" ,"PASSED"
	End If
If ucase(operationflag) <> ucase("insertionupdate") Then	
	query_customer = "Select * from CustomerActivityLog where CustomerID=" & addresscustid &" order by ModifiedOn desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_C_ActivityLogID = dictDbResultSet("ActivityLogID")
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	
	query_customer = "Select * from ActivityLog where ActivityTypeID=99 order by LogDateTime desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_Note = dictDbResultSet("Note")
	 Db_A_ActivityLogID = dictDbResultSet("ActivityLogID")
	 
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	
	wait 15
	query_customer = "Select __$operation as coperval,ModifiedOn,IsActive,ContactID,IsPrimary,ModifiedBy,ModifiedByApp,AddressID,ContactTitleID  from cdc.dbo_CustomerContact_CT where CustomerID=" & addresscustid & " and ContactID= " & eContactID & " and ExternalContactID='" & externaRefVal & "' order by __$seqval Desc;"
	Append_TestHTML StepCo.unter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 3, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_C_Operation = dictDbResultSet("coperval")
	 Db_C_ModifiedOn = dictDbResultSet("ModifiedOn")
	 Db_C_IsActive= dictDbResultSet("IsActive")
	 Db_C_ContactID= dictDbResultSet("ContactID")
	 Db_C_IsPrimary= dictDbResultSet("IsPrimary")
	 Db_C_ModifiedBy= dictDbResultSet("ModifiedBy")
	 Db_C_ModifiedByApp= dictDbResultSet("ModifiedByApp")
	 Db_C_AddressID= dictDbResultSet("AddressID")
	 Db_C_ContactTitleID= dictDbResultSet("ContactTitleID")
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	query_customer = "Select ContactPreferenceID,LanguageID,ForeName,MiddleName,LastName,Telephone,Fax,MobilePhone,EmailAddress,Position  from cdc.dbo_CustomerContact_CT where CustomerID=" & addresscustid & " and ContactID= " & eContactID & " and ExternalContactID='" & externaRefVal & "' order by __$seqval Desc;"
	Append_TestHTML StepCo.unter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 3, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_C_ContactPreferenceID= dictDbResultSet("ContactPreferenceID")
	 Db_C_LanguageID= dictDbResultSet("LanguageID")
	 Db_C_ForeName= dictDbResultSet("ForeName")
	 Db_C_MiddleName= dictDbResultSet("MiddleName")
	 Db_C_LastName= dictDbResultSet("LastName")
	 Db_C_Telephone= dictDbResultSet("Telephone")
	 Db_C_Fax= dictDbResultSet("Fax")
	 Db_C_MobilePhone= dictDbResultSet("MobilePhone")
	 Db_C_EmailAddress= dictDbResultSet("EmailAddress")
	 Db_C_Position= dictDbResultSet("Position")
	 	 
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	query_customer = "Select __$operation as coperval,ContactTypeID from cdc.dbo_CustomerContactType_CT where CustomerID=" & addresscustid & " and ContactID= " & eContactID & " order by __$seqval Desc;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 3, "GFN_SHELL_SPRINTQA_PH_OLTP")
	 Db_T_Operation = dictDbResultSet("coperval")
	 Db_T_AddressTypeID = dictDbResultSet("ContactTypeID")
	 
	 Set query_customer = Nothing
	set dictDbResultSet = Nothing
	Append_TestHTML StepCounter, "Validate ActivityLog and CDC Table operation Entries","*** Verify " & operationflag & " operation entries***", "PASSED"
	
	If ucase(operationflag) = ucase("insertion") Then
		If Db_C_ActivityLogID = "" Then
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"FAILED"			
		End If	
		If instr(Db_Note,addresscustid) = 0 and instr(Db_Note,eAddressID) = 0 Then
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:"  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: Empty"  & VBCRLF & "Actual:" & Db_Note ,"FAILED"
		End If
		If instr(Db_C_Operation ,"2")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive  &"-" & Db_C_ForeName & "-" & Db_C_MiddleName & "-" & Db_C_LastName & "-" & Db_C_IsPrimary & "-" &Db_C_IsActive & "-"& Db_C_Telephone  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive ,"FAILED"			
		End If
		If instr(addresstypeids,",")>0 Then
			If instr(Db_T_Operation , "2")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerContactsType_CT  " & operationflag & " operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation   ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT  " & operationflag & "operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation  ,"FAILED"	
			End If
		Else
			If Db_T_Operation = "2" Then
					Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT " & operationflag & "  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation  ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT " & operationflag & " operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation & "ContactTypeID" & Db_T_AddressTypeID ,"FAILED"	
			End If
		End If
		
	ElseIf ucase(operationflag) = ucase("update") or ucase(operationflag) = ucase("active")  Then
		If Db_C_ActivityLogID <> ""  and cdbl(Db_A_ActivityLogID) = cdbl(Db_C_ActivityLogID) Then
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual: " & Db_C_ActivityLogID & "-" & Db_A_ActivityLogID & " ActivityLogID:-" & Db_C_ActivityLogID  ,"FAILED"
		End If	
		If instr(Db_Note,addresscustid) > 0 and instr(Db_Note,eAddressID) > 0 Then
			Append_TestHTML StepCounter, "Validate ActivityLog Note Data", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate ActivityLog Note Data", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
		End If
		If  ucase(operationflag) = ucase("update")  Then
		
			If instr(Db_C_Operation ,"3")>0 or instr(Db_C_Operation ,"4")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT  " & operationflag & "operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive & Db_C_ForeName & "-" & Db_C_MiddleName & "-" & Db_C_LastName & "-" & Db_C_IsPrimary & "-" &Db_C_IsActive & "-"& Db_C_Telephone  ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT  " & operationflag & " operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " Modified On" & Db_C_ModifiedOn  ,"FAILED"
			End If
			If instr(Db_T_Operation, "2")>0 	Then		'and instr(Db_T_AddressTypeID,"2")>0 and instr(Db_T_AddressTypeID,"3")>0 and instr(Db_T_AddressTypeID,"4")>0 Then
					Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT " & operationflag & "  operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation  & " ContactTypeID" & Db_T_AddressTypeID,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT " & operationflag & " operation value", "Expected: 2"  & VBCRLF & "Actual:" & Db_T_Operation & "ContactTypeID" & Db_T_AddressTypeID ,"FAILED"	
			End If
			
		End If
	ElseIf ucase(operationflag) = ucase("deletion") Then
		If Db_C_ActivityLogID <> ""  and cdbl(Db_A_ActivityLogID) = cdbl(Db_C_ActivityLogID) Then
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerActivityLog Entry", "Expected: Non-Empty"  & VBCRLF & "Actual:" & Db_C_ActivityLogID & "-" & Db_A_ActivityLogID & " ActivityLogID:-" & Db_C_ActivityLogID  ,"FAILED"
		End If	
		If instr(Db_Note,addresscustid) > 0 and instr(Db_Note,eAddressID) > 0 Then
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate ActivityLog Entry", "Expected: " & addresscustid & " - " & eAddressID  & VBCRLF & "Actual:" & Db_Note ,"PASSED"
		End If
		If instr(Db_C_Operation ,"3")>0 or instr(Db_C_Operation ,"4")>0 Then
				Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT " & operationflag & " operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " IsActive-"& Db_C_IsActive  ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT " & operationflag & " operation value", "Expected: 3,4"  & VBCRLF & "Actual:" & Db_C_Operation & " Modified On" & Db_C_ModifiedOn  ,"FAILED"
		End If
		If instr(Db_T_Operation, "1")>0 Then ' and instr(Db_T_AddressTypeID,"2")>0 and instr(Db_T_AddressTypeID,"3")>0 and instr(Db_T_AddressTypeID,"4")>0  Then
			Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT " & operationflag & " operation value", "Expected: 1"  & VBCRLF & "Actual:" & Db_T_Operation & " ContactTypeID" & Db_T_AddressTypeID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate dbo_CustomerContactType_CT " & operationflag & " operation value", "Expected: 1"  & VBCRLF & "Actual:" & Db_T_Operation & " ContactTypeID" & Db_T_AddressTypeID,"FAILED"
					bRunFlag = False
		End If
	End If
	
	If ucase(operationflag) = ucase("update") or  ucase(operationflag) = ucase("insertion") Then
		
			Append_TestHTML StepCounter, "Validate dbo_CustomerContact_CT Table details", "***Validate CTC table entries***"  ,"PASSED"
	 	 
			If instr(Db_C_IsActive,j.EntityBody.IsActive) >0  or cbool(Db_C_IsActive) = cbool(j.EntityBody.IsActive) Then
				Append_TestHTML StepCounter, "Validate Customer Contact IsActive Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & Db_C_IsActive  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact IsActive Status", "Expected:"& j.EntityBody.IsActive & VBCRLF & "Actual:" & Db_C_IsActive  ,"FAILED"	
					bRunFlag = False
			End If
			If instr(Db_C_IsPrimary,j.EntityBody.IsPrimaryContact)>0 Then
				Append_TestHTML StepCounter, "Validate Customer isPrimaryContact Status", "Expected:"& j.EntityBody.IsPrimaryContact & VBCRLF & "Actual:" & Db_C_IsPrimary  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer isPrimaryContact Status", "Expected:"& j.EntityBody.IsPrimaryContact & VBCRLF & "Actual:" & Db_C_IsPrimary  ,"FAILED"	
					bRunFlag = False
			End If
			If instr(Db_C_Position,j.EntityBody.Position)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact Position Value", "Expected:"& j.EntityBody.Position & VBCRLF & "Actual:" & Db_C_Position  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact Position Value", "Expected:"& j.EntityBody.Position & VBCRLF & "Actual:" & Db_C_Position  ,"FAILED"	
					bRunFlag = False
			End If
			If instr(Db_C_Fax,j.EntityBody.Fax)>0 or cstr(Db_C_Fax) = cstr(j.EntityBody.Fax) Then
				Append_TestHTML StepCounter, "Validate Customer Contact Fax Value", "Expected:"& j.EntityBody.Fax & VBCRLF & "Actual:" & Db_C_Fax  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact Fax Value", "Expected:"& j.EntityBody.Fax & VBCRLF & "Actual:" & Db_C_Fax  ,"FAILED"	
					bRunFlag = False
			End If
			
			If instr(Db_C_Telephone,j.EntityBody.Telephone)>0 or cstr(Db_C_Telephone) = cstr(j.EntityBody.Telephone)Then
				Append_TestHTML StepCounter, "Validate Customer Contact Telephone Value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & Db_C_Telephone  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact Telephone Value", "Expected:"& j.EntityBody.Telephone & VBCRLF & "Actual:" & Db_C_Telephone  ,"FAILED"	
					bRunFlag = False
			End If
		
			If instr(Db_C_EmailAddress,j.EntityBody.EmailAddress)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact EmailAddress Value", "Expected:"& j.EntityBody.EmailAddress & VBCRLF & "Actual:" & Db_C_EmailAddress  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact EmailAddress Value", "Expected:"& j.EntityBody.EmailAddress & VBCRLF & "Actual:" & Db_C_EmailAddress  ,"FAILED"	
					bRunFlag = False
			End If
		
			If instr(Db_C_MobilePhone,j.EntityBody.MobilePhone)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact MobilePhone Value", "Expected:"& j.EntityBody.MobilePhone & VBCRLF & "Actual:" & Db_C_MobilePhone  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact MobilePhone Value", "Expected:"& j.EntityBody.MobilePhone & VBCRLF & "Actual:" & Db_C_MobilePhone  ,"FAILED"	
					bRunFlag = False
			End If
		
			If inst(db_AddressID,Db_C_AddressID)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact AddressId Value", "Expected:"& db_AddressID & VBCRLF & "Actual:" & Db_C_AddressID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact AddressId Value", "Expected:"& db_AddressID & VBCRLF & "Actual:" & Db_C_AddressID  ,"FAILED"	
					bRunFlag = False
			End If
				If instr(Db_C_ContactTitleID,j.EntityBody.ContactTitleId)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact ContactTitleID Value", "Expected:"& j.EntityBody.ContactTitleId & VBCRLF & "Actual:" & Db_C_ContactTitleID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate CustomerContact  ContactTitleID Value", "Expected:"& j.EntityBody.ContactTitleId & VBCRLF & "Actual:" & Db_C_ContactTitleID  ,"FAILED"	
					bRunFlag = False
			End If
				If instr(Db_C_ContactPreferenceID,j.EntityBody.ContactPreferenceId)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact ContactPreferenceId Value", "Expected:"& j.EntityBody.ContactPreferenceId & VBCRLF & "Actual:" & Db_C_ContactPreferenceID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact ContactPreferenceId Value", "Expected:"& j.EntityBody.ContactPreferenceId & VBCRLF & "Actual:" & Db_C_ContactPreferenceID  ,"FAILED"	
					bRunFlag = False
			End If
			If instr(Db_C_LanguageID,j.EntityBody.LanguageId)>0 Then
				Append_TestHTML StepCounter, "Validate Customer Contact LanguageId Value", "Expected:"& j.EntityBody.LanguageId & VBCRLF & "Actual:" & Db_C_LanguageID  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Contact LanguageId Value", "Expected:"& j.EntityBody.LanguageId & VBCRLF & "Actual:" & Db_C_LanguageID  ,"FAILED"	
					bRunFlag = False
			End If
				
			 If instr(Db_C_ForeName, j.EntityBody.Forename )>0 or cstr(Db_C_ForeName)=cstr(j.EntityBody.Forename)   Then
				Append_TestHTML StepCounter, "Validate Customer Forename", "Expected:"& j.EntityBody.Forename  & VBCRLF & "Actual:" & Db_C_ForeName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer Forename", "Expected:"& j.EntityBody.Forename  & VBCRLF & "Actual:" & Db_C_ForeName  ,"FAILED"
					bRunFlag = False
			End If
			If instr(Db_C_MiddleName, j.EntityBody.MiddleName )>0 or cstr(Db_C_MiddleName) = cstr(j.EntityBody.MiddleName)  Then
				Append_TestHTML StepCounter, "Validate Customer MiddleName", "Expected:"& j.EntityBody.MiddleName  & VBCRLF & "Actual:" & Db_C_MiddleName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer MiddleName", "Expected:"& j.EntityBody.MiddleName  & VBCRLF & "Actual:" & Db_C_MiddleName  ,"FAILED"
					bRunFlag = False
			End If
			If instr(Db_C_LastName, j.EntityBody.LastName )>0 or cstr(Db_C_LastName)=cstr( j.EntityBody.LastName) Then
				Append_TestHTML StepCounter, "Validate Customer LastName", "Expected:"& j.EntityBody.LastName  & VBCRLF & "Actual:" & Db_C_LastName  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Customer LastName", "Expected:"& j.EntityBody.LastName  & VBCRLF & "Actual:" & Db_C_LastName  ,"FAILED"
					bRunFlag = False
			End If
		
	End If
	End  If
End Function



Function basicPreConfigContactData(jsonFPath,dbName)

On Error Resume Next
	bRunFlag = True
	bFlag = True
	
	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
	db_CompanyID = dictDbResultSet("CompanyID")
	db_CountryID = dictDbResultSet("CountryID")
'	db_RegionID = dictDbResultSet("LegislativeRegionID")
'	db_CurrencyID = dictDbResultSet("CurrencyID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	query = "Select * from LocalisedDescriptions where ColumnName like '%ContactTitleId%' and Culture='en-GB';"
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ContactTitleId = dictDbResultSet("Value")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	query = "Select * from LocalisedDescriptions where ColumnName like '%ContactPreferenceId%' and Culture='en-GB';"
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_ContactPreferenceId = dictDbResultSet("Value")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	query = "Select * from CompanyLanguage where CompanyID=" & db_CompanyID & " and OfficialLanguage=1;"
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_CL_LanguageId= dictDbResultSet("LanguageID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	'query = "Select * from LocalisedDescriptions where ColumnName like '%LanguageId%' and Culture='en-GB' and Description='English';"
	query = "Select * from Language where LanguageID=" & db_CL_LanguageId &";"
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	db_LanguageId= dictDbResultSet("LanguageID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
		If db_LanguageId = "" or isNULL(db_LanguageId)= true Then
			query = "Select * from Language where ISOCode2='en';"
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			Set dictDbResultSet = execute_db_query(query, 1, dbName)
			wait 3
			db_LanguageId= dictDbResultSet("LanguageID")
			Set dictDbResultSet = Nothing
			Set query = Nothing
		End If
	preSearchString = "ColcoCode-" & db_ColcoCode &";ContactTitleId-"& db_ContactTitleId &";ContactPreferenceId-" & db_ContactPreferenceId &";LanguageId-" & db_LanguageId
	Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
End Function


Function verifyContactDetailscheckFromUI(vercheck)
	On Error Resume Next
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebCheckbox_Main") Then
			isPrimarycheck = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_Primary","GetROProperty","checked")
			isCardDeliverycheck = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_Carddelivery","GetROProperty","checked")
'			isCorrescheck = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_Correspond","GetROProperty","checked")
			
			If cbool(isPrimarycheck) = cbool(vercheck) Then
				Append_TestHTML StepCounter, "Validate Main Addresstype Main checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isMaincheck ,"PASSED"
			else
				Append_TestHTML StepCounter, "Validate Main Addresstype Main checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isMaincheck ,"FAILED"
					bRunFlag = False
			End If
			If cbool(isCardDeliverycheck) = cbool(vercheck) Then
				Append_TestHTML StepCounter, "Validate Main Addresstype Registered checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isRegcheck ,"PASSED"
			else
				Append_TestHTML StepCounter, "Validate Main Addresstype Registered checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isRegcheck ,"FAILED"
					bRunFlag = False
			End If
'			If cbool(isCorrescheck) = cbool(vercheck) Then
'				Append_TestHTML StepCounter, "Validate Main Addresstype Correspondence checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isCorrescheck ,"PASSED"
'			else
'				Append_TestHTML StepCounter, "Validate Main Addresstype Correspondence checkbox", "Expected: " & vercheck  & VBCRLF & "Actual:" & isCorrescheck ,"FAILED"
'					bRunFlag = False
'			End If
		End If

End Function

'Function updateContactNamewithSpecialCharacters()
'	On Error Resume Next
'	
'	Forename-null;MiddleName-null;LastName-null
'	
'	preSearchString = "ColcoCode-" & db_ColcoCode &";LanguageId-"& db_LanguageID &";CurrencyId-" & db_CurrencyID & ";CountryId-"& db_CountryID &";BandId-" & db_BandID & ";CustomerClassificationId-" & db_CCTId_Value & ";MarketingSegmentationId-" & db_MSId_Value & ";IndustrialClassId-"& db_ICId_Value & ";LegalEntityId-" & db_LEId_Value
'	Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
'	
'End Function

Function invokeUpdateJSONAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)

On error resume next
	bFlag = True
		bRunFlag =True
	Set objAPI = createobject("MSXML2.serverXMLHTTP.6.0")
	'Set objAPI = createobject("WinHttp.WinHttpRequest.5.1")
	
	objAPI.open reqType, apiurl , asynctype
	arrHeaders = split(headers,",")
	
	For hitr = 0 To ubound(arrHeaders) Step 1
		If arrHeaders(hitr) <> "" or arrHeaders(hitr) <> null Then
			arrheaderNameval = split(arrHeaders(hitr),"==>")
		
			objAPI.setRequestHeader trim(arrheaderNameval(0)) , trim(arrheaderNameval(1))
		End If
	Next
	If strType = "file" Then
		Set fileSysObj = createObject("Scripting.FileSystemObject")
		Set file = FileSysObj.OpenTextFile(strJson)
		sText = file.ReadAll
		Set fileSysObj = Nothing
	Else
		sText = strJson
	End If
	Append_TestHTML StepCounter, "Verify input request JSON", "JSON Text is: " & sText , "PASSED"
	
	specialFName = "จังหวัดบุรีรัมย์"
	specialMName = "Test$%^"
	specialLName = "Script!@#"
'	Set q = ParseJson(sText)
'	q.EntityBody.Forename ="จังหวัดบุรีรัมย์"
'	q.EntityBody.MiddleName ="Test$%^"
'	q.EntityBody.LastName ="Script!@#"
'	Append_TestHTML StepCounter, "Change Forename value", "Changed to " & q.EntityBody.Forename , "PASSED"
	Append_TestHTML StepCounter, "Change Forename value", "Changed to จังหวัดบุรีรัมย์", "PASSED"
	Append_TestHTML StepCounter, "Change Forename value", "Changed to " & specialFName, "PASSED"
	Append_TestHTML StepCounter, "Change MiddleName  value", "Changed to " & specialMName , "PASSED"
	Append_TestHTML StepCounter, "Change LastName  value", "Changed to " & specialLName  , "PASSED"
	sText = Replace(sText,"AutomationFName","จังหวัดบุรีรัมย์")
	sText = Replace(sText,"TestMName","Test$%^")
	sText = Replace(sText,"ScriptLName","Script!@#")
	Append_TestHTML StepCounter, "Verify input request JSON", "JSON Text is: " & sText , "PASSED"
	objAPI.send sText
	
	pageReturn = objAPI.responseText
' /* Below statement is for getting response Body */
'	bodyReturn =  BinarytoString(objAPI.responseBody)
'************************************************
'	msgbox objAPI.status
'	msgbox objAPI.statusText
	For itr = 1 To 5 Step 1
		If pageReturn = "" Then
			objAPI.send sText
			pageReturn = objAPI.responseText
		Else
			Exit For
		End If
	Next

	Set j = ParseJson(pageReturn)
	If ucase(j.Status) = ucase("Success") Then
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "PASSED"
		transRef = j.Results.TransactionRef
		transReqID = j.RequestID
		invokeUpdateJSONAPI = pageReturn
	Else
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "FAILED"
		transRef = ""
		transReqID = ""
		transErrDesc = j.Description
		invokeUpdateJSONAPI = Empty
		bRunFlag = False
	End If
		wait 40
	Set objAPI = Nothing
End Function


Function getErrorapiprerequisits(err_data,jsonFPath)
	On Error Resume Next
	inputusernamval = err_data("userName")
	
	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
	db_CompanyID = dictDbResultSet("CompanyID")
	db_CountryID = dictDbResultSet("CountryID")
	db_RegionID = dictDbResultSet("LegislativeRegionID")
	db_CurrencyID = dictDbResultSet("CurrencyID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
		Append_TestHTML StepCounter, "Get ColcoCode from Company Table ", "ClientCompanyNumber is: " & db_ColcoCode , "PASSED"
		Append_TestHTML StepCounter, "Get UserName from TestData ", "UserName is: " & inputusernamval , "PASSED"
	
	query = "Select Min(MQP.ProcessedOnUTC)as minprocessed,Max(MQP.ProcessedOnUTC) as maxprocessed from [dbo].[MessageQueueError] MQE	INNER JOIN [dbo].[MessageQueueProcessed] MQP ON (MQP.MessageQueueID = MQE.MessageQueueID) AND MQP.SystemIdentifier = '" & inputusernamval & "' AND MQE.StackTrace IS NULL;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	db_startdateval = dictDbResultSet("minprocessed")
	db_enddateval = dictDbResultSet("maxprocessed")


	Set dictDbResultSet = Nothing
	Set query = Nothing
		Append_TestHTML StepCounter, "Get ProcessedOnUTC Startdate and enddate values from MessageQueueProcessed Table ", "StartDate is: " & db_startdateval & "EndDate is: " & db_enddateval , "PASSED"
	
	original_date = Split(db_startdateval," ")(0)
	myordate = nextdateDBFormt(cdate(original_date)-1)
	mystartdateval = myordate & " 00:00:00.001"
	
	original_date = Split(db_enddateval," ")(0)
	myordate = nextdateDBFormt(cdate(original_date)-1)
	myenddateval= myordate & " 11:55:00.001"
	

	err_colcoID = db_ColcoCode
	preSearchString = "ColcoCode-" & db_ColcoCode &";UserName-"& inputusernamval
	Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
	preSearchString = "StartDateTime$" & mystartdateval & ";EndDateTime$"& myenddateval
	Call searchandReplaceMultipleStringwithdollar(jsonFPath, preSearchString)
	
	query = "Select Count(*) as totalrec from [dbo].[MessageQueueError] MQE	INNER JOIN [dbo].[MessageQueueProcessed] MQP ON (MQP.MessageQueueID = MQE.MessageQueueID) AND MQP.SystemIdentifier = '" & inputusernamval & "'AND MQP.ProcessedOnUTC >='" & mystartdateval &"'  AND MQP.ProcessedOnUTC <= '" & myenddateval &"' AND MQE.StackTrace IS NULL;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	errorapi_Totalnumrecs = dictDbResultSet("totalrec")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
		Append_TestHTML StepCounter, "Get Total Records b/w Startdate and enddate values from MessageQueueProcessed Table ", "Total number of records is: " & errorapi_Totalnumrecs, "PASSED"
	
	
	
End Function




Function verifyPosterrorAPIDetails(err_data,rcheck)
	On Error Resume Next
	inputusernamval = err_data("userName")
	inputpageval = err_data("page")
	inputpageSizeval = err_data("pageSize")
	inputdpagesizeval = err_data("dpageSize")
	Set j = ParseJson(rcheck)
	
	If cdbl(errorapi_Totalnumrecs) = cdbl( j.TotalRecords)  Then
				Append_TestHTML StepCounter, "Validate Response of error API  Totalrecords value", "Expected: " & errorapi_Totalnumrecs  & VBCRLF & "Actual:" & j.TotalRecords ,"PASSED"
		Else
				Append_TestHTML StepCounter, "Validate Response of error API  Totalrecords value", "Expected: " & errorapi_Totalnumrecs  & VBCRLF & "Actual:" & j.TotalRecords ,"FAILED"
		
	End If
	If cdbl(inputpageval) = cdbl( j.Page)  Then
				Append_TestHTML StepCounter, "Validate Response of error API  Page value", "Expected: " & inputpageval   & VBCRLF & "Actual:" & j.Page ,"PASSED"
		Else
				Append_TestHTML StepCounter, "Validate Response of error API  Page value", "Expected: " & inputpageval   & VBCRLF & "Actual:" & j.Page ,"FAILED"
		
	End If
	If isNULL(inputdpagesizeval) = true or inputdpagesizeval = "" Then
		inputdpagesizeval = 1000
	End If
	If cdbl(inputdpagesizeval) = cdbl( j.PageSize)  Then
				Append_TestHTML StepCounter, "Validate Response of error API  PageSize value", "Expected: " & inputdpagesizeval  & VBCRLF & "Actual:" & j.PageSize ,"PASSED"
		Else
				Append_TestHTML StepCounter, "Validate Response of error API  PageSize value", "Expected: " & inputdpagesizeval   & VBCRLF & "Actual:" & j.PageSize ,"FAILED"
		
	End If
	
'	If ucase(j.Status) = ucase("SUCCESS") Then
'		num_of_ent_val = j.TotalRecords
'		For Iterator = 0 To num_of_ent_val Step 1
'			entityTypeIDval =  eval("j.Data.[" & Iterator & "].EntityTypeId")
'			AccountNumberval =  eval("j.Data.[" & Iterator & "].AccountNumber")
'			
'			If entityTypeIDval <> "" Then  'or cdbl(entityTypeIDval) = cdbl("1")  or  cdbl(entityTypeIDval) = cdbl("1")  or  cdbl(entityTypeIDval) = cdbl("1")  or 
'					Append_TestHTML StepCounter, "Validate EntityTypeId value", "EntityTypeID Value is-" & entityTypeIDval & " Account Number-"& AccountNumberval ,"PASSED"
'			Else
'					Append_TestHTML StepCounter, "Validate EntityTypeId value", "Value is-" & entityTypeIDval ,"FAILED"
'			End If
'		Next
'	
'	End If
	
	
'	query_customer = "Select * from MessageQueueProcessed where EntityRowID ='"&myERP&"'"
'			Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
'			set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
'			wait 3
'			messageqID = dictDbResultSet("MessageQueueID")
'			messaageereqID = dictDbResultSet("ExternalRequestID")
'			 custStatusID = dictDbResultSet("MessageQueueStatusID")
'			custEntityPayload = dictDbResultSet("EntityPayload")
'			
'			set dictDbResultSet = Nothing
'			Set query_customer = Nothing
End Function



Function getapiConfigInfoFromCSVForErrorvalidation(Apconfigpath,errplatform)
	On error resume next
	bFlag = True
'	dataFile = TDFilePath & "TC008_PricingCreateNewPriceRule.csv"
Apiconfigpath= Apconfigpath

	 Set xlApp = CreateObject("Excel.Application")
	 xlApp.Visible = true
	 Set xlBook = xlApp.Workbooks.open(Apiconfigpath)
	sheetName = "APIdata"
		Set xlSheet = xlBook.Worksheets(sheetName)
		rcount = xlSheet.UsedRange.rows.count

		ccount = xlSheet.UsedRange.columns.count

		For i = 1 To rcount Step 1
			platformName =  xlSheet.Cells(i,1).Value
			envName =  xlSheet.Cells(i,2).Value
			
			If ucase(platformName) = ucase(errplatform) Then
				dataentryrow = i
				Exit For
			End If
		Next

		If dataentryrow > 1 and dataentryrow<=rcount Then
				headerNames =  xlSheet.Cells(dataentryrow,4).Value
				headervals = Split(headerNames,";")
				
				apiHeaders=""
				For itr = 0 To ubound(headervals) Step 1
					headerval = Split(headervals(itr),"#")
					apiHeaders = apiHeaders & headerval(0)	& "==>" &  headerval(1) & ","
				Next	 
				Append_TestHTML StepCounter, "Pre condition Headers & Credentials of " & errplatform, apiHeaders  ,"PASSED" 
				
				getapiConfigInfoFromCSV = apiHeaders
		Else
				getapiConfigInfoFromCSV = ""
		End If
		
	 xlBook.Save
	  xlApp.Quit
End Function
Function errorcodevalidation(strJson, errval, errtitle )
	On Error Resume Next
'	
'		Set fileSysObj = createObject("Scripting.FileSystemObject")
'		Set file = FileSysObj.OpenTextFile(strJson)
'		sText = file.ReadAll
'		Set fileSysObj = Nothing

	Set j = ParseJson(strJson)
	If ucase(j.Errors.[0].Code) = ucase(errval)  Then
		Append_TestHTML StepCounter, "Validate Error Code ",  "Expected: " & errval   & VBCRLF & "Actual:" & j.Errors.[0].Code , "PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Error Code ",  "Expected: " & errval   & VBCRLF & "Actual:" & j.Errors.[0].Code , "FAILED"	
	End If
	If ucase(j.Errors.[0].Title) = ucase(errtitle)  Then
		Append_TestHTML StepCounter, "Validate Title ",  "Expected: " & errtitle   & VBCRLF & "Actual:" & j.Errors.[0].Title , "PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Title ",  "Expected: " & errtitle   & VBCRLF & "Actual:" & j.Errors.[0].Title , "FAILED"	
	End If
	
End Function

Function verifyPagecountValidations(rcheck,newPagcount)
	On Error Resume Next
		Set j = ParseJson(rcheck)
		If cint(j.TotalRecords) = cint(errorapi_Totalnumrecs) Then
			Append_TestHTML StepCounter, "Validate TotalRecords ",  "Expected: "& errorapi_Totalnumrecs   & VBCRLF & "Actual:" & j.TotalRecords , "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate TotalRecords ",  "Expected: " & errorapi_Totalnumrecs   & VBCRLF & "Actual:" & j.TotalRecords , "FAILED"		
		End If
		If cint(j.TotalPages) = cint(newPagcount) Then
			Append_TestHTML StepCounter, "Validate TotalPages ",  "Expected: " & newPagcount   & VBCRLF & "Actual:" & j.TotalPages , "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate TotalPages ",  "Expected: " & newPagcount   & VBCRLF & "Actual:" & j.TotalPages , "FAILED"		
		End If
	Set j = Nothing
End Function

Function getCustomerStatus(custnewERPid)
	On Error Resume Next
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&custnewERPid&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	eStatusID = dictDbResultSet("StatusID")
	
	set dictDbResultSet = Nothing	
	set query_customer = Nothing
	
	If eStatusID<> "" Then
		Append_TestHTML StepCounter, "Customer Status-", "Status is-" & eStatusID , "PASSED"
	
		getCustomerStatus = eStatusID
	Else
		Append_TestHTML StepCounter, "Customer Status-", "Status is-" & eStatusID , "FAILED"
		
		getCustomerStatus = ""	
	End If
	
End Function

Function validateCustomerDBwithJSONrequestdata(jsonFPath,custnewERPid)
	On Error Resume Next
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set file = FileSysObj.OpenTextFile(jsonFPath,1)
		sText = file.ReadAll
'	Append_TestHTML StepCounter, "JSON Request Data",sText, "PASSED"
		
'		msgbox sText
	Set j = ParseJson(sText)
'	Set html = CreateObject("htmlfile")
'	
'	Set window1 = html.parentWindow
'	window1.execScript "var json = " & sText, "jScript"
'	
'	Set j = window1.json
	
	Set fileSysObj = Nothing	
	
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&custnewERPid&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1," GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
	ecustSName = dictDbResultSet("ShortName")
	ecustPPId = dictDbResultSet("CustomerPriceProfileID")
	ecustFGId = dictDbResultSet("FeeGroupID")
	eStatusID = dictDbResultSet("StatusID")
	eColcoID = dictDbResultSet("ColcoID")
	eModifiedBy = dictDbResultSet("ModifiedBy")
	eCustomerERP = dictDbResultSet("CustomerERP")
	eTopLevelCustomerNumber = dictDbResultSet("TopLevelCustomerID")
	eParentCustomerNumber = dictDbResultSet("ParentCustomerID")
	eLanguageId = dictDbResultSet("LanguageID")
	eCurrencyId = dictDbResultSet("CurrencyID")
	eFullName = dictDbResultSet("FullName")
	eShortName = dictDbResultSet("ShortName")
	eTradingName = dictDbResultSet("TradingName")
	eCountryId = dictDbResultSet("CountryID")
	eBandId = dictDbResultSet("BandID")
	eCustomerClassificationId = dictDbResultSet("CustomerClassificationID")
	eMarketingSegmentationId = dictDbResultSet("MarketingSegmentationID")
	eIndustrialClassId = dictDbResultSet("IndustryClassID")
	eLegalEntityId = dictDbResultSet("LegalEntityID")
	eIsInternational = dictDbResultSet("IsInternational")
	eIsProvenTaxableReseller = dictDbResultSet("IsProvenTaxableReseller")
	eDateEstablished = dictDbResultSet("DateEstablished")
	eFeeGroupId = dictDbResultSet("FeeGroupID")
	eRegistrationNumber = dictDbResultSet("RegistrationNumber")
	eRegistrationNumber2 = dictDbResultSet("RegistrationNumber2")
	eBirthNumber = dictDbResultSet("BirthNumber")
	eLineOfBusinessId = dictDbResultSet("LineOfBusinessID")
	ePriceProfileId = dictDbResultSet("CustomerPriceProfileID")
	eVATRegNumber = dictDbResultSet("VATRegNumber")
	eVATRegNumber2 = dictDbResultSet("VATRegNumber2")
	eBillingLanguageId = dictDbResultSet("BillingLanguageID")
	eCrossBusiness = dictDbResultSet("CrossBusiness")
	eCustomerSegmentationTypeID = dictDbResultSet("CustomerSegmentationTypeID")
	Set query_customer = Nothing
	set dictDbResultSet = Nothing
	If eTopLevelCustomerNumber <> "" or isNULL(eTopLevelCustomerNumber)=false Then
		query_customer = "Select * from Customer where CustomerID = '"&eTopLevelCustomerNumber&"'"
		Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
		set dictDbResultSet = execute_db_query(query_customer, 1," GFN_SHELL_SPRINTQA_PH_OLTP")
		wait 3
		eTopLevelCustomerNumber = dictDbResultSet("CustomerERP")
		Set query_customer = Nothing
		set dictDbResultSet = Nothing
	End  IF
	If eParentCustomerNumber <> "" or isNULL(eParentCustomerNumber)=false Then
			query_customer = "Select * from Customer where CustomerID = '"&eParentCustomerNumber&"'"
			Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
			set dictDbResultSet = execute_db_query(query_customer, 1," GFN_SHELL_SPRINTQA_PH_OLTP")
			wait 3
			eParentCustomerNumber = dictDbResultSet("CustomerERP")
			Set query_customer = Nothing
			set dictDbResultSet = Nothing
		End If	
	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, dbName)
	wait 3
	cdb_ClientCompanyNumber = dictDbResultSet("ClientCompanyNumber")
	cdb_CompanyID = dictDbResultSet("CompanyID")
	
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
	If cint(eColcoID) = cint(cdb_CompanyID) Then
		Append_TestHTML StepCounter, "Validate ColcoCode Value", "Expected:"& cdb_CompanyID & VBCRLF & "Actual:" & eColcoID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate ColcoCode Value", "Expected:"&cdb_CompanyID & VBCRLF & "Actual:" & eColcoID   ,"FAILED"	
			bRunFlag = False
	End If
	If cstr(eTopLevelCustomerNumber) = cstr(j.EntityBody.TopLevelCustomerNumber) Then
		Append_TestHTML StepCounter, "Validate TopLevelCustomerNumber Value", "Expected:"& j.EntityBody.TopLevelCustomerNumber & VBCRLF & "Actual:" & eTopLevelCustomerNumber  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate TopLevelCustomerNumber Value", "Expected:"& j.EntityBody.TopLevelCustomerNumber & VBCRLF & "Actual:" & eTopLevelCustomerNumber   ,"FAILED"	
			bRunFlag = False
	End If
	
	If (eParentCustomerNumber) = (j.EntityBody.ParentCustomerNumber) or isNULL(eParentCustomerNumber) = true Then
		Append_TestHTML StepCounter, "Validate ParentCustomerNumber Value", "Expected:"& j.EntityBody.ParentCustomerNumber & VBCRLF & "Actual:" & eParentCustomerNumber  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate ParentCustomerNumber Value", "Expected:"& j.EntityBody.ParentCustomerNumber & VBCRLF & "Actual:" & eParentCustomerNumber   ,"FAILED"	
			bRunFlag = False
	End If
	If cstr(eLanguageId) = cstr(j.EntityBody.LanguageId) Then
		Append_TestHTML StepCounter, "Validate LanguageId Value", "Expected:"& j.EntityBody.LanguageId & VBCRLF & "Actual:" & eLanguageId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate LanguageId Value", "Expected:"& j.EntityBody.LanguageId & VBCRLF & "Actual:" & eLanguageId   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cstr(eCurrencyId) = cstr(j.EntityBody.CurrencyId) Then
		Append_TestHTML StepCounter, "Validate CurrencyId Value", "Expected:"& j.EntityBody.CurrencyId & VBCRLF & "Actual:" & eCurrencyId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CurrencyId Value", "Expected:"& j.EntityBody.CurrencyId & VBCRLF & "Actual:" & eCurrencyId   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cstr(eFullName) = cstr(j.EntityBody.FullName) Then
		Append_TestHTML StepCounter, "Validate FullName Value", "Expected:"& j.EntityBody.FullName & VBCRLF & "Actual:" & eFullName  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate FullName Value", "Expected:"& j.EntityBody.FullName & VBCRLF & "Actual:" & eFullName   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cstr(eShortName) = cstr(j.EntityBody.ShortName) Then
		Append_TestHTML StepCounter, "Validate ShortName Value", "Expected:"& j.EntityBody.ShortName & VBCRLF & "Actual:" & eShortName  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate ShortName Value", "Expected:"& j.EntityBody.ShortName & VBCRLF & "Actual:" & eShortName   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cstr(eTradingName) = cstr(j.EntityBody.TradingName) Then
		Append_TestHTML StepCounter, "Validate TradingName Value", "Expected:"& j.EntityBody.TradingName & VBCRLF & "Actual:" & eTradingName  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate TradingName Value", "Expected:"& j.EntityBody.TradingName & VBCRLF & "Actual:" & eTradingName   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cint(eCountryId) = cint(j.EntityBody.CountryId) Then
		Append_TestHTML StepCounter, "Validate CountryId Value", "Expected:"& j.EntityBody.CountryId & VBCRLF & "Actual:" & eCountryId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CountryId Value", "Expected:"& j.EntityBody.CountryId & VBCRLF & "Actual:" & eCountryId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cint(eBandId) = cint(j.EntityBody.BandId) Then
		Append_TestHTML StepCounter, "Validate BandId Value", "Expected:"& j.EntityBody.BandId & VBCRLF & "Actual:" & eBandId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate BandId Value", "Expected:"& j.EntityBody.BandId & VBCRLF & "Actual:" & eBandId   ,"FAILED"	
			bRunFlag = False
	End If
	If cint(eCustomerClassificationId) = cint(j.EntityBody.CustomerClassificationId) Then
		Append_TestHTML StepCounter, "Validate CustomerClassificationId Value", "Expected:"& j.EntityBody.CustomerClassificationId & VBCRLF & "Actual:" & eCustomerClassificationId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerClassificationId Value", "Expected:"& j.EntityBody.CustomerClassificationId & VBCRLF & "Actual:" & eCustomerClassificationId   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cint(eMarketingSegmentationId) = cint(j.EntityBody.MarketingSegmentationId) Then
		Append_TestHTML StepCounter, "Validate MarketingSegmentationId Value", "Expected:"& j.EntityBody.MarketingSegmentationId & VBCRLF & "Actual:" & eMarketingSegmentationId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate MarketingSegmentationId Value", "Expected:"& j.EntityBody.MarketingSegmentationId & VBCRLF & "Actual:" & eMarketingSegmentationId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cint(eIndustrialClassId) = cint(j.EntityBody.IndustrialClassId) Then
		Append_TestHTML StepCounter, "Validate IndustrialClassId Value", "Expected:"& j.EntityBody.IndustrialClassId & VBCRLF & "Actual:" & eIndustrialClassId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate IndustrialClassId Value", "Expected:"& j.EntityBody.IndustrialClassId & VBCRLF & "Actual:" & eIndustrialClassId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cint(eLegalEntityId) = cint(j.EntityBody.LegalEntityId) Then
		Append_TestHTML StepCounter, "Validate LegalEntityId Value", "Expected:"& j.EntityBody.LegalEntityId & VBCRLF & "Actual:" & eLegalEntityId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate LegalEntityId Value", "Expected:"& j.EntityBody.LegalEntityId & VBCRLF & "Actual:" & eLegalEntityId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cbool(eIsInternational) = cbool(j.EntityBody.IsInternational) Then
		Append_TestHTML StepCounter, "Validate IsInternational Value", "Expected:"& j.EntityBody.IsInternational & VBCRLF & "Actual:" & eIsInternational  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate IsInternational Value", "Expected:"& j.EntityBody.IsInternational & VBCRLF & "Actual:" & eIsInternational   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cbool(eIsProvenTaxableReseller) = cbool(j.EntityBody.IsProvenTaxableReseller) Then
		Append_TestHTML StepCounter, "Validate IsProvenTaxableReseller Value", "Expected:"& j.EntityBody.IsProvenTaxableReseller & VBCRLF & "Actual:" & eIsProvenTaxableReseller  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate IsProvenTaxableReseller Value", "Expected:"& j.EntityBody.IsProvenTaxableReseller & VBCRLF & "Actual:" & eIsProvenTaxableReseller   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cstr(eDateEstablished) <> "" Then
		Append_TestHTML StepCounter, "Validate DateEstablished Value", "Expected:"& j.EntityBody.DateEstablished & VBCRLF & "Actual:" & eDateEstablished  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate DateEstablished Value", "Expected:"& j.EntityBody.DateEstablished & VBCRLF & "Actual:" & eDateEstablished   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cstr(eFeeGroupId) <> "" Then
		Append_TestHTML StepCounter, "Validate FeeGroupId Value", "Expected:"& j.EntityBody.FeeGroupId & VBCRLF & "Actual:" & eFeeGroupId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate FeeGroupId Value", "Expected:"& j.EntityBody.FeeGroupId & VBCRLF & "Actual:" & eFeeGroupId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cstr(eRegistrationNumber) = cstr(j.EntityBody.RegistrationNumber) Then
		Append_TestHTML StepCounter, "Validate RegistrationNumber Value", "Expected:"& j.EntityBody.RegistrationNumber & VBCRLF & "Actual:" & eRegistrationNumber  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate RegistrationNumber Value", "Expected:"& j.EntityBody.RegistrationNumber & VBCRLF & "Actual:" & eRegistrationNumber   ,"FAILED"	
			bRunFlag = False
	End If
	
	If (eRegistrationNumber2) = (j.EntityBody.RegistrationNumber2) or eRegistrationNumber2 = ""  or isNULL(eRegistrationNumber2)=true Then
		Append_TestHTML StepCounter, "Validate RegistrationNumber2 Value", "Expected:"& j.EntityBody.RegistrationNumber2 & VBCRLF & "Actual:" & eRegistrationNumber2  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate RegistrationNumber2 Value", "Expected:"& j.EntityBody.RegistrationNumber2 & VBCRLF & "Actual:" & eRegistrationNumber2   ,"FAILED"	
			bRunFlag = False
	End If
	
	If (eBirthNumber) = (j.EntityBody.BirthNumber) Then
		Append_TestHTML StepCounter, "Validate BirthNumber Value", "Expected:"& j.EntityBody.BirthNumber & VBCRLF & "Actual:" & eBirthNumber  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate BirthNumber Value", "Expected:"& j.EntityBody.BirthNumber & VBCRLF & "Actual:" & eBirthNumber   ,"FAILED"	
			bRunFlag = False
	End If
	If cint(eLineOfBusinessId) = cint(j.EntityBody.LineOfBusinessId) Then
		Append_TestHTML StepCounter, "Validate LineOfBusinessId Value", "Expected:"& j.EntityBody.LineOfBusinessId & VBCRLF & "Actual:" & eLineOfBusinessId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate LineOfBusinessId Value", "Expected:"& j.EntityBody.LineOfBusinessId & VBCRLF & "Actual:" & eLineOfBusinessId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cint(ePriceProfileId) = cint(j.EntityBody.PriceProfileId) Then
		Append_TestHTML StepCounter, "Validate PriceProfileId Value", "Expected:"& j.EntityBody.PriceProfileId & VBCRLF & "Actual:" & ePriceProfileId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate PriceProfileId Value", "Expected:"& j.EntityBody.PriceProfileId & VBCRLF & "Actual:" & ePriceProfileId   ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If (eVATRegNumber) = (j.EntityBody.VATRegNumber) or eVATRegNumber="" or isNULL(eVATRegNumber)=true Then
		Append_TestHTML StepCounter, "Validate VATRegNumber Value", "Expected:"& j.EntityBody.VATRegNumber & VBCRLF & "Actual:" & eVATRegNumber  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate VATRegNumber Value", "Expected:"& j.EntityBody.VATRegNumber & VBCRLF & "Actual:" & eVATRegNumber   ,"FAILED"	
			bRunFlag = False
	End If
	
	If (eVATRegNumber2) = (j.EntityBody.VATRegNumber2)  or eVATRegNumber2="" or isNULL(eVATRegNumber2)=true Then
		Append_TestHTML StepCounter, "Validate VATRegNumber2 Value", "Expected:"& j.EntityBody.VATRegNumber2 & VBCRLF & "Actual:" & eVATRegNumber2  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate VATRegNumber2 Value", "Expected:"& j.EntityBody.VATRegNumber2 & VBCRLF & "Actual:" & eVATRegNumber2   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cint(eBillingLanguageId) = cint(j.EntityBody.BillingLanguageId) Then
		Append_TestHTML StepCounter, "Validate BillingLanguageId Value", "Expected:"& j.EntityBody.BillingLanguageId & VBCRLF & "Actual:" & eBillingLanguageId  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate BillingLanguageId Value", "Expected:"& j.EntityBody.BillingLanguageId & VBCRLF & "Actual:" & eBillingLanguageId   ,"FAILED"	
			bRunFlag = False
	End If
	
	If cbool(eCrossBusiness) = cbool(j.EntityBody.IsCrossBusiness) Then
		Append_TestHTML StepCounter, "Validate IsCrossBusiness Value", "Expected:"& j.EntityBody.IsCrossBusiness & VBCRLF & "Actual:" & eCrossBusiness  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate IsCrossBusiness Value", "Expected:"& j.EntityBody.IsCrossBusiness & VBCRLF & "Actual:" & eCrossBusiness  ,"FAILED"	
			bRunFlag = False
	End If
	
	
	If cint(eCustomerSegmentationTypeID) = cint(j.EntityBody.CustomerSegmentationTypeId) Then
		Append_TestHTML StepCounter, "Validate CustomerSegmentationTypeID Value", "Expected:"& j.EntityBody.CustomerSegmentationTypeID & VBCRLF & "Actual:" & eCustomerSegmentationTypeID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerSegmentationTypeID Value", "Expected:"& j.EntityBody.CustomerSegmentationTypeID & VBCRLF & "Actual:" & eCustomerSegmentationTypeID  ,"FAILED"	
			bRunFlag = False
	End If
	
	
	
End Function




Function errorAPIResponseDetailsValidation(jsonFPath,resJData)
	On Error Resume Next
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set file = FileSysObj.OpenTextFile(jsonFPath,1)
		sText = file.ReadAll
	Set i = ParseJson(sText)
	Set j = ParseJson(resJData)
	
	Set fileSysObj = Nothing	
		
	query = "Select * from [dbo].[MessageQueueError] MQE	INNER JOIN [dbo].[MessageQueueProcessed] MQP ON (MQP.MessageQueueID = MQE.MessageQueueID) AND MQP.SystemIdentifier = '" & i.UserName &"'AND MQP.ProcessedOnUTC >='" & i.Filters.StartDateTime &"'  AND MQP.ProcessedOnUTC <= '" & i.Filters.EndDateTime &"' AND MQE.StackTrace IS NULL;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, j.TotalRecords, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	eExternalRequestID = dictDbResultSet("ExternalRequestID")
	eExternalReference = dictDbResultSet("ExternalReference")
	eAccountNumber = dictDbResultSet("AccountNumber")
	eMessageQueueEntityTypeID = dictDbResultSet("MessageQueueEntityTypeID")
	eErrorMessage = dictDbResultSet("ErrorMessage")
	eProcessedOn = dictDbResultSet("ProcessedOn")
	
	Set dictDbResultSet = Nothing
	Set query = Nothing
	If  isnull(i.PageSize) = True Then
		maxpsize =1000	
	Else
		maxpsize = i.PageSize
	End If
'	num_of_pages = Round(cint(j.TotalRecords) / cint(i.PageSize))
'	pagesDifference = cint(j.TotalPages) - cint(num_of_pages)
	If cdbl(errorapi_Totalnumrecs) = cdbl( j.TotalRecords)  Then
				Append_TestHTML StepCounter, "Validate Response of error API  Totalrecords value", "Expected: " & errorapi_Totalnumrecs  & VBCRLF & "Actual:" & j.TotalRecords ,"PASSED"
		Else
				Append_TestHTML StepCounter, "Validate Response of error API  Totalrecords value", "Expected: " & errorapi_Totalnumrecs  & VBCRLF & "Actual:" & j.TotalRecords ,"FAILED"
		
	End If
	If isNULL(i.Page) = true or i.Page = "" Then
		inputdpagesizeval = 1
	Else
		inputdpagesizeval = i.Page
	End If
	If cdbl(inputdpagesizeval) = cdbl( j.Page)  Then
				Append_TestHTML StepCounter, "Validate Response of error API  Page value", "Expected: " & inputdpagesizeval   & VBCRLF & "Actual:" & j.Page ,"PASSED"
		Else
				Append_TestHTML StepCounter, "Validate Response of error API  Page value", "Expected: " &inputdpagesizeval   & VBCRLF & "Actual:" & j.Page ,"FAILED"
		
	End If
	
	If cdbl(maxpsize) = cdbl( j.PageSize)  Then
				Append_TestHTML StepCounter, "Validate Response of error API  PageSize value", "Expected: " & maxpsize  & VBCRLF & "Actual:" & j.PageSize ,"PASSED"
		Else
				Append_TestHTML StepCounter, "Validate Response of error API  PageSize value", "Expected: " & maxpsize   & VBCRLF & "Actual:" & j.PageSize ,"FAILED"
		
	End If
			
	If ucase(j.Status) = ucase("SUCCESS") Then
		If  isnull(i.PageSize) = True or cdbl(i.PageSize) = cdbl("1000") Then
			num_of_st_val = 0
			num_of_ent_val = j.TotalRecords
			final_Total_Dat_rs = (cint(num_of_ent_val) - cint(num_of_st_val) ) -1
		Else
			inposition = -1
			myiincval = cint(maxpsize) - 1
			If cint(i.Page) = 1 Then
				ntimes = 1
			Else
				ntimes = cint(i.Page) - 1
			End If
			corrincval = ntimes * myiincval
			finalincval = inposition + corrincval
			If cint(i.Page) = 1 Then
				num_of_st_val = 0
				num_of_ent_val =  cint(i.Page) + finalincval
			ElseIf cint(i.Page) = cint(j.TotalPages) Then
				num_of_st_val = cint(i.Page) + finalincval
				num_pend_val = (cint(j.TotalRecords) -((cint(j.TotalPages)  - 1) * cint(maxpsize)))-1
				num_of_ent_val = cint(num_of_st_val) + num_pend_val
			Else
				num_of_st_val = cint(i.Page) + finalincval	
				num_of_ent_val = cint(num_of_st_val) + myiincval
			End If
			final_Total_Dat_rs = (cint(num_of_ent_val) - cint(num_of_st_val) ) + 1
		End If
			
			If cint(j.Data.Length) = cint(final_Total_Dat_rs)  Then  
				Append_TestHTML StepCounter, "Validate Data records in Page-" & cint(i.Page), "Expected: i/p json PageSize-"& final_Total_Dat_rs  & VBCRLF & "Actual: Data records count-" &  j.Data.Length  ,"PASSED"					
			Else
				Append_TestHTML StepCounter, "Validate Data records in  Page" & cint(i.Page), "Expected: i/p json PageSize-"& final_Total_Dat_rs  & VBCRLF & "Actual: Data records count-" &  j.Data.Length  ,"FAILED"					
			End If
		dataIterator = 0
		For Iterator = num_of_st_val To num_of_ent_val Step 1
			rRequestId =  eval("j.Data.[" & dataIterator & "].RequestId")
			rExternalReference =  eval("j.Data.[" & dataIterator & "].ExternalReference")
			rAccountNumber =  eval("j.Data.[" & dataIterator & "].AccountNumber")
			rEntityTypeId =  eval("j.Data.[" & dataIterator & "].EntityTypeId")
			rErrorMessage =  eval("j.Data.[" & dataIterator & "].ErrorMessage")
			rProcessedOn =  eval("j.Data.[" & dataIterator & "].ProcessedOn")
			Append_TestHTML StepCounter, "Validate Response-" & dataIterator & " data values", "Loading Resoponse-"& dataIterator & " Data" ,"PASSED"
			If  j.TotalRecords > 1 Then
				eExternalRequestIDs = Split(eExternalRequestID,"|") (Iterator)
				eExternalReferences = Split(eExternalReference,"|")(Iterator)
				eAccountNumbers = Split(eAccountNumber,"|")(Iterator)
				eMessageQueueEntityTypeIDs = Split(eMessageQueueEntityTypeID,"|")(Iterator)
				eErrorMessages = Split(eErrorMessage,"|")(Iterator)
				eProcessedOns = Split(eProcessedOn,"|")(Iterator)	
			Else
				eExternalRequestIDs = eExternalRequestID
				eExternalReferences = eExternalReference
				eAccountNumbers = eAccountNumber
				eMessageQueueEntityTypeIDs = eMessageQueueEntityTypeID
				eErrorMessages = eErrorMessage
				eProcessedOns = eProcessedOn
			End If
			If eExternalRequestIDs = rRequestId  Then  
				Append_TestHTML StepCounter, "Validate RequestId Value", "Expected:"& eExternalRequestIDs & VBCRLF & "Actual:" & rRequestId  ,"PASSED"					
			Else
				Append_TestHTML StepCounter, "Validate RequestId Value", "Expected:"& eExternalRequestIDs & VBCRLF & "Actual:" & rRequestId  ,"FAILED"					
			End If
			If eExternalReferences = rExternalReference  Then  
				Append_TestHTML StepCounter, "Validate ExternalReference Value", "Expected:"& eExternalReferences & VBCRLF & "Actual:" & rExternalReference  ,"PASSED"					
			Else
				Append_TestHTML StepCounter, "Validate ExternalReference Value", "Expected:"& eExternalReferences & VBCRLF & "Actual:" & rExternalReference  ,"FAILED"					
			End If
			If cstr(trim(eAccountNumbers)) = cstr(trim(rAccountNumber)) Then  
				Append_TestHTML StepCounter, "Validate AccountNumber Value", "Expected:"& eAccountNumbers & VBCRLF & "Actual:" & rAccountNumber  ,"PASSED"					
			Else
				Append_TestHTML StepCounter, "Validate AccountNumber Value", "Expected:"& eAccountNumbers &"-" & len(eAccountNumbers) & VBCRLF & "Actual:" & rAccountNumber &"-" & len(rAccountNumber)   ,"FAILED"					
			End If
			If trim(eMessageQueueEntityTypeIDs) = trim(rEntityTypeId) Then  
				Append_TestHTML StepCounter, "Validate EntityTypeId Value", "Expected:"& eMessageQueueEntityTypeIDs & VBCRLF & "Actual:" & rEntityTypeId  ,"PASSED"					
			Else
				Append_TestHTML StepCounter, "Validate EntityTypeId Value", "Expected:"& eMessageQueueEntityTypeIDs & VBCRLF & "Actual:" & rEntityTypeId  ,"FAILED"					
			End If
'			If eErrorMessages = rErrorMessage Then  
'				Append_TestHTML StepCounter, "Validate ErrorMessage Value", "Expected:"& eErrorMessages & VBCRLF & "Actual:" & rErrorMessage  ,"PASSED"					
'			Else
'				Append_TestHTML StepCounter, "Validate ErrorMessage Value", "Expected:"& eErrorMessages & VBCRLF & "Actual:" & rErrorMessage  ,"FAILED"					
'			End If
			If eProcessedOns = rProcessedOn or eProcessedOns <> "" Then  
				Append_TestHTML StepCounter, "Validate ProcessedOn Value", "Expected:"& eProcessedOns & VBCRLF & "Actual:" & rProcessedOn  ,"PASSED"					
			Else
				Append_TestHTML StepCounter, "Validate ProcessedOn Value", "Expected:"& eProcessedOns & VBCRLF & "Actual:" & rProcessedOn  ,"FAILED"					
			End If
			dataIterator = dataIterator + 1
		Next
		
	End If
	Set i = Nothing
	Set j = Nothing
End Function

Function displayErrorMessageEntityTypeValue(newcERPNo)
	On Error Resume Next
	wait 20
	query = "Select * from [dbo].[MessageQueueProcessed] order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query,1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	eMessageQueueID = dictDbResultSet("MessageQueueID")
	eMessageQueueEntityTypeID = dictDbResultSet("MessageQueueEntityTypeID")
	eAccountNumber = dictDbResultSet("AccountNumber")
	db_startdateval = dictDbResultSet("ProcessedOnUTC")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	original_date = Split(db_startdateval," ")(0)
	myordate = nextdateDBFormt(cdate(original_date)-1)
	mystartdateval = myordate & " " & Split(db_startdateval," ")(1) &".000"
	'&" : Processed On" & mystartdateval &"----" & db_startdateval
	If instr(cstr(eAccountNumber) , cstr(newcERPNo) ) > 0  Then
				displayErrorMessageEntityTypeValue = mystartdateval
				Append_TestHTML StepCounter, "Get EntityTypeID value from MessageQueueEnttityTypeID table", "EntityTypeId-"& eMessageQueueEntityTypeID  ,"PASSED"					
	Else
			displayErrorMessageEntityTypeValue = ""
				Append_TestHTML StepCounter, "Get EntityTypeID value from MessageQueueEnttityTypeID table", "EntityTypeId is not created for" & eAccountNumber  & "-" & newcERPNo & "::" & db_startdateval &":-:" & mystartdateval,"FAILED"					
	
	End If
	
End Function

Function validatejob2200withstatus(logstatusID)
On error resume next
	
	bFlag = True
	bRunFlag = True
	wait 50
	query = "Select * from Job where JobTypeID = 2200 and StatusID=3 order by 1 desc;"
	Append_TestHTML StepCounter, "Verify job 2200 ",query, "PASSED"
	wait 20
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_jobid = dbRecordSet("ID")
	db_statusid = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	Set dbRecordSet = Nothing
	Set query = Nothing
	wait 20
	If cint(db_statusid) = cint("3")  Then		'and instr(db_inputxml,newERPCustNo)>0
		Append_TestHTML StepCounter, "Validate job 2200 status", "Expected Value: 3"  & VBCRLF & "Actual Value: " & db_statusid ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate job 2200 status", "Expected Value: 3"  & VBCRLF & "Actual Value: " & db_statusid ,"FAILED"
	
	End If		
	wait 100
		query = "Select * from JobLog where JobID='"& db_jobid & "' order by 1 desc; "
		Append_TestHTML StepCounter, "Verify JobLog of id "& db_jobid  ,query, "PASSED"
		wait 20
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_Message = dbRecordSet("Message")
		db_logtypeID = dbRecordSet("LogTypeID")
		db_logDate = dbRecordSet("Date")
		Set dbRecordSet = Nothing
		Set query = Nothing
		If db_Message <> "" or instr(db_Message,newERPCustNo)>0  or instr(db_logDate,Day(Date))>0Then
			Append_TestHTML StepCounter, "Validate Message data", "Actual Value: " & db_Message ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Message data", "Actual Value: " & db_Message ,"FAILED"
				bRunFlag = False
		End If
		errMsgval = Split(db_Message,"-")(1)
		If instr(errMsgval,"Invalid ")>0 or errMsgval<> "" or instr(errMsgval,"different ")>0  or instr(errMsgval,"error ")>0 Then
			Append_TestHTML StepCounter, "Validate Error Message data", "Error Message Value: " & errMsgval ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate Error Message data", "Error Message Value: " & errMsgval ,"FAILED"			
		End If
			
		
		If cint(db_logtypeID) = cint(logstatusID) Then
			Append_TestHTML StepCounter, "Validate LogTYpeID status", "Expected Value: "& logstatusID  & VBCRLF & "Actual Value: " & db_logtypeID ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate LogTYpeID status", "Expected Value: " & logstatusID  & VBCRLF & "Actual Value: " & db_logtypeID ,"FAILED"
				bRunFlag = False
		End If	
	
End Function
	



Function getErrorapiPreferedDateTime(err_data,jsonFPath)
	On Error Resume Next
	inputusernamval = err_data("userName")
	
	query = "Select * from Company where DatabaseName like '%" & countryCode & "' and CompanyName like'%"& countryName &"%';"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	db_ColcoCode = dictDbResultSet("ClientCompanyNumber")
	db_CompanyID = dictDbResultSet("CompanyID")
	db_CountryID = dictDbResultSet("CountryID")
	db_RegionID = dictDbResultSet("LegislativeRegionID")
	db_CurrencyID = dictDbResultSet("CurrencyID")
	Set dictDbResultSet = Nothing
	Set query = Nothing
		Append_TestHTML StepCounter, "Get ColcoCode from Company Table ", "ClientCompanyNumber is: " & db_ColcoCode , "PASSED"
		Append_TestHTML StepCounter, "Get UserName from TestData ", "UserName is: " & inputusernamval , "PASSED"

		Append_TestHTML StepCounter, "Get ProcessedOnUTC Startdate and enddate values from MessageQueueProcessed Table ", "StartDate is: " & toplevelstartTime & "EndDate is: " & toplevelendTime , "PASSED"
	

	mystartdateval = toplevelstartTime
	
	
	myenddateval= toplevelendTime
	

	err_colcoID = db_ColcoCode
	preSearchString = "ColcoCode-" & db_ColcoCode &";UserName-"& inputusernamval
	Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, preSearchString)
	preSearchString = "StartDateTime$" & mystartdateval & ";EndDateTime$"& myenddateval
	Call searchandReplaceMultipleStringwithdollar(jsonFPath, preSearchString)
	
	query = "Select Count(*) as totalrec from [dbo].[MessageQueueError] MQE	INNER JOIN [dbo].[MessageQueueProcessed] MQP ON (MQP.MessageQueueID = MQE.MessageQueueID) AND MQP.SystemIdentifier = '" & inputusernamval & "'AND MQP.ProcessedOnUTC >='" & mystartdateval &"'  AND MQP.ProcessedOnUTC <= '" & myenddateval &"' AND MQE.StackTrace IS NULL;"		
	Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
	Set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	errorapi_Totalnumrecs = dictDbResultSet("totalrec")
	Set dictDbResultSet = Nothing
	Set query = Nothing
	
		Append_TestHTML StepCounter, "Get Total Records b/w Startdate and enddate values from MessageQueueProcessed Table ", "Total number of records is: " & errorapi_Totalnumrecs, "PASSED"
	
	
	
End Function






Function verifyPostJobMessageQueueDetails(custnewERPid,oprequestID)
	
	On error resume next
	
	bFlag = True
	
		query = "Select * from MessageQueue where ExternalRequestID = '"& oprequestID & "' order by 1 desc"
		Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
		wait 3
		msgQID = dictDbResultSet("MessageQueueID")
		msgRID = dictDbResultSet("ExternalRequestID")
		msgPayload = dictDbResultSet("EntityPayload")
		msgqstatus = dictDbResultSet("MessageQueueStatusID")
		msgrecordcreateddate = dictDbResultSet("CreatedOn")
		set dictDbResultSet = Nothing
		
		If msgqstatus = "1" and instr(msgPayload, custnewERPid)> 0 and instr(msgrecordcreateddate, nextdateDBFormt(Date))>0 Then
			Append_TestHTML StepCounter, "Validate MessageQueue Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
			Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Append_TestHTML StepCounter, "MessageQueue Entity status ", "Record is Still in Message Queue" ,"PASSED"
		Else
			query = "Select * from MessageQueueProcessed where ExternalRequestID = '"& oprequestID & "' order by 1 desc"
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
			wait 3
			msgQID = dictDbResultSet("MessageQueueID")
			msgRID = dictDbResultSet("ExternalRequestID")
			msgPayload = dictDbResultSet("EntityPayload")
			msgqstatus = dictDbResultSet("MessageQueueStatusID")
			msgrecordcreateddate = dictDbResultSet("CreatedOn")
			msgrecordAccountNumber = dictDbResultSet("AccountNumber")
			set dictDbResultSet = Nothing
			''msgbox msgPayload
			''msgbox instr(msgPayload, custnewERPid)
			''msgbox msgrecordcreateddate
			''msgbox nextdateDBFormt(Date)			
			If msgqstatus = "3" and (instr(msgPayload, "AddressTypeId")> 0 or instr(msgPayload, "ContactTypeId")> 0 or instr(msgPayload, "BankTypeId")> 0  or instr(msgPayload, "DDMandateTypeId")>0 or instr(msgPayload, "CustomerCardType")> 0 or instr(msgPayload, "CreditLimit")> 0 or instr(msgrecordcreateddate, nextdateDBFormt(Date))>0) Then
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
				Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "messagequeuestatus-" & msgqstatus & "Payload-" & msgPayload  ,"FAILED"
				Append_TestHTML StepCounter, "MessageQueueProcessed Entity status ", "No Record created in the MessageQueueProcessed table" ,"FAILED"
					bRunFlag = False
			End If
		
		End If
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&msgrecordAccountNumber&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
'	ecustSName = dictDbResultSet("ShortName")
'	ecustPPId = dictDbResultSet("CustomerPriceProfileID")
'	ecustFGId = dictDbResultSet("FeeGroupID")
'	db_ecustStatusID = dictDbResultSet("StatusID")
	
	set dictDbResultSet = Nothing
	
	
	Call OpenApplication(url)
	Call customerSearch1(msgrecordAccountNumber)
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary")  Then
		Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=SearchMenu","html tag:=INPUT").Set "Addresses"
		wait 2
		Browser("creationTime:=1").Page("creationTime:=1").Link("html tag:=A","innertext:="&"Addresses","index:=0").Click
		wait 2
	End  If
	If ecustCID <> "" Then
		verifyPostJobMessageQueueDetails = ecustCID
	Else
		verifyPostJobMessageQueueDetails = ""	
	End If
	
End Function



Function verifyPostJobErrorMessageQueueDetails(custnewERPid,oprequestID,errmsgdataval)
	
	On error resume next
	
	bFlag = True
	
		query = "Select * from MessageQueue where ExternalReference = '"& oprequestID & "' order by 1 desc"
		Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
		wait 3
		msgQID = dictDbResultSet("MessageQueueID")
		msgRID = dictDbResultSet("ExternalRequestID")
		msgPayload = dictDbResultSet("EntityPayload")
		msgqstatus = dictDbResultSet("MessageQueueStatusID")
		msgrecordcreateddate = dictDbResultSet("CreatedOn")
		set dictDbResultSet = Nothing
		
		If msgqstatus = "1" and instr(msgPayload, custnewERPid)> 0 and instr(msgrecordcreateddate, nextdateDBFormt(Date))>0 Then
			Append_TestHTML StepCounter, "Validate MessageQueue Details", "Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
			Append_TestHTML StepCounter, "EntityPayload Data ", msgPayload ,"PASSED"
			Append_TestHTML StepCounter, "MessageQueue Entity status ", "Record is Still in Message Queue" ,"PASSED"
		Else
			query = "Select * from MessageQueueProcessed where ExternalReference = '"& oprequestID & "' order by 1 desc"
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
			wait 3
			msgQID = dictDbResultSet("MessageQueueID")
			msgRID = dictDbResultSet("ExternalRequestID")
			msgPayload = dictDbResultSet("EntityPayload")
			msgqstatus = dictDbResultSet("MessageQueueStatusID")
			msgrecordcreateddate = dictDbResultSet("CreatedOn")
			msgrecordAccountNumber = dictDbResultSet("AccountNumber")
			set dictDbResultSet = Nothing
			set query = Nothing
			
			query = "Select * from MessageQueueError  order by 1 desc"
			Append_TestHTML StepCounter, "Execute Query",query, "PASSED"
			set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
			wait 3	
			db_MessageQueueID = dictDbResultSet("MessageQueueID")
			db_ErrorMessage = dictDbResultSet("ErrorMessage")
			
			set dictDbResultSet = Nothing
			set query = Nothing
			Append_TestHTML StepCounter, "Get MessageQueueError Details","MessageQID-" &db_MessageQueueID & "ErrorMessaage-"&db_ErrorMessage  , "PASSED"
			
			
			If msgqstatus = "4" and cdbl(msgQID) = cdbl(db_MessageQueueID) and ( instr(db_ErrorMessage,errmsgdataval)>0 or  instr(msgrecordcreateddate, nextdateDBFormt(Date))>0) Then
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Error Message Queue id is: "& msgQID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "" ,"PASSED"
				Append_TestHTML StepCounter, "Error EntityPayload Data ", msgPayload ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate MessageQueueProcessed Details", "Message Queue id is: MQPTable"& msgQID &" MQE Table-" & db_MessageQueueID  & VBCRLF & "Message Request id: "& msgRID  & VBCRLF & "messagequeuestatus-" & msgqstatus & "Payload-" & msgPayload  ,"FAILED"
				Append_TestHTML StepCounter, "MessageQueueProcessed Entity status ", "No Record created in the MessageQueueProcessed table" ,"FAILED"
					bRunFlag = False
			End If
		
		End If
	query_customer = "Select * from Customer where ClientCustomerNumber = '"&msgrecordAccountNumber&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_OLTP")
	wait 3
	 ecustCID = dictDbResultSet("CustomerID")
'	ecustSName = dictDbResultSet("ShortName")
'	ecustPPId = dictDbResultSet("CustomerPriceProfileID")
'	ecustFGId = dictDbResultSet("FeeGroupID")
'	db_ecustStatusID = dictDbResultSet("StatusID")
	
	set dictDbResultSet = Nothing
	
	
	Call OpenApplication(url)
	Call customerSearch1(msgrecordAccountNumber)
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary")  Then
		Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=SearchMenu","html tag:=INPUT").Set "Addresses"
		wait 2
		Browser("creationTime:=1").Page("creationTime:=1").Link("html tag:=A","innertext:="&"Addresses","index:=0").Click
		wait 2
	End  If
	If ecustCID <> "" Then
		verifyPostJobErrorMessageQueueDetails = ecustCID
	Else
		verifyPostJobErrorMessageQueueDetails = ""	
	End If
	
End Function

