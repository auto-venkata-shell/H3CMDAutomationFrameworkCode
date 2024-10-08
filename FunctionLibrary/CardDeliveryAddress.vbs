Public saminput,transRef,transReqID,addtype_10_id,addtype_13_id,transErrDesc

'******************************* HEADER ******************************************
' Description : The function to verify json response status of input json request of POST metod type
' Creator :  Venkata Srinivasa Rao. K
' Date : 20th December, 2022
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Function invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)

On error resume next
	bFlag = True
		bRunFlag =True
	'Set objAPI = createobject("MSXML2.serverXMLHTTP.6.0")
	Set objAPI = createobject("WinHttp.WinHttpRequest.5.1")
	
	objAPI.open reqType, apiurl , asynctype
	arrHeaders = split(headers,",")
	
	For hitr = 0 To ubound(arrHeaders) Step 1
		If arrHeaders(hitr) <> "" or arrHeaders(hitr) <> null Then
			arrheaderNameval = split(arrHeaders(hitr),"==>")
'			msgbox arrheaderNameval(0)
'			msgbox arrheaderNameval(1)
			
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
	If j.Status = "Success" Then
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "PASSED"
		transRef = j.Results.TransactionRef
		transReqID = j.RequestID
		invokeAPI = pageReturn
		wait 20
	Else
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "FAILED"
		transRef = ""
		transReqID = ""
		transErrDesc = j.Description
		invokeAPI = Empty
		bRunFlag = False
	End If
	
	Set objAPI = Nothing
End Function






Function BinarytoString(byVal Binary)
On Error Resume Next
	BinarytoString = ""
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type =1
	BinaryStream.Open
	BinaryStream.Write Binary
	BinaryStream.Position = 0
	BinaryStream.Type =2
	BinaryStream.Charset = "UTF-8"
	BinarytoString = BinaryStream.ReadText
End Function



Function ParseJson(pageJson)
On Error Resume Next
	Set html = CreateObject("htmlfile")
	
	Set window1 = html.parentWindow
	window1.execScript "var json = " & pageJson, "jScript"
	
	Set ParseJson = window1.json
End Function


Function searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)
On Error Resume Next
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(jsonFPath) Then
		Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
		Newcontent = fileRead.ReadAll
		fileRead.Close
		Set fileSysObj = Nothing
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
		If strFlag Then
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
	Else
		Append_TestHTML StepCounter, "Verify json input File Path", "File not exist in the path:" & jsonFPath , "FAILED"
	End If
End Function


Function verifyPreCardActiveDetails(cardid)
	
	On error resume next
	
	bFlag = True
	
	query = "Select * from CardAddress where CardID in (" & cardid & ") and IsActive = 1 order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card Addreess query",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_addressids = dbRecordSet("AddressID")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	db_address = Split(db_addressids,"|")
	
	If db_address(0) <> "" and db_address(1) <> "" Then
		Append_TestHTML StepCounter, "Get Active Card AddressIDs", "AddressID1: " & db_address(0) & VBCRLF & "AddressID2" & db_address(1) , "PASSED"
		query = "Select * from CardAddressType where CardID in (" & cardid & ") and AddressID = " & db_address(0) &";"
		Append_TestHTML StepCounter, "Execute Card AddreessType query",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		db_addresstypeid = dbRecordSet("AddressTypeID")
		Set dbRecordSet = Nothing
		Set query = Nothing
		Append_TestHTML StepCounter, "Get AddressTypeID", "AddressID1: " & db_address(0) & VBCRLF & "AddressType- " & db_addresstypeid , "PASSED"
		If cint(db_addresstypeid) = cint("10") Then
			addtype_10_id = db_address(0)
		ElseIf cint(db_addresstypeid) = cint("13") Then
			addtype_13_id = db_address(0)
		End If
		query = "Select * from CardAddressType where CardID in (" & cardid & ") and AddressID = " & db_address(1) & ";"
		Append_TestHTML StepCounter, "Execute Card AddreessType query",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		db_addresstypeid = dbRecordSet("AddressTypeID")
		Set dbRecordSet = Nothing
		Set query = Nothing
		Append_TestHTML StepCounter, "Get AddressTypeID", "AddressID1: " & db_address(1) & VBCRLF & "AddressType- " & db_addresstypeid , "PASSED"
		If cint(db_addresstypeid) = cint("10") Then
			addtype_10_id = db_address(1)
		ElseIf cint(db_addresstypeid) = cint("13") Then
			addtype_13_id = db_address(1)
		End If
	else
		Append_TestHTML StepCounter, "Get Active Card AddressIDs", "AddressID1: " & db_address(0) & VBCRLF & "AddressID2" & db_address(1) , "FAILED"
		bFlag = False
	End If
	
End Function


Function validatejob2013()
On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeID = 2013 order by 1 desc;"
	Append_TestHTML StepCounter, "Verify job 2013 ",query, "PASSED"
	wait 10
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_jobid = dbRecordSet("ID")
	db_statusid = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If cint(db_statusid) = cint("4") and instr(db_inputxml,transRef)>0 Then
		Append_TestHTML StepCounter, "Validate job 2013 status", "Expected Value: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"PASSED"
			
		query = "Select * from GenericUpdate_WWW where GenericUpdateID = " & transRef & ";"
		Append_TestHTML StepCounter, "Verify GenericUpdate_WWW table entries ",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_updatexml = dbRecordSet("UpdateXML")
		Set dbRecordSet = Nothing
		Set query = Nothing
		If db_updatexml <> "" Then
			Append_TestHTML StepCounter, "Validate UpdateXML data", "Actual Value: " & db_updatexml ,"PASSED"
		Else
			Append_TestHTML StepCounter, "Validate UpdateXML data", "Actual Value: " & db_updatexml ,"FAILED"
		
		End If
		query = "Select * from GenericUpdateStatus_WWW where GenericUpdateID =  " & transRef & ";"
		Append_TestHTML StepCounter, "Verify GenericUpdateStatus_WWW entries ",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_creationInfoxml = dbRecordSet("CreationInfoXML")
		db_IsSuccessl = dbRecordSet("IsSuccess")
		db_statusID = dbRecordSet("StatusID")
		
		
		Set dbRecordSet = Nothing
		Set query = Nothing
		
		If db_creationInfoxml <> "" Then
			Append_TestHTML StepCounter, "Validate CreationInfoXML data", "Actual Value: " & db_creationInfoxml ,"PASSED"
			Append_TestHTML StepCounter, "Validate IsSuccess data",  "Expected Value: 1"  & VBCRLF &"Actual Value: " & db_IsSuccessl ,"PASSED"
			Append_TestHTML StepCounter, "Validate StatusID data",  "Expected Value: 4"  & VBCRLF &"Actual Value: " & db_statusID ,"PASSED"
			
		Else
			Append_TestHTML StepCounter, "Validate CreationInfoXML data", "Actual Value: " & db_creationInfoxml ,"FAILED"
		
		End If
	Else
		Append_TestHTML StepCounter, "Validate job 2013 status", "Expected Value: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"FAILED"
	
	End If
	
End Function
	


Function verifyPostCardActiveDetails(cardid)
	
	On error resume next
	
	bFlag = True
	
	query = "Select * from CardAddress where CardID in (" & cardid & ") and AddressID = " & addtype_10_id & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card Addreess query",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_add_10_status = dbRecordSet("IsActive")
	Set dbRecordSet = Nothing
	Set query = Nothing
	If cint(db_add_10_status) = cint(0) or db_add_10_status = False Then
		Append_TestHTML StepCounter, "Validate Existing Addresstype 10 card status", "Expected Value: 0"  & VBCRLF & "Actual Value: " & db_add_10_status ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Existing Addresstype 10 card status", "Expected Value: 0"  & VBCRLF & "Actual Value: " & db_add_10_status , "FAILED"
	End If
	query = "Select * from CardAddress where CardID in (" & cardid & ") and AddressID = " & addtype_13_id & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card Addreess query",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_add_13_status = dbRecordSet("IsActive")
	Set dbRecordSet = Nothing
	Set query = Nothing
	If cint(db_add_13_status) = cint(1)  or db_add_13_status = True Then
		Append_TestHTML StepCounter, "Validate Existing Addresstype 10 card status", "Expected Value: 1"  & VBCRLF & "Actual Value: " & db_add_13_status ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Existing Addresstype 10 card status", "Expected Value: 1"  & VBCRLF & "Actual Value: " & db_add_13_status , "FAILED"
	End If
	
	
End Function

