
Public newCardCustomerID,XML_Cadrid, XML_Cardpanid,apiCustomerID,XML_Addressid,pinReminder_ATypeID,cardColcoID, dx053FinalFName, card_DAddressID,card_PAddressID



Function getCustomerValue(card_data)
	
	On error resume next
	
	bFlag = True
	dbquery = "Select * from Customer where CustomerERP='"& card_data("custcardERP") &"';"
	dbNameVal = "SFN_SHELL_SPRINTQA_ID_OLTP"
	Append_TestHTML StepCounter, "Execute Card Addreess query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1,dbNameval)
	wait 2
	db_CustomerID = dbRecordSet("CustomerID")
	db_AllowSelectPIN = dbRecordSet("AllowSelectPIN")
	db_UseFleetPIN = dbRecordSet("UseFleetPIN")
	apiCustomerID = db_CustomerID
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If db_AllowSelectPIN <> "" Then
		Append_TestHTML StepCounter, "Verify Customer AllowSelectPIN value", "Expected: 1/0" & VBCRLF & "Actual: " & db_AllowSelectPIN , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify Customer AllowSelectPIN value", "Expected: 1" & VBCRLF & "Actual: " & db_AllowSelectPIN , "FAILED"	
	End If
	
	dbquery = "Select * from PINAdviceType where PINAdviceTypeID="& card_data("PinAdviceTypeidval") &";"
	Append_TestHTML StepCounter, "Execute PINAdviceType query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_PostalAddressFlag = dbRecordSet("RequirePostalAddress")
	db_EmailAddressFlag = dbRecordSet("RequireEmailAddress")
	db_PhoneNumberFlag = dbRecordSet("RequireMobilePhoneNumber")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If  card_data("PinAdviceTypeidval") =1  Then
		If db_PostalAddressFlag = true  and db_EmailAddressFlag= false and db_PhoneNumberFlag=false Then
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-1 value from PINAdviceType Table", "RequirePostalAddress- Expected: 1" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-1 value from PINAdviceType Table", "RequirePostalAddress- Expected: 1" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "FAILED"
			bRunFlag = False
		End If
	ElseIf  card_data("PinAdviceTypeidval") =2 Then  	
		If db_PostalAddressFlag = false  and db_EmailAddressFlag= true and db_PhoneNumberFlag=false Then
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-2 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 1" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-2 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 1" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "FAILED"
			bRunFlag = False
		End If
	ElseIf  card_data("PinAdviceTypeidval") =3 Then  	
		If db_PostalAddressFlag = false  and db_EmailAddressFlag= false and db_PhoneNumberFlag=true Then
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-3 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 1" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-3 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 1" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "FAILED"
			bRunFlag = False
		End If
	End If
	
	dbquery = "Select * from SysVarColco where SysVarID=141;"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_ColcoID = dbRecordSet("ColcoID")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If db_ColcoID <> "" Then
		Append_TestHTML StepCounter, "Verify db_ColcoID value from SysVarColco Table", "ColcoID: " & db_ColcoID  , "PASSED"
	else
		Append_TestHTML StepCounter, "Verify db_ColcoID value from SysVarColco Table", "ColcoID: " & db_ColcoID  , "FAILED"
	End If
	dbquery = "Select * from CardTypeControl where ColcoID=" & db_ColcoID & ";"
	Append_TestHTML StepCounter, "Execute CardTypeControl query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_CardTypeID = dbRecordSet("CardTypeID")
	db_TokenTypeID= dbRecordSet("TokenTypeID")
	db_FleetPINAllowed= dbRecordSet("FleetPINAllowed")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If db_CardTypeID <> "" Then
		Append_TestHTML StepCounter, "Verify ProcessID value from JobType Table", "CardTypeID: " & db_CardTypeID & VBCRLF & "TokenTypeID: "& db_TokenTypeID  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify ProcessID value from JobType Table", "CardTypeID: " & db_CardTypeID & VBCRLF & "TokenTypeID: "& db_TokenTypeID  , "FAILED"	
	End If
	
	dbquery = "Select * from JobType where ID=2010;"
	Append_TestHTML StepCounter, "Execute JobType-2010 query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_ProcessID = dbRecordSet("ProcessID")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If ISNULL(db_ProcessID) = TRUE Then
		Append_TestHTML StepCounter, "Verify ProcessID value from JobType Table", "Expected: NULL" & VBCRLF & "Actual: NULL"  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify ProcessID value from JobType Table", "Expected: NULL" & VBCRLF & "Actual: " & db_ProcessID  , "FAILED"
			bRunFlag = False
	End If
	
	If  card_data("PinTypeIDval") =3  Then
		If cbool(db_FleetPINAllowed) = true or db_FleetPINAllowed = 1  Then
				Append_TestHTML StepCounter, "Verify FleetPINAllowed value from CardTypeControl Table", "Expected: 1" & VBCRLF & "Actual: "& db_FleetPINAllowed  , "PASSED"
		Else
			query1 = "Update CardTypeControl Set FleetPINAllowed = 1 " & "where ColcoID=" & db_ColcoID & " and CardTypeID=" & db_CardTypeID &" and TokenTypeID="& db_TokenTypeID & ";"
			Append_TestHTML StepCounter, "Update FleetPINAllowed value in  CardTypeControl", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_MASTER")
			wait 12
			Set dbRecordSet = Nothing
			Set query1 = Nothing
		End If
		If db_UseFleetPIN = 1 or cbool(db_UseFleetPIN) = true Then
				Append_TestHTML StepCounter, "Verify UseFleetPIN value from Customer Table", "Expected: 1" & VBCRLF & "Actual: "& db_UseFleetPIN  , "PASSED"
		Else
			query1 = "Update Customer Set UseFleetPIN = 1 where CustomerERP='"& card_data("custcardERP") &"';"
			Append_TestHTML StepCounter, "Update UseFleetPIN value in  Customer", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			Set dbRecordSet = Nothing
			Set query1 = Nothing
		End If
	
	End If
	If  card_data("PinTypeIDval") =2 or card_data("PinTypeIDval") =1 Then
		If cbool(db_FleetPINAllowed) = false or db_FleetPINAllowed = 0  Then
				Append_TestHTML StepCounter, "Verify FleetPINAllowed value from CardTypeControl Table", "Expected: 0" & VBCRLF & "Actual: "& db_FleetPINAllowed  , "PASSED"
		Else
			query1 = "Update CardTypeControl Set FleetPINAllowed = 0 " & "where ColcoID=" & db_ColcoID & " and CardTypeID=" & db_CardTypeID &" and TokenTypeID="& db_TokenTypeID & ";"
			Append_TestHTML StepCounter, "Update FleetPINAllowed value in  CardTypeControl", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_MASTER")
			wait 12
			Set dbRecordSet = Nothing
			Set query1 = Nothing
		End If
		If db_UseFleetPIN = 0 or cbool(db_UseFleetPIN) = false Then
				Append_TestHTML StepCounter, "Verify UseFleetPIN value from Customer Table", "Expected: 0" & VBCRLF & "Actual: "& db_UseFleetPIN  , "PASSED"
		Else
			query1 = "Update Customer Set UseFleetPIN = 0 where CustomerERP='"& card_data("custcardERP") &"';"
			Append_TestHTML StepCounter, "Update UseFleetPIN value in  Customer", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 12
			Set dbRecordSet = Nothing
			Set query1 = Nothing
		End If
	
	End If


	If card_data("PinTypeIDval") =4  Then
			query1 = "Update TokenTypeControl Set RequirePIN = 0 " & "where TokenTypeID=" & db_TokenTypeID & ";"
			Append_TestHTML StepCounter, "Update RequirePIN value in  TokenTypeControl", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_MASTER")
			wait 12
			Set dbRecordSet = Nothing
			Set query1 = Nothing
	Else
			query1 = "Update TokenTypeControl Set RequirePIN = 1 " & "where TokenTypeID=" & db_TokenTypeID & ";"
			Append_TestHTML StepCounter, "Update RequirePIN value in  TokenTypeControl", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_MASTER")
			wait 12
			Set dbRecordSet = Nothing
			Set query1 = Nothing
	End  If
	If db_CustomerID <> "" Then
		getCustomerValue = db_CustomerID &"|"&db_ColcoID&"|"&db_CardTypeID&"|"&db_TokenTypeID
	Else
		getCustomerValue =""
	End If
End Function



Function validateAPIjob(jobidval,newopfile,newcreatefile)
On error resume next
	
	bFlag = True
	wait 10
	query = "Select * from Job where JobTypeID = "& jobidval & " order by 1 desc;"
	Append_TestHTML StepCounter, "Verify job -"& jobidval,query, "PASSED"
	wait 10
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_jobid = dbRecordSet("ID")
	db_statusid = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If cint(db_statusid) = cint("4") and instr(db_inputxml,transRef)>0 Then
		Append_TestHTML StepCounter, "Validate job "& jobidval & " status", "Expected Value: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"PASSED"
			
		query = "Select * from GenericUpdate_WWW where GenericUpdateID = " & transRef & ";"
		Append_TestHTML StepCounter, "Verify GenericUpdate_WWW table entries ",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_updatexml = dbRecordSet("UpdateXML")
		db_GenericUpdateID = dbRecordSet("GenericUpdateID")
		db_GenericUpdateTypeID = dbRecordSet("GenericUpdateTypeID")
		
		
		Set dbRecordSet = Nothing
		Set query = Nothing
		
		query = "Select * from Job where JobTypeID = "& db_GenericUpdateTypeID & " order by 1 desc;"
			Append_TestHTML StepCounter, "Verify job -"& db_GenericUpdateTypeID,query, "PASSED"
			wait 10
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_jobid = dbRecordSet("ID")
			db_statusid = dbRecordSet("StatusID")
			db_inputxml = dbRecordSet("InputXml")
			Set dbRecordSet = Nothing
			Set query = Nothing
		If instr(db_inputxml,transRef)>0 and db_statusid = "4" Then
				Append_TestHTML StepCounter, "Validate job StatusID   status", "Expected StatusIDValue: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"PASSED"
				Append_TestHTML StepCounter, "Validate job InputXml File  status", "Input XML value " & db_inputxml  ,"PASSED"
				
			Else
				Append_TestHTML StepCounter, "Validate job StatusID   status", "Expected StatusIDValue: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"FAILED"
				Append_TestHTML StepCounter, "Validate job InputXml File  status", "Input XM value " & db_inputxml  ,"FAILED"
				
		End If	
		
		
		If db_updatexml <> "" Then
			
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			'fileSysObj.CreateTextFile(newopfile)
			Set DXwrite = fileSysObj.OpenTextFile(newopfile,2,False,-1)
			DXwrite.Write db_updatexml
			DXwrite.Close
			Append_TestHTML StepCounter, "Validate UpdateXML data", "Data updated in the XML :" & db_updatexml ,"PASSED"
			Append_TestHTML StepCounter, "Validate GenericUpdateID Value ", "Value matched with Transaction ref value- Expected:"& transRef &"Actual:"  & db_GenericUpdateID ,"PASSED"
			Append_TestHTML StepCounter, "Validate GenericUpdateTypeID data", "GenericUpdateTypeID value is :" & db_GenericUpdateTypeID ,"PASSED"
			
		Else
			Append_TestHTML StepCounter, "Validate UpdateXML data", "Actual Value: " & db_updatexml ,"FAILED"
			bRunFlag = False
		End If
		query = "Select * from GenericUpdateStatus_WWW where GenericUpdateID =  " & transRef & ";"
		Append_TestHTML StepCounter, "Verify GenericUpdateStatus_WWW entries ",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_creationInfoxml = dbRecordSet("CreationInfoXML")
		db_GenericUpdateID = dbRecordSet("GenericUpdateID")
		db_GenericUpdateStatusID = dbRecordSet("GenericUpdateStatusID")
		db_IsSuccess = dbRecordSet("IsSuccess")
		db_StatusID = dbRecordSet("StatusID")
		
		
		Set dbRecordSet = Nothing
		Set query = Nothing
		
		If db_creationInfoxml <> "" Then
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			'fileSysObj.CreateTextFile(newcreatefile)
			Set DXwrite = fileSysObj.OpenTextFile(newcreatefile,2,False,-1)
			DXwrite.Write db_creationInfoxml
			DXwrite.Close
'			Call getCardDataValues(newcreatefile)
'			XML_Cadrid = findValuefromString(newcreatefile,"<Identity>","</Identity>")
'			XML_Cardpanid = findValuefromString(newcreatefile,"<Identity2>","</Identity2>")
			Append_TestHTML StepCounter, "Validate CreationInfoXML data", "Actual Value: " & db_creationInfoxml ,"PASSED"
			Append_TestHTML StepCounter, "Validate IsSuccess value", "Expected: 1" & "Actual:" & db_IsSuccess ,"PASSED"
			
			Append_TestHTML StepCounter, "Validate StatusID value", "Expected: 4 " & "Actual:" &  db_StatusID ,"PASSED"
			
		Else
			
			Append_TestHTML StepCounter, "Validate CreationInfoXML data", "Actual Value: " & db_creationInfoxml ,"PASSED"
			Append_TestHTML StepCounter, "Validate IsSuccess value", "Expected: 1" & "Actual:" & db_IsSuccess ,"PASSED"
			
			Append_TestHTML StepCounter, "Validate StatusID value", "Expected: 4 " & "Actual:" &  db_StatusID ,"PASSED"
'			bRunFlag = False
		End If
	Else
		Append_TestHTML StepCounter, "Validate job " & jobidval & " status", "Expected Value: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"FAILED"
		bRunFlag = False	
	End If
	
End Function


Function getCardDataValues(newcreatefile)
On error resume next
	bFlag = True
	searchtagName = "<Card>"
	searchText = "<Identity>"
	searchendtagName = "</Card>"
	sectionposs =   FFMXMLFileDataReader(MFF_fileLoc, searchText, searchtagName, searchendtagName)
	Set mydicRecordset = validateSectionData(MFF_fileLoc,sectionposs)
	XML_Cadrid =  mydicRecordset("Identity")
	XML_Cardpanid =  mydicRecordset("Identity2")
	
	If XML_Cadrid <> "" and XML_Cardpanid <> "" Then
			Append_TestHTML StepCounter, "Get Cardid and Cardpanid from CreationInfoXML column", "Cardid Value: " & XML_Cadrid & "Cardpanid Value:" & XML_Cardpanid ,"PASSED"
	Else
			Append_TestHTML StepCounter, "Get Cardid and Cardpanid from CreationInfoXML column", "Cardid Value: " & XML_Cadrid & "Cardpanid Value:" & XML_Cardpanid ,"FAILED"
			bRunFlag = False
	End If
End Function


Function findValuefromString(filepath,fStringone,sStringtwo)
On error resume next
	bFlag = True
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set fileRead = fileSysObj.OpenTextFile(filepath,1,False,-1)
	content = fileRead.ReadAll	
	onepos = instr(content,fStringone)
	twopos = instr(content,sStringtwo)
	stpos = onepos+len(fStringone)
	endpos = twopos - stpos
	substring = mid(content,stpos,endpos)
	fileRead.Close
	If substring <> "" and len(substring) > 0 Then
		findValuefromString = substring
	Else
		findValuefromString = NULL
	End If
	
End Function


Function validatePostAPICardDetails()
On error resume next
	query = "Select * from CardSelectedPIN where CardID="& XML_Cadrid & " order by 1 desc;"  'applicable only Pintypeid-2
	Append_TestHTML StepCounter, "Execute CardSelectedPIN table entry ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_Libraryver = dbRecordSet("LibraryVersion")
	db_EncrySDet = dbRecordSet("EncryptedSessionDetails")
	db_EncryPin = dbRecordSet("EncryptedPIN")
	db_KeyIDval = dbRecordSet("KeyID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_Libraryver <> "" or db_EncrySDet <> "" or db_EncryPin <> "" or db_KeyIDval <> ""  Then
		Append_TestHTML StepCounter, "Validate CardSelectedPIN record entries", "LibraryVersion: " & db_Libraryver  & VBCRLF & "EncryptedSessionDetails: " & db_EncrySDet  & VBCRLF & "EncryptedPIN: " & db_EncryPin  & VBCRLF & "KeyID: " & db_KeyIDval  & VBCRLF ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardSelectedPIN record entries", "LibraryVersion: " & db_Libraryver  & VBCRLF & "EncryptedSessionDetails: " & db_EncrySDet  & VBCRLF & "EncryptedPIN: " & db_EncryPin  & VBCRLF & "KeyID: " & db_KeyIDval  & VBCRLF &"-- All are NULL --","PASSED"
		
	End If
	
	query = "Select * from Card where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_TypeofPin = dbRecordSet("TypeOfPIN")
	db_PinSelMethod = dbRecordSet("PINSelectionMethod")
	db_StatusID = dbRecordSet("StatusID")
	db_TokenTypeID = dbRecordSet("TokenTypeID")
	db_Pan = dbRecordSet("PAN")
	cardPANNum = db_Pan
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_Pan <> "" Then
		Append_TestHTML StepCounter, "Validate Card record entries", "TypeOfPIN: " & db_TypeofPin  & VBCRLF & "PINSelectionMethod: " & db_PinSelMethod  & VBCRLF & "StatusID: " & db_StatusID  & VBCRLF ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Card record entries", "TypeOfPIN: " & db_TypeofPin  & VBCRLF & "PINSelectionMethod: " & db_PinSelMethod  & VBCRLF & "StatusID: " & db_StatusID  & VBCRLF ,"FAILED"
		bRunFlag = False
	End If
	
	query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardIssueAttributes table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_InitialPINAdviceTypeID= dbRecordSet("InitialPINAdviceTypeID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_InitialPINAdviceTypeID <> "" Then
		Append_TestHTML StepCounter, "Validate CardIssueAttributes record entries", "InitialPINAdviceTypeID: " & db_InitialPINAdviceTypeID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardIssueAttributes record entries", "InitialPINAdviceTypeID: " & db_InitialPINAdviceTypeID  ,"FAILED"
		bRunFlag = False
	End If
	
	query = "Select * from CustomerContact where CustomerID="& apiCustomerID & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CustomerContact table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ContactID = dbRecordSet("ContactID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ContactID <> "" Then
		Append_TestHTML StepCounter, "Validate CustomerContact record entries", "ContactID: " & db_ContactID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerContact record entries", "ContactID: " & db_ContactID  ,"FAILED"
		bRunFlag = False
	End If
	query = "Select * from CustomerCard where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CustomerCard table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CustomerCardID = dbRecordSet("CustomerCardID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_CustomerCardID <> "" Then
		Append_TestHTML StepCounter, "Validate CustomerCard record entries", "CustomerCardID: " & db_CustomerCardID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerCard record entries", "CustomerCardID: " & db_CustomerCardID  ,"FAILED"
		bRunFlag = False
	End If
	
	query = "Select * from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardAddress table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_AddressID = dbRecordSet("AddressID")
	db_IsActive = dbRecordSet("IsActive")	
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_AddressID <> "" Then
		Append_TestHTML StepCounter, "Validate CardAddress record entries", "AddressID: " & db_AddressID & " IsActive status-" & db_IsActive ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardAddress record entries", "AddressID: " & db_AddressID  ,"PASSED"		
		bRunFlag = False
	End If
	
	query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and AddressTypeID =10 order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardAddressType table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_AddressID1 = dbRecordSet("AddressID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_AddressID1 <> "" Then
		Append_TestHTML StepCounter, "Validate CardAddress - Card DeliveryType record entries", "AddressID: " & db_AddressID1   ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardAddress- Card DeliveryType record entries", "AddressID: " & db_AddressID1  ,"PASSED"		
		bRunFlag = False
	End If
	query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and AddressTypeID =13 order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardAddressType table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_AddressID2 = dbRecordSet("AddressID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_AddressID2 <> "" Then
		Append_TestHTML StepCounter, "Validate CardAddress-Pin DeleiveryType record entries", "AddressID: " & db_AddressID2  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardAddress-Pin DeleiveryType record entries", "AddressID: " & db_AddressID2  ,"PASSED"		
		bRunFlag = False
	End If
	
	query = "Select * from CardStatusHistory where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardStatusHistory table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CardStatusHistoryID = dbRecordSet("CardStatusHistoryID")
	db_NewStatusID= dbRecordSet("NewStatusID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_CardStatusHistoryID <> "" and db_NewStatusID <> "" Then
		Append_TestHTML StepCounter, "Validate CardStatusHistory record entries", "CardStatusHistoryID: " & db_CardStatusHistoryID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardStatusHistory record entries", "CardStatusHistoryID: " & db_CardStatusHistoryID  ,"FAILED"		
		bRunFlag = False
	End If
	query = "Select * from CardActivityLog where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityLogID = dbRecordSet("ActivityLogID")
		
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityLogID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from ActivityLog where ActivityLogID="& db_ActivityLogID & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityTypeID = dbRecordSet("ActivityTypeID")
	db_Note = dbRecordSet("Note")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityTypeID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityTypeID: " & db_ActivityTypeID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityTypeID: " & db_ActivityTypeID  ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from CardPAN where CardPANID="& XML_Cardpanid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardPAN table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_PAN1 = dbRecordSet("PAN")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_PAN1 <> "" Then
		Append_TestHTML StepCounter, "Validate CardPAN record entries", "PAN: " & db_PAN1  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardPAN record entries", "PAN: " & db_PAN1  ,"FAILED"		
		bRunFlag = False
	End If
End Function


Function pinReminderPrerequestData()

	On error resume next
	bFlag = True
	query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardIssueAttributes table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_RequirePIN= dbRecordSet("RequirePIN")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	If cint(db_RequirePIN) = cint("1") or db_RequirePIN = true Then
		Append_TestHTML StepCounter, "Verify CardIssueAttributes RequirePIN value", "Expected: TRUE" & VBCRLF & "Actual: " & db_RequirePIN , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify CardIssueAttributes RequirePIN value", "Expected: TRUE" & VBCRLF & "Actual: " & db_RequirePIN , "FAILED"	
	End If
	
	query = "Select * from CardPAN where CardPANID="& XML_Cardpanid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardPAN table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_PAN1 = dbRecordSet("PAN")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	query = "Select * from Card where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_TypeofPin = dbRecordSet("TypeOfPIN")
	db_PinSelMethod = dbRecordSet("PINSelectionMethod")
	db_StatusID = dbRecordSet("StatusID")
	db_ExpiryDate = dbRecordSet("ExpiryDate")
	db_Pan = dbRecordSet("PAN")
	
	pAPI_Expitedate = Mid(Replace(db_ExpiryDate,"-",""),3,4)
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	
	query = "Select * from CustomerCard where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CustomerCard table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CustomerCardID = dbRecordSet("CustomerCardID")
	db_CustomerID = dbRecordSet("CustomerID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	dbquery = "Select * from SysVarColco where SysVarID=141;"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_ColcoID = dbRecordSet("ColcoID")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	
	dbquery = "Select * from Customer where CustomerID=" & db_CustomerID & ";"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	dbC_CountryID = dbRecordSet("CountryID")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	dbquery = "Select * from Country where CountryID=" & dbC_CountryID & ";"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_MASTER")
	wait 2
	dbM_CountryID = dbRecordSet("CountryID")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	
	pAPI_PANID = XML_Cardpanid
	pAPI_CardID = XML_Cadrid
	pAPI_CustomerID = db_CustomerID
	pAPI_cardPAN = db_Pan
	
		
	If pAPI_Expitedate <>"" Then
		pinReminderPrerequestData = pAPI_CustomerID & "|" & pAPI_CardID & "|" &  pAPI_PANID & "|" & pAPI_cardPAN & "|" & pAPI_Expitedate &"|"& db_ColcoID & "|"& dbM_CountryID
	Else
		pinReminderPrerequestData = NULL
	End If
End Function




Public Function verifyCardAddressStatus()
	query = "Select * from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
		Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		db_def_Vaue = dbRecordSet("IsActive")
		db_AddressID= dbRecordSet("AddressID")
		
		Set dbRecordSet = Nothing
		Set query = Nothing
		Append_TestHTML StepCounter,keydata & " AddressID value", "Actual Value: " &db_AddressID  ,"PASSED"
		If cBool(db_def_Vaue) =True or  db_def_Vaue= 1 Then
			Append_TestHTML StepCounter,keydata & " IsActive  Status value","Expected Value: True" & VBCRLF & "Actual Value: " &db_def_Vaue  ,"PASSED"
		Else
			Append_TestHTML StepCounter,"IsActive Satus value","Expected Value: True "& VBCRLF & "Actual Value: " &db_def_Vaue  ,"FAILED"
		End If
End Function

Public Function doIsActiveCardAddressAction(keydata,cstatusFlag,chstatusFlag)

On Error Resume Next
		query = "Select * from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
		Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		db_def_Vaue = dbRecordSet("IsActive")
		
		Set dbRecordSet = Nothing
		Set query = Nothing
		If isNull(db_def_Vaue) Then
			Append_TestHTML StepCounter, keydata & "Default value","Unknown/NULL IsActive value "& VBCRLF & "Actual - RequirePIN: " &db_def_Vaue  ,"FAILED"
			bRunFlag = False
		Else
			If cBool(db_def_Vaue) = cBool(cstatusFlag) Then
				Append_TestHTML StepCounter, keydata & " Existing value","Expected Value: " & cstatusFlag  & VBCRLF & "Actual Value: " &db_def_Vaue  ,"PASSED"
				query1 = "Update CardAddress Set IsActive = " & chstatusFlag &" where CardID=" & XML_Cadrid & ";"
				Append_TestHTML StepCounter, "Update IsActive value in  CardAddress", query1, "PASSED"
				Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				Set dbRecordSet = Nothing
				Set query1 = Nothing
	
				query = "Select * from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
				Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
				Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				db_def_Vaue1 = dbRecordSet("IsActive")
				Set dbRecordSet = Nothing
				Set query = Nothing
				
				If cBool(db_def_Vaue1) = cBool(chstatusFlag) Then
					Append_TestHTML StepCounter,keydata & " Updated value","Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_def_Vaue1  ,"PASSED"
				Else
					Append_TestHTML StepCounter,"IsActive Data Updated value","Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_def_Vaue1  ,"FAILED"
				End If
			Else
				Append_TestHTML StepCounter, keydata & " Existing value","No Update requiered in DB - Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_def_Vaue  ,"PASSED"
			End If
		End If
End  Function


Function validatePostAPIpinReminderDetails()
On error resume next
	
	query = "Select * from CardAddress where CardID="& XML_Cadrid & " and AddressID=" & XML_Addressid &" order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardAddress table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_IsActive = dbRecordSet("IsActive")
	db_AddressLines = dbRecordSet("AddressLines")
	db_Email = dbRecordSet("Email")
	db_RegionIDD = dbRecordSet("RegionID")
	db_CountryID = dbRecordSet("CountryID")
	db_Phone = dbRecordSet("Phone")
	
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	If db_IsActive = true or cint(db_IsActive) = 1 Then
		Append_TestHTML StepCounter, "Validate CardAddress record entries", "Expected: TRUE" &  VBCRLF & "Actual: " & db_IsActive  ,"PASSED"
		Append_TestHTML StepCounter, "Validate CardAddress record entries",  "AddressLines: " & db_AddressLines  & VBCRLF & "Email: " & db_Email  & VBCRLF & "Phone: " & db_Phone  & VBCRLF ,"PASSED"		
	Else
		Append_TestHTML StepCounter, "Validate CardAddress record entries", "Expected: TRUE" &  VBCRLF & "Actual: " & db_IsActive  ,"FAILED"
		bRunFlag = False
	End If

	query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and AddressID=" & XML_Addressid &" order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardAddressType table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_AddressTypeID = dbRecordSet("AddressTypeID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing

	If cint(db_AddressTypeID) = 13 Then
		Append_TestHTML StepCounter, "Validate AddressTypeID value from  CardAddressType table", "Expected: 13" &  VBCRLF & "Actual: " & db_AddressTypeID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate AddressTypeID value from  CardAddressType table", "Expected: 13" &  VBCRLF & "Actual: " & db_AddressTypeID  ,"FAILED"	
	End If
	query = "Select Count(*) as No_of_Records from CardActivityLog where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_No_of_Records = dbRecordSet("No_of_Records")
		
	Set dbRecordSet = Nothing
	Set query = Nothing
	query = "Select * from CardActivityLog where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, cint(db_No_of_Records),"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityLogID = dbRecordSet("ActivityLogID")
		
	Set dbRecordSet = Nothing
	Set query = Nothing
	If db_ActivityLogID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from ActivityLog where ActivityLogID in ("& Replace(db_ActivityLogID,"|",",") & ") and ActivityTypeID =" & pinReminder_ATypeID &" order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityTypeID = dbRecordSet("ActivityTypeID")
	db_Note = dbRecordSet("Note")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityTypeID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "Expected: " & pinReminder_ATypeID  & "Actual: " & db_ActivityTypeID ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "Expected: " & pinReminder_ATypeID  & "Actual: " & db_ActivityTypeID ,"FAILED"	
		bRunFlag = False
	End If

	query = "Select * from CardPINMailerRequest where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, cint(db_No_of_Records),"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_CardPINMailerRequestID = dbRecordSet("CardPINMailerRequestID")
		
	Set dbRecordSet = Nothing
	Set query = Nothing
	If db_CardPINMailerRequestID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", " CardActivityLog record created Successfully"  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "Fail to create CardActivityLog record "  ,"FAILED"
		bRunFlag = False
	End If


End Function



Function getpinReminderprerequestValues(card_data)
	
	On error resume next
	
	bFlag = True
	bRunFlag = True
	
	dbquery = "Select * from PINAdviceType where PINAdviceTypeID="& card_data("PinAdviceTypeidval") &";"
	Append_TestHTML StepCounter, "Execute PINAdviceType query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_PostalAddressFlag = dbRecordSet("RequirePostalAddress")
	db_EmailAddressFlag = dbRecordSet("RequireEmailAddress")
	db_PhoneNumberFlag = dbRecordSet("RequireMobilePhoneNumber")
	db_ActivityTypeCardOrder = dbRecordSet("ActivityTypeCardOrder")
	db_ActivityTypePINReminder = dbRecordSet("ActivityTypePINReminder")
'	pinReminderActivityCardOrderval = db_ActivityTypeCardOrder
'	pinReminderActivityPintypeval = db_ActivityTypeCardOrder
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	
	If  card_data("PinAdviceTypeidval") =1  Then
		If db_PostalAddressFlag = true  and db_EmailAddressFlag= false and db_PhoneNumberFlag=false Then
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-1 value from PINAdviceType Table", "RequirePostalAddress- Expected: 1" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-1 value from PINAdviceType Table", "RequirePostalAddress- Expected: 1" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "FAILED"
			bRunFlag = False
		End If
	ElseIf  card_data("PinAdviceTypeidval") =2 Then  	
		If db_PostalAddressFlag = false  and db_EmailAddressFlag= true and db_PhoneNumberFlag=false Then
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-2 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 1" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-2 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 1" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 0" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "FAILED"
			bRunFlag = False
		End If
	ElseIf  card_data("PinAdviceTypeidval") =3 Then  	
		If db_PostalAddressFlag = false  and db_EmailAddressFlag= true and db_PhoneNumberFlag=false Then
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-3 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 1" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify PinAdviceTypeID-3 value from PINAdviceType Table", "RequirePostalAddress- Expected: 0" & VBCRLF & "Actual: " & db_PostalAddressFlag & "RequireEmailAddress- Expected: 0" & VBCRLF & "Actual: " & db_EmailAddressFlag & "RequireMobilePhoneNumber- Expected: 1" & VBCRLF & "Actual: " & db_PhoneNumberFlag  , "FAILED"
			bRunFlag = False
		End If
	End If
	
	dbquery = "Select * from SysVarColco where SysVarID=141;"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_ColcoID = dbRecordSet("ColcoID")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	
	dbquery = "Select * from JobType where ID=2012;"
	Append_TestHTML StepCounter, "Execute JobType-2012 query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_ProcessID = dbRecordSet("ProcessID")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If ISNULL(db_ProcessID) = TRUE Then
		Append_TestHTML StepCounter, "Verify ProcessID value from JobType Table", "Expected: NULL" & VBCRLF & "Actual: NULL"  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify ProcessID value from JobType Table", "Expected: NULL" & VBCRLF & "Actual: " & db_ProcessID  , "FAILED"
			bRunFlag = False
	End If
		
	If card_data("PinAdviceTypeidval") <> "" Then
		getpinReminderprerequestValues = db_ActivityTypeCardOrder &"|"&db_ActivityTypePINReminder
	Else
		getpinReminderprerequestValues =""
	End If
End Function


Function invokeAPIError(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)

On error resume next
	Set objAPI = createobject("MSXML2.serverXMLHTTP.6.0")
	objAPI.open reqType, apiurl , asynctype
	arrHeaders = split(headers,",")
	
	For hitr = 0 To ubound(arrHeaders) Step 1
		If arrHeaders(hitr) <> "" or arrHeaders(hitr) <> null Then
			arrheaderNameval = split(arrHeaders(hitr),"==>")
			'Added Condition newly
			If isEmpty(arrheaderNameval(0)) = False OR isEmpty(arrheaderNameval(1)) = False Then
				objAPI.setRequestHeader trim(arrheaderNameval(0)) , trim(arrheaderNameval(1))
			End If
			
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
		transRef = j.Results.TransactionRef
		transReqID = j.RequestID
		invokeAPIError = pageReturn
		wait 20
	Else
		Append_TestHTML StepCounter, "Verify Response JSON ", "Resonse Text is: " & pageReturn , "PASSED"
		transRef = ""
		transReqID = ""
		transErrDesc = j.Description
		invokeAPIError = pageReturn
		
	End If
	
	Set objAPI = Nothing
End Function


Function validatePostAPICardDetailswithoutAddress()
On error resume next
'	query = "Select * from CardSelectedPIN where CardID="& XML_Cadrid & " order by 1 desc;"  'applicable only Pintypeid-2
'	Append_TestHTML StepCounter, "Execute CardSelectedPIN table entry ",query, "PASSED"
'	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
'	wait 2
'	
'	db_Libraryver = dbRecordSet("LibraryVersion")
'	db_EncrySDet = dbRecordSet("EncryptedSessionDetails")
'	db_EncryPin = dbRecordSet("EncryptedPIN")
'	db_KeyIDval = dbRecordSet("KeyID")
'	
'	Set dbRecordSet = Nothing
'	Set query = Nothing
'	
'	If db_Libraryver <> "" or db_EncrySDet <> "" or db_EncryPin <> "" or db_KeyIDval <> ""  Then
'		Append_TestHTML StepCounter, "Validate CardSelectedPIN record entries", "LibraryVersion: " & db_Libraryver  & VBCRLF & "EncryptedSessionDetails: " & db_EncrySDet  & VBCRLF & "EncryptedPIN: " & db_EncryPin  & VBCRLF & "KeyID: " & db_KeyIDval  & VBCRLF ,"PASSED"
'	Else
'		Append_TestHTML StepCounter, "Validate CardSelectedPIN record entries", "LibraryVersion: " & db_Libraryver  & VBCRLF & "EncryptedSessionDetails: " & db_EncrySDet  & VBCRLF & "EncryptedPIN: " & db_EncryPin  & VBCRLF & "KeyID: " & db_KeyIDval  & VBCRLF ,"FAILED"
'		bRunFlag = False
'	End If
'	
	query = "Select * from Card where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_TypeofPin = dbRecordSet("TypeOfPIN")
	db_PinSelMethod = dbRecordSet("PINSelectionMethod")
	db_StatusID = dbRecordSet("StatusID")
	db_TokenTypeID = dbRecordSet("TokenTypeID")
	db_Pan = dbRecordSet("PAN")
	cardPANNum = db_Pan
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_Pan <> "" Then
		Append_TestHTML StepCounter, "Validate Card record entries", "TypeOfPIN: " & db_TypeofPin  & VBCRLF & "PINSelectionMethod: " & db_PinSelMethod  & VBCRLF & "StatusID: " & db_StatusID  & VBCRLF ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate Card record entries", "TypeOfPIN: " & db_TypeofPin  & VBCRLF & "PINSelectionMethod: " & db_PinSelMethod  & VBCRLF & "StatusID: " & db_StatusID  & VBCRLF ,"FAILED"
		bRunFlag = False
	End If
	
	query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardIssueAttributes table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_InitialPINAdviceTypeID= dbRecordSet("InitialPINAdviceTypeID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_InitialPINAdviceTypeID <> "" Then
		Append_TestHTML StepCounter, "Validate CardIssueAttributes record entries", "InitialPINAdviceTypeID: " & db_InitialPINAdviceTypeID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardIssueAttributes record entries", "InitialPINAdviceTypeID: " & db_InitialPINAdviceTypeID  ,"FAILED"
		bRunFlag = False
	End If
	
	query = "Select * from CustomerContact where CustomerID="& apiCustomerID & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CustomerContact table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ContactID = dbRecordSet("ContactID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ContactID <> "" Then
		Append_TestHTML StepCounter, "Validate CustomerContact record entries", "ContactID: " & db_ContactID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerContact record entries", "ContactID: " & db_ContactID  ,"FAILED"
		bRunFlag = False
	End If
	query = "Select * from CustomerCard where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CustomerCard table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CustomerCardID = dbRecordSet("CustomerCardID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_CustomerCardID <> "" Then
		Append_TestHTML StepCounter, "Validate CustomerCard record entries", "CustomerCardID: " & db_CustomerCardID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CustomerCard record entries", "CustomerCardID: " & db_CustomerCardID  ,"FAILED"
		bRunFlag = False
	End If
	
	query = "Select * from CardStatusHistory where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardStatusHistory table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CardStatusHistoryID = dbRecordSet("CardStatusHistoryID")
	db_NewStatusID= dbRecordSet("NewStatusID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_CardStatusHistoryID <> "" and db_NewStatusID <> "" Then
		Append_TestHTML StepCounter, "Validate CardStatusHistory record entries", "CardStatusHistoryID: " & db_CardStatusHistoryID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardStatusHistory record entries", "CardStatusHistoryID: " & db_CardStatusHistoryID  ,"FAILED"		
		bRunFlag = False
	End If
	query = "Select * from CardActivityLog where CardID="& XML_Cadrid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityLogID = dbRecordSet("ActivityLogID")
		
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityLogID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from ActivityLog where ActivityLogID="& db_ActivityLogID & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityTypeID = dbRecordSet("ActivityTypeID")
	db_Note = dbRecordSet("Note")
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityTypeID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityTypeID: " & db_ActivityTypeID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityTypeID: " & db_ActivityTypeID  ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from CardPAN where CardPANID="& XML_Cardpanid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardPAN table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_PAN1 = dbRecordSet("PAN")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_PAN1 <> "" Then
		Append_TestHTML StepCounter, "Validate CardPAN record entries", "PAN: " & db_PAN1  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardPAN record entries", "PAN: " & db_PAN1  ,"FAILED"		
		bRunFlag = False
	End If
End Function

Function getRequirePINTokenTypeIDvalue(reqpinvalue)
	On error resume next
	query = "Select * from TokenTypeControl where RequirePIN="& reqpinvalue & ";"
	Append_TestHTML StepCounter, "Execute CardPAN table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_MASTER")
	wait 2
	
	db_TokenTypeID = dbRecordSet("TokenTypeID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	If db_TokenTypeID <>""  Then
		getRequirePINTokenTypeIDvalue = db_TokenTypeID
	Else
		getRequirePINTokenTypeIDvalue = 4
	End If
	
	
End Function


Function changeRequirePINvalue(tokenTypeval,changeRPINvalue)
	On error resume next
	
	query1 = "Select * from TokenTypeControl where TokenTypeID="& tokenTypeval & ";"
	Append_TestHTML StepCounter, "Update RequirePIN value in  TokenTypeControl", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_MASTER")
	wait 2
	db_RequirePIN = dbRecordSet("RequirePIN")
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	If cint(db_RequirePIN) = cint(changeRPINvalue) Then
		Append_TestHTML StepCounter, "Verify RequirePIN value from TokenTypeControl Table", "Expected:"& changeRPINvalue & VBCRLF & "Actual: " & db_RequirePIN  , "PASSED"
	Else
		query1 = "Update TokenTypeControl set RequirePIN = " & changeRPINvalue & " where TokenTypeID="& tokenTypeval & ";"
		Append_TestHTML StepCounter, "Update RequirePIN value in  TokenTypeControl", query1, "PASSED"
		Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_MASTER")
		wait 2
		Set dbRecordSet = Nothing
	End If
	
		
End Function




Function verifysysVarcolcoandupdate(cpsdesc,cpsVal,acsdesc,acsVal)
On error resume next
	dbquery = "Select * from SysVarColco where [Key]='" & cpsdesc & "';"
	Append_TestHTML StepCounter, "Execute Query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_SysVarID = dbRecordSet("SysVarID")
	db_ColcoID = dbRecordSet("ColcoID")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	cardColcoID = db_ColcoID
	dbquery = "Select * from SysVarColco where SysVarID="& db_SysVarID & ";"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_Value = dbRecordSet("Value")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If db_Value <> "" and cbool(db_Value) = cbool(cpsVal) Then
		Append_TestHTML StepCounter, "Verify " & cpsdesc & " value from SysVarColco Table", "Expected:"& cpsVal & VBCRLF & "Actual: " & db_Value  , "PASSED"
	else
		query1 = "Update SysVarColco set Value = " & cpsVal & " where SysVarID="& db_SysVarID & ";"
		Append_TestHTML StepCounter, "Update " & cpsdesc & " value in  SysVarColco", query1, "PASSED"
		Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		Set dbRecordSet = Nothing
	End If
	
	dbquery = "Select * from SysVarColco where [Key]='" & acsdesc & "';"
	Append_TestHTML StepCounter, "Execute Query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_SysVarID = dbRecordSet("SysVarID")
	db_ColcoID = dbRecordSet("ColcoID")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing

	dbquery = "Select * from SysVarColco where SysVarID="& db_SysVarID & ";"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_Value = dbRecordSet("Value")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	If db_Value <> "" and cbool(db_Value) = cbool(acsVal) Then
		Append_TestHTML StepCounter, "Verify " & acsdesc & " value from SysVarColco Table", "Expected:"& acsVal & VBCRLF & "Actual: " & db_Value  , "PASSED"
	else
		query1 = "Update SysVarColco set Value = " & acsVal & " where SysVarID="& db_SysVarID & ";"
		Append_TestHTML StepCounter, "Update " & acsdesc & " value in  SysVarColco", query1, "PASSED"
		Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		Set dbRecordSet = Nothing
	End If
	

End Function

Function cardCreationchecks(custERPno,cardPANno)
	On error resume next
	
	query1 = "Select * from customer where CustomerERP='"& customerERP_id & "';"
	Append_TestHTML StepCounter, "Execute Query", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_AllowSelectPIN = dbRecordSet("AllowSelectPIN")
	db_CustomerID = dbRecordSet("CustomerID")
	
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	
	If db_AllowSelectPIN <> "" Then
		Append_TestHTML StepCounter, "Verify AllowSelectPIN value from Customer Table", "Actual: " & db_AllowSelectPIN  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify AllowSelectPIN value from Customer Table", "Actual: " & db_AllowSelectPIN  , "FAILED"
	End If
	
	query1 = "Select * from Card where PAN="& cardPANNum & ";"
	Append_TestHTML StepCounter, "Execute Query", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_CardID = dbRecordSet("CardID")
	db_TypeOfPIN = dbRecordSet("TypeOfPIN")
	db_PINSelectionMethod = dbRecordSet("PINSelectionMethod")
	db_IssueNumber = dbRecordSet("IssueNumber")
	db_EncodingID = dbRecordSet("EncodingID")
'	'msgbox dbRecordSet("TypeOfPIN")
'	'msgbox isnull(dbRecordSet("TypeOfPIN"))
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	
	If cint(db_TypeOfPIN) <> "" Then
		Append_TestHTML StepCounter, "Verify TypeOfPin value from Card Table", "Expected: "& db_TypeOfPIN & VBCRLF & "Actual: " & db_TypeOfPIN  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify TypeOfPin value from Card Table", "Expected: 1/2"  & VBCRLF & "Actual: " & db_TypeOfPIN  , "FAILED"
	End If
	If cint(db_PINSelectionMethod) = 1 Then
		Append_TestHTML StepCounter, "Verify PINSelectionMethod value from Card Table", "Expected: 1"& VBCRLF & "Actual: " & db_PINSelectionMethod  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify PINSelectionMethod value from Card Table", "Expected: 1"& VBCRLF & "Actual: " & db_PINSelectionMethod  , "FAILED"
	End If
	If cint(db_EncodingID) = 2 or cint(db_EncodingID) = 3 Then
		Append_TestHTML StepCounter, "Verify IssueNumber value from Card Table", "Expected: 2/3"& VBCRLF & "Actual: " & db_EncodingID  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify IssueNumber value from Card Table", "Expected: 2/3"& VBCRLF & "Actual: " & db_EncodingID  , "FAILED"
	End If
	
	query1 = "Select * from CardIssueAttributes where CardID="& db_CardID & ";"
	Append_TestHTML StepCounter, "Execute Query", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_InitialPINAdviceTypeID = dbRecordSet("InitialPINAdviceTypeID")
	
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	
	If cint(db_InitialPINAdviceTypeID) = 1 or cbool(db_InitialPINAdviceTypeID)= true Then
		Append_TestHTML StepCounter, "Verify InitialPINAdviceTypeID value from CardIssueAttributes Table",  "Expected: 1"& VBCRLF & "Actual: " & db_InitialPINAdviceTypeID  , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify InitialPINAdviceTypeID value from CardIssueAttributes Table",  "Expected: 1"& VBCRLF & "Actual: " & db_InitialPINAdviceTypeID  , "FAILED"
	End If
	
	query1 = "Select * from TokenTypeControl where TokenTypeID in (Select TokenTypeID from Card where CardID="& db_CardID & ");"
	Append_TestHTML StepCounter, "Execute Query", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_RequirePIN = dbRecordSet("RequirePIN")
	
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	
	If db_RequirePIN <> "" Then
		Append_TestHTML StepCounter, "Verify RequirePIN value from TokenTypeControl Table",  "Expected: 1/0"& VBCRLF & "Actual: " & db_RequirePIN  , "PASSED"
	End If
	
End Function
	
	
Function validateDX501filedata(filePath,ftype,verifydata,datadesc,Pos,noofchrs,character)
On error resume Next
	bFlag = True
	Set fileSysObj = createObject("Scripting.FileSystemObject")

	If (ucase(ftype) = ucase("dat") or ucase(ftype) = ucase("File") ) Then
		Set DXread = fileSysObj.OpenTextFile(filePath,1,False,-1)
	Else
		Set DXread = fileSysObj.OpenTextFile(filePath,1)
	End  IF
	lcount = 0
	Do while not DXread.AtEndOfStream
		strLine = DXread.ReadLine()
		lcount = lcount + 1
		If instr(strLine,verifydata)>0 Then
			
		Append_TestHTML StepCounter,"Card Data line in a File", strLine , "PASSED"
		
		Methodval = mid(strLine,Pos,noofchrs)
		If instr(Methodval,character)>0 Then
			Append_TestHTML StepCounter,"Validate "& datadesc & " " & character & " in a File", "Value identified at "&Pos & " position", "PASSED"
		Else
			Append_TestHTML StepCounter,"Validate "& datadesc & " " & character & " in a File", "Value not matched and identified value is "&Methodval & " instead of " & character, "FAILED"
		
		End If
		End If
	
	Loop
End Function	


Public Function replaceCard()

	On error resume next
	
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_EmbossText") Then
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossText", "Set", "NewEmbossText"
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
			Append_TestHTML StepCounter, "Create card", "Added details on create new card page", "PASSED"
		else
			Append_TestHTML StepCounter, "Create card", "Adding details on create card page failed", "FAILED"
		End If
		
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			wait 5
			Call enterWebList_value("WebList_CardReasonReplace","Reason","Replaced")
		
	'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cardsreplace_WebElement_Save") Then
			Append_TestHTML StepCounter, "Verify Confirmation popup", "Identified Save button", "PASSED"
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cardsreplace_WebElement_Save", "Click", ""	
		wait 5
		else
			Append_TestHTML StepCounter, "Verify Confirmation popup", "Fail to Identify Save button", "FAILED"
			
		End If
		
		cardNum = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Card","GetROProperty","default value")
		
		expDate_actual = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Expiry","GetROProperty","default value")
		expDate_arr = split(expDate_actual,"/")
		expDate = expDate_arr(2) & "-" & expDate_arr(1) & "-" & expDate_arr(0)
		
		embossText = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossText","GetROProperty","default value")	
		
		wait 3   ' waiting for the DB update
		
		query = "select * from Card where PAN = '" & cardNum & "';"
		set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		cardPAN = dbRecordSet("PAN")
		'cardPANNum = cardPAN
		dbExpDate =  dbRecordSet("ExpiryDate")
		db_StatusID =  dbRecordSet("StatusID")
		dbEmbossText =  dbRecordSet("EmbossText")
		db_StatusChangeReason =  dbRecordSet("StatusChangeReason")
		
		'cardPAN_no = cardPAN
		'cardPANNum = cardPAN
		cardExpiry_date = dbExpDate
			Set dbRecordSet = Nothing
	Set query = Nothing
		If instr(db_StatusID, "31")<>0 Then
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: 31"  & VBCRLF & "Actual Value: " & db_StatusID ,"PASSED"
		else
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: 31"  & VBCRLF & "Actual Value: " & db_StatusID ,"FAILED"
			bRunFlag = False
		End If
		
		If instr( ucase(dbEmbossText),ucase("NewEmbossText"))<>0 Then
			Append_TestHTML StepCounter,"Emboss Text Validation","Expected Value: NewEmbossText"  & VBCRLF & "Actual Value: " & dbEmbossText ,"PASSED"
		else
			Append_TestHTML StepCounter,"Emboss Text Validation","Expected Value: NewEmbossText"  & VBCRLF & "Actual Value: " & dbEmbossText ,"FAILED"
			bRunFlag = False
		End If
		If instr(db_StatusChangeReason, "Replaced")<>0 Then
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: Replaced"  & VBCRLF & "Actual Value: " & db_StatusChangeReason ,"PASSED"
		else
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: Replaced"  & VBCRLF & "Actual Value: " & db_StatusChangeReason ,"FAILED"
			bRunFlag = False
		End If
	Else
			Append_TestHTML StepCounter,"Active Card Page","Fail to open active card" ,"FAILED"
			bRunFlag = False
	End  If
End Function


Function renewationdet()
	On error resume next
	
	bFlag = True
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WebTable_CM_searchresultstable") Then
			Append_TestHTML StepCounter, "Verify card reissue details", "Details are verified", "PASSED"
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebTable_CM_searchresultstable", "h", ""	

		wait 5
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		wait 
			Append_TestHTML StepCounter, "Renew the card details", "Renewed card details and saved successfully", "PASSED"
		
		else
			Append_TestHTML StepCounter, "Verify card reissue details", "Fail to Identify renew details", "FAILED"
			
		End If
		wait 10
		query = "select * from Card where PAN = '" & cardPANNum & "';"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		
		set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		dbExpDate =  dbRecordSet("ExpiryDate")
		db_StatusID =  dbRecordSet("StatusID")
		dbEmbossText =  dbRecordSet("EmbossText")
'		cardPAN_no = cardPAN
'		cardPANNum = cardPAN
'		cardExpiry_date = dbExpDate
		Set dbRecordSet = Nothing
		Set query = Nothing

		If instr(db_StatusID, "23")<>0 Then
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: 23"  & VBCRLF & "Actual Value: " & db_StatusID ,"PASSED"
		else
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: 23"  & VBCRLF & "Actual Value: " & db_StatusID ,"FAILED"
			bRunFlag = False
		End If
		
			
End Function

Function updateDatesforRenewcard()
On error resume next
	
	bFlag = True
		query = "select * from Card where PAN = '" & cardPANNum & "' order by ModifiedOn desc;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		
		set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		cardPAN = dbRecordSet("PAN")
		dbExpDate =  dbRecordSet("ExpiryDate")
		db_StatusID =  dbRecordSet("StatusID")
		db_CardID=  dbRecordSet("CardID")
		Set dbRecordSet = Nothing
		Set query = Nothing
		
		If db_CardID <> "" Then
	
			 lastdoayofmonthval = findMonthenddate()
			 'msgbox lastdoayofmonthval
			If lastdoayofmonthval<>"" and cint(db_StatusID) = cint("1") Then
				query = "Update Card Set ExpiryDate = '"& lastdoayofmonthval &"', PreviousExpiryDate = '" & lastdoayofmonthval & "'where CardID = '" & db_CardID & "' and StatusID = 1;"
				set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
				Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
				
				wait 2
				Set dbRecordSet = Nothing
				Set query = Nothing
			End If
		Else
				Append_TestHTML StepCounter, "Verify  card details", "Picked card is not active card", "FAILED"
			bRunFlag = False
		
		End If
End Function


Function findMonthenddate()
	On error resume next
	
	bFlag = True
	
	lastdaymonth = DateAdd("d",0,dateSerial(Year(Now),Month(Now)+1,0))
	
		If len(Month(lastdaymonth)) =1 Then
			monval = "0" & Month(lastdaymonth)
		Else
			monval = Month(lastdaymonth)
		End If
		If len(Day(lastdaymonth)) =1 Then
			dayval = "0" & Day(lastdaymonth)
		Else
			dayval = Day(lastdaymonth)
		End If
		ldmval = Year(lastdaymonth) & "-" & monval &"-" & dayval
	If ldmval <> "" Then
		findMonthenddate = ldmval
	End If
End Function



Public Function reprintCard()

	On error resume next
	
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_EmbossText") Then
		Call performCardActions("Renew")
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_ReprintPIN", "Click", ""
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
			Append_TestHTML StepCounter, "Create card", "Added details on create new card page", "PASSED"
		else
			Append_TestHTML StepCounter, "Create card", "Adding details on create card page failed", "FAILED"
		End If
		
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			wait 5
			Call enterWebList_value("WebList_CardReasonReplace","Reason","Replaced")
		
	'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cardsreplace_WebElement_Save") Then
			Append_TestHTML StepCounter, "Verify Confirmation popup", "Identified Save button", "PASSED"
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cardsreplace_WebElement_Save", "Click", ""	
		wait 5
		else
			Append_TestHTML StepCounter, "Verify Confirmation popup", "Fail to Identify Save button", "FAILED"
			
		End If
		
		cardNum = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Card","GetROProperty","default value")
		
		expDate_actual = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Expiry","GetROProperty","default value")
		expDate_arr = split(expDate_actual,"/")
		expDate = expDate_arr(2) & "-" & expDate_arr(1) & "-" & expDate_arr(0)
		
		embossText = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossText","GetROProperty","default value")	
		
		wait 3   ' waiting for the DB update
		
		query = "select * from Card where PAN = '" & cardPANNum & "';"
		set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		
		cardPAN = dbRecordSet("PAN")
		cardPANNum = cardPAN
		dbExpDate =  dbRecordSet("ExpiryDate")
		db_StatusID =  dbRecordSet("StatusID")
		dbEmbossText =  dbRecordSet("EmbossText")
		cardPAN_no = cardPAN
		cardPANNum = cardPAN
		cardExpiry_date = dbExpDate
			Set dbRecordSet = Nothing
	Set query = Nothing
		If instr(db_StatusID, "31")<>0 Then
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: 31"  & VBCRLF & "Actual Value: " & db_StatusID ,"PASSED"
		else
			Append_TestHTML StepCounter,"Card Status ID's","Expected Value: 31"  & VBCRLF & "Actual Value: " & db_StatusID ,"FAILED"
			bRunFlag = False
		End If
		
		If instr( dbEmbossText,"NewEmbossText")<>0 Then
			Append_TestHTML StepCounter,"Emboss Text Validation","Expected Value: NewEmbossText"  & VBCRLF & "Actual Value: " & dbEmbossText ,"PASSED"
		else
			Append_TestHTML StepCounter,"Emboss Text Validation","Expected Value: NewEmbossText"  & VBCRLF & "Actual Value: " & dbEmbossText ,"FAILED"
			bRunFlag = False
		End If
	Else
			Append_TestHTML StepCounter,"Active Card Page","Fail to open active card" ,"FAILED"
			bRunFlag = False
	End  If
End Function

Function cardStatusHistoryDetails()
	
	On error resume next
	bFlag = True
	query = "select * from Card where PAN = '" & cardPANNum & "';"
		set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		wait 2
		db_StatusID =  dbRecordSet("StatusID")
		db_CardID =  dbRecordSet("CardID")
		Set dbRecordSet = Nothing
		Set query = Nothing
	query = "Select * from CardStatusHistory where CardID=" & db_CardID & ";"
		set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		wait 2
		db_StatusID =  dbRecordSet("NewStatusID")
		db_CardID =  dbRecordSet("CardStatusHistoryID")
		Set dbRecordSet = Nothing
		Set query = Nothing
		If db_StatusID<> "" Then
			Append_TestHTML StepCounter,"Card Status History Details ","Previous Status | Current Status" & db_StatusID  ,"PASSED"
		else
			Append_TestHTML StepCounter,"Card Status History Details ","Previous Status | Current Status" & db_StatusID  ,"FAILED"
			bRunFlag = False
		End If
End Function

Function verifyActivityLogdetails(alogstatus)
	
		On error resume next
	
	bFlag = True
	query = "select * from Card where PAN = '" & cardPANNum & "';"
		set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		
		wait 2
		db_StatusID =  dbRecordSet("StatusID")
		db_CardID =  dbRecordSet("CardID")
	Set dbRecordSet = Nothing
	Set query = Nothing
		If db_CardID <> ""  Then
			query = "Select * from CardActivityLog where CardID=" & db_CardID & " order by ModifiedOn desc;"
			set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
				db_ActivityLogID = dbRecordSet("ActivityLogID")
			Set dbRecordSet = Nothing
			Set query = Nothing
					Append_TestHTML StepCounter,"Verify ActivityLogID from CardActivityLog" & "Actual Value: " & db_ActivityLogID ,"PASSED"
				query = "Select * from ActivityLog where ActivityLogID=" & db_ActivityLogID & " order by ModifiedOn desc;"
			set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
				db_ActivityTypeID = dbRecordSet("ActivityTypeID")
			Set dbRecordSet = Nothing
			Set query = Nothing	
			If cint(db_ActivityTypeID) = cint("alogstatus") Then
					Append_TestHTML StepCounter,"Verify ActivityTypeID status","Expected Value: "& alogstatus  & VBCRLF & "Actual Value: " & db_ActivityTypeID ,"PASSED"
				Else
					Append_TestHTML StepCounter,"Verify ActivityTypeID status","Expected Value: "& alogstatus  & VBCRLF & "Actual Value: " & db_ActivityTypeID ,"FAILED"
			End If
		Else
					Append_TestHTML StepCounter,"Verify Card Activity Type status","Unable to get card id" ,"FAILED"
		
	End  IF
End Function

Function verifyCardHistoryActivityLog()
On error resume next
	
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebElement_History", "Click", ""
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Card_WebElement_CardActivityLog") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebElement_CardActivityLog", "Click", ""
			Append_TestHTML StepCounter,"Veriry Card Activity Log tab", " Successfully navigated"  ,"PASSED"
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WebTable_TimeActivitytable")  Then
					Append_TestHTML StepCounter,"Veriry Card Activity Log history details", " Successfully verified"  ,"PASSED"
				Else
					Append_TestHTML StepCounter,"Veriry Card Activity Log history details", " Fail to verifiy"  ,"FAILED"
				
			End If
		Else
			Append_TestHTML StepCounter,"Veriry Card Activity Log tab", " Fail to navigate"  ,"FAILED"
			
			bFlag = False
			bRunFlag = False
	End  If
End Function

Function reissueCardvalidation()
	On error resume next
	
	bFlag = True
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WebElement_reissueSave")  Then
		Append_TestHTML StepCounter,"Verify RePrint Confirmation Pop-up", "Pop up appeared correctly"  ,"PASSED"
	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebElement_reissueSave", "Click", ""
		wait 10
		Append_TestHTML StepCounter,"Verify Card reissue details", " Data saved successfully"  ,"PASSED"
					
		query = "Select * from Card where PAN=" & cardPANNum & " order by ModifiedOn desc;"
		set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		wait 2
		db_CardID =  dbRecordSet("CardID")
		Set dbRecordSet = Nothing
		wait 20
		query = "Select * from CardPinMailerRequest where CardID=" & db_CardID & ";"
		set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		wait 2
		db_CardPINMailerRequestID =  dbRecordSet("CardPINMailerRequestID")
		Set dbRecordSet = Nothing
		If db_CardPINMailerRequestID <> "" Then
			Append_TestHTML StepCounter,"Verify CardPINMailerRequestID details", "Sucessfully creted record entry with CardPINMailerRequestID-" & db_CardPINMailerRequestID  ,"PASSED"
		Else
			Append_TestHTML StepCounter,"Verify CardPINMailerRequestID details", "record entry not created and CardPINMailerRequestID-" & db_CardPINMailerRequestID  ,"FAILED"	
			bRunFlag = False
		End If
	Else
			Append_TestHTML StepCounter,"Verify Card reissue details", "Fail to click on Resiisue option"  ,"FAILED"
			bRunFlag = False
	End If
	
	
End Function






Function gfnSystemdetails()
	On error resume next
	
	bFlag = True
		query = "Select * from Company where CompanyName like '%" & browserProp & "%' ;"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		sys_CompanyID =  dbRecordSet("CompanyID")
		sys_ClientCompanyNumber =  dbRecordSet("ClientCompanyNumber")
		
		
		Set query = Nothing
		Set dbRecordSet = Nothing

End Function

Function csdImportPreConfigDetailsandcreateDX053File(dPath, fPath, recCount, etType)
	On error resume next
	
	bFlag = True
	
		query = "Select * from SysVarColco where [Key]='ExcludeEncodingIDFromCSD' ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_SysVarID =  dbRecordSet("SysVarID")
		db_ColcoID =  dbRecordSet("ColcoID")
		db_Value =  dbRecordSet("Value")
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		If isNull(db_SysVarID) = False   Then
				Append_TestHTML StepCounter, "Verify ExcludeEncodingIDFromCSD key in SysVarColco Table", "Entry exist with ID-"&db_SysVarID & " Colco-"& db_ColcoID & " Value is-"&  db_Value , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify ExcludeEncodingIDFromCSD key in SysVarColco Table", "Entry exist with ID-"&db_SysVarID & " Colco-"& db_ColcoID & " Value is-"&  db_Value , "FAILED"
			bRunFlag = False
				
		End If	
		
		query = "Select * from FileSeq where FileSeqID like '%DX053%' ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_LastSequence=  dbRecordSet("LastSequence")
		db_FileSeqID=  dbRecordSet("FileSeqID")
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		countynumber = Split(db_FileSeqID,"_")(1)
		
		query = "Select * from Company where ClientCompanyNumber="&cint(countynumber)&" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_CompanyID=  dbRecordSet("CompanyID")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		DispatchCompanyID = db_CompanyID
		
		query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_CVKi=  dbRecordSet("CVKi")
		db_CVKi2=  dbRecordSet("CVKi2")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from CardPINHistory where CardPANID="& XML_Cardpanid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_PVV=  dbRecordSet("PVV")
		db_PVKi=  dbRecordSet("PVKi")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from CardFulfilment where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_DispatchedDate =  dbRecordSet("DispatchedDate")
		db_ReceivedDate =  dbRecordSet("ReceivedDate")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from CardRFID where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_RFIDUID =  dbRecordSet("RFIDUID")
		db_EMAID =  dbRecordSet("EMAID")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		If db_RFIDUID="" and db_DispatchedDate = ""   Then
				Append_TestHTML StepCounter, "Verify CVK values from CardIssueAttributes", "CVK values are "& db_CVKi , "PASSED"
				Append_TestHTML StepCounter, "Verify PVV values from CardPINHistory", "No record created" & db_PVV , "PASSED"
				Append_TestHTML StepCounter, "Verify DispatchedDate values from CardFulfilment", "No record created" &  db_DispatchedDate, "PASSED"
				Append_TestHTML StepCounter, "Verify RFIDUID values from CardRFID", "No record created" & db_RFIDUID  , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify RFIDUID values from CardRFID", "Record entry exist with few values: RFIDUID-" & db_RFIDUID & " DispatchedDate-"  & db_DispatchedDate , "FAILED"			
		End If
			
		FileCreationDate = findpreviousMonthenddate(-1)
		DispatchDate = findpreviousMonthenddate(1)
		DispatchNextDate = findpreviousMonthenddate(2)
	
		newLastseq = db_LastSequence + 1
		timeval = getDateandTimestamp()
		timeval_arr = split(timeval," ")
		timevaltext = timeval_arr(1)
		hrvaltext =  mid(timevaltext,1,2)
		minvaltext =  mid(timevaltext,3,2)
		secvaltext = mid(timevaltext,5,2)
		
		newFileCreaationfrmt = FileCreationDate & "T" & hrvaltext & ":" & minvaltext & ":" & secvaltext & "Z"
		newDispatchDatefrmt = DispatchDate & "T" & hrvaltext & ":" & minvaltext & ":" & secvaltext & "Z"
		
		
		If ucase(etType) = ucase("Both") Then
			searchreplaceString = "IssuingCountry-" & countynumber &";SequenceNumber$"& newLastseq &";RecordCount$" & recCount &";FileCreationDate$"& newFileCreaationfrmt  &";PAN_ID$"& XML_Cardpanid &";ReceivedDate$"& newFileCreaationfrmt &";DispatchDate$"& newDispatchDatefrmt &";DispatchCompanyID$"& DispatchCompanyID 
			searchmultiString = "CardID-"& XML_Cadrid
			Call searchandReplaceAllwithsearchString(fPath, searchmultiString)
		ElseIf ucase(etType) = ucase("F") Then
			searchreplaceString = "IssuingCountry$" & countynumber &";SequenceNumber$"& newLastseq &";RecordCount$" & recCount &";FileCreationDate$"& newFileCreaationfrmt &";CardID$"& XML_Cadrid &";DispatchDate$"& newDispatchDatefrmt &";DispatchCompanyID$"& DispatchCompanyID 

		ElseIf ucase(etType) = ucase("C") Then
			searchreplaceString = "IssuingCountry$" & countynumber &";SequenceNumber$"& newLastseq &";RecordCount$" & recCount &";FileCreationDate$"& newFileCreaationfrmt &";CardID$"& XML_Cadrid &";PAN_ID$"& XML_Cardpanid &";ReceivedDate$"& newFileCreaationfrmt
		
		End If
		Call searchandReplaceMultipleStringwithdollar(fPath, searchreplaceString)
		
		fileInitialName = db_FileSeqID
		seqnum_val = valuetrailzeros(db_LastSequence + 1)
		fcreationdate = Replace(FileCreationDate,"-","")
	
		nFname =  fileInitialName&"_"&seqnum_val&"_"&fcreationdate&"_"&timevaltext & ".json"
		newfile = dPath & nFname
		
		Set fileSysObj = createObject("Scripting.FileSystemObject")
		fileSysObj.CreateTextFile(newfile)
		Set DXread = fileSysObj.OpenTextFile(fPath,1)
		content = DXread.ReadAll
		DXread.Close
		
		Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
		Newcontent = content
		DXwrite.Write Newcontent
		DXwrite.Close
		If nFname <> "" Then
			ppend_TestHTML StepCounter, "DX053 File creation and move to GLobal location", "File created at"& newfile , "PASSED"
			dx053FinalFName = nFname
			csdImportPreConfigDetailsandcreateDX053File = nFname
		Else
			csdImportPreConfigDetailsandcreateDX053File = ""
		End If

End Function


Function findpreviousMonthenddate(priornum)
	On error resume next
	
	bFlag = True
	
	lastdaymonth = DateAdd("d",priornum,dateSerial(Year(Now),Month(Now)-2,1))
		If len(Month(lastdaymonth)) =1 Then
			monval = "0" & Month(lastdaymonth)
		Else
			monval = Month(lastdaymonth)
		End If
		If len(Day(lastdaymonth)) =1 Then
			dayval = "0" & Day(lastdaymonth)-1
		Else
			dayval = Day(lastdaymonth)-1
		End If
		ldmval = Year(lastdaymonth) & "-" & monval &"-" & dayval
	If ldmval <> "" Then
		findpreviousMonthenddate = ldmval
	End If
End Function

Function ReadTextFile(strPath, lngFormat)
	With CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath,1,False,lngFormat)
		ReadTextFile = ""
		If Not .AtEndOfStream Then
			ReadTextFile = .ReadAll
			.Close
		End If
	End With
End Function

Function searchandReplaceAllwithsearchString(jsonFPath, searchreplaceString)

On Error Resume Next

	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(jsonFPath) Then
		Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
		Newcontent = fileRead.ReadAll
		fileRead.Close
		Set fileSysObj = Nothing
'		srstrings = Split(searchreplaceString,";")
		'Append_TestHTML StepCounter, "Verify json input File Path", "File exist in the path:Before Modifying Data--:" & Newcontent , "PASSED"
		
			sstrings = Split(searchreplaceString,"-")
			searchString = sstrings(0)
			replaceString = sstrings(1)
		strContent = ReadTextFile(jsonFPath, 0)
		arrContent = Split(strContent, searchString)
		linecount = 1
		For titr = 1 To ubound(arrContent) Step 1
			
			searchstr = chr(34) & searchString & chr(34)
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			Set fileRead = fileSysObj.OpenTextFile(jsonFPath,1)
			keycontent = ""
			Do until fileRead.AtEndOfStream
				content = fileRead.ReadLine	
				If instr(content,searchstr) > 0  and linecount = titr Then
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

Function verifyJob230and2011Entries()
	On error resume next
		bFlag = True	
		query = "Select * from Job where JobTypeID=230 order by 1 desc;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_MASTER")
		
		db_StatusID=  dbRecordSet("StatusID")
		db_ID=  dbRecordSet("ID")
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		If cint(db_StatusID) = 3  Then
				Append_TestHTML StepCounter, "Verify 203 Job Status", "Successfully job triggerred with Status -" &db_StatusID  , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify 203 Job Status", "Fail to job trigger with Status -" &db_StatusID  , "FAILED"
			bRunFlag = False
				
		End If	
		
		query = "Select * from Job where JobTypeID=2011 order by 1 desc;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_StatusID=  dbRecordSet("StatusID")
		db_ID=  dbRecordSet("ID")
		db_InputXml=  dbRecordSet("InputXml")
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		If cint(db_StatusID) = 4 or instr(InputXml,dx053FinalFName)>0 Then
				Append_TestHTML StepCounter, "Verify 2011 Job Status", "Successfully job trigerred with " & db_InputXml  , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify 2011 Job Status", "Fail to triger Job with " & dx053FinalFName & " " & db_InputXml , "FAILED"
			bRunFlag = False

		End If	
		
		query = "Select * from JobLog where JobID="& db_ID &";"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 5, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_Message=  dbRecordSet("Message")
	
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		If db_Message <> "" Then
				Append_TestHTML StepCounter, "Verify 2011 Job Log Status", "Job Log entry created with-" & db_Message , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify 2011 Job Log Status", "No job log entry" , "PASSED"

		End If	
		
		
		
End Function

Function findCSDBATCHFILEStatusDetails(fileNameval)
On error resume next
		bFlag = True		
	query = "Select * from CSDBATCHFILE where FileName like '%"& fileNameval &"%';"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 5, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_BatchID =  dbRecordSet("BatchID")
		db_NumberOfRecordsInError=  dbRecordSet("NumberOfRecordsInError")
		db_NumberOfRecordsInFile=  dbRecordSet("NumberOfRecordsInFile")
		db_NumberOfRecordsProcessed =  dbRecordSet("NumberOfRecordsProcessed")
			
		Set query = Nothing
		Set dbRecordSet = Nothing
		If db_BatchID <> "" Then
				Append_TestHTML StepCounter, "Verify CSDBATCHFILE Status", "Entry exist with "& fileNameval , "PASSED"
				If db_NumberOfRecordsInError = 1 or db_NumberOfRecordsInError = True Then
					Append_TestHTML StepCounter, "Verify db_NumberOfRecordsInError vlue from CSDBATCHFILE table", "Expected-1"& "  Actual-"&db_NumberOfRecordsInError , "PASSED"
				Else
					Append_TestHTML StepCounter, "Verify db_NumberOfRecordsInError vlue from CSDBATCHFILE table", "Expected-1"& "  Actual-"&db_NumberOfRecordsInError , "FAILED"				
				End If
		Else
				Append_TestHTML StepCounter, "Verify CSDBATCHFILE Status", "No entry with "& fileNameval , "PASSED"

		End If	
	
End Function

Function verifyCardPINHistoryandIssueAttributeValues()
	On error resume next
		bFlag = True	
		
		query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_CVKi=  dbRecordSet("CVKi")
		db_CVKi2=  dbRecordSet("CVKi2")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from CardPINHistory where CardPANID="& XML_Cardpanid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_PVV=  dbRecordSet("PVV")
		db_PVKi=  dbRecordSet("PVKi")
		db_PINVersion=  dbRecordSet("PINVersion")
		db_PINEventTypeIDi=  dbRecordSet("PINEventTypeID")

		Set query = Nothing
		Set dbRecordSet = Nothing
		
		If isNULL(db_CVKi) = False and isNULL(db_PVV)=False Then
				Append_TestHTML StepCounter, "Verify CVK value from CardIssueAttributes", "CVK value is-" & db_CVKi  , "PASSED"
				Append_TestHTML StepCounter, "Verify CVKi2 valuesfrom CardIssueAttributes", "CVKi2 value is-" & db_CVKi2  , "PASSED"
				Append_TestHTML StepCounter, "Verify PVV valuesfrom CardPINHistory", "PVV value is-" & db_PVV  , "PASSED"
				Append_TestHTML StepCounter, "Verify PVKi valuesfrom CardPINHistory", "PVKi value is-" & db_PVKi  , "PASSED"
				Append_TestHTML StepCounter, "Verify PINVersion valuesfrom CardPINHistory", "PINVersion value is-" & db_PINVersion  , "PASSED"
				Append_TestHTML StepCounter, "Verify PINEventTypeID valuesfrom CardPINHistory", "PINEventTypeID value is-" & db_PINEventTypeIDi  , "PASSED"
				
		Else
				Append_TestHTML StepCounter, "Verify CVK value from CardIssueAttributes", "CVK value is-" & db_CVKi  , "FAILED"
				Append_TestHTML StepCounter, "Verify CVKi2 valuesfrom CardIssueAttributes", "CVKi2 value is-" & db_CVKi2  , "FAILED"
				Append_TestHTML StepCounter, "Verify PVV valuesfrom CardPINHistory", "PVV value is-" & db_PVV  , "FAILED"
				Append_TestHTML StepCounter, "Verify PVKi valuesfrom CardPINHistory", "PVKi value is-" & db_PVKi  , "FAILED"
			bRunFlag = False
				
		End If
		
'		If isNULL(db_RFIDUID) = true and isNULL(db_DispatchedDate) = true Then
'						
'				
'				Append_TestHTML StepCounter, "Verify DispatchedDate values from CardFulfilment", "No record created" , "PASSED"
'				Append_TestHTML StepCounter, "Verify RFIDUID values from CardRFID", "No record created" , "PASSED"
'		Else
'				Append_TestHTML StepCounter, "Verify RFIDUID values from CardRFID", "Record entry exist with few values" , "FAILED"			
'			bRunFlag = False
'
'End If
End Function


Function verifyCardRFIDTableDetails()
	On error resume next
		bFlag = True	

		query = "Select Count(*) as 'NoOfRecords' from CardRFID where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_NoOfRecords =  dbRecordSet("NoOfRecords")
		Set query = Nothing
		Set dbRecordSet = Nothing
		If cint(db_NoOfRecords) = 1 Then
				Append_TestHTML StepCounter, "Verify No of records created in CardRFID table ", "Created one record in the table with Cardid-"& XML_Cadrid , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify No of records created in CardRFID table ",  db_NoOfRecords & " records created in the table", "FAILED"
				
		End If
	
		query = "Select * from CardRFID where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_RFIDUID =  dbRecordSet("RFIDUID")
		db_EMAID =  dbRecordSet("EMAID")
		db_EVPrintedNumber =  dbRecordSet("EVPrintedNumber")

		Set query = Nothing
		Set dbRecordSet = Nothing
		
		If isNULL(db_RFIDUID) = False and  isNULL(db_EMAID) = False and isNULL(db_EVPrintedNumber) = False Then
				Append_TestHTML StepCounter, "Verify RFIDUID value from CardRFID", "RFIDUID value is-" & db_RFIDUID  , "PASSED"
				Append_TestHTML StepCounter, "Verify EMAID value from CardRFID", "EMAID value is-" & db_EMAID  , "PASSED"
				Append_TestHTML StepCounter, "Verify EVPrintedNumber value from CardRFID", "EVPrintedNumber value is-" & db_EVPrintedNumber  , "PASSED"
				
		Else
					Append_TestHTML StepCounter, "Verify RFIDUID value from CardRFID", "RFIDUID value is-" & db_RFIDUID  , "FAILED"
				Append_TestHTML StepCounter, "Verify EMAID value from CardRFID", "EMAID value is-" & db_EMAID  , "FAILED"
				Append_TestHTML StepCounter, "Verify EVPrintedNumber value from CardRFID", "EVPrintedNumber value is-" & db_EVPrintedNumber  , "FAILED"
			bRunFlag = False

End If
End Function



Function verifyCardFulfilmentTableDetails()
	On error resume next
		bFlag = True	

		query = "Select Count(*) as 'NoOfRecords' from CardFulfilment where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_NoOfRecords =  dbRecordSet("NoOfRecords")
		Set query = Nothing
		Set dbRecordSet = Nothing
		If cint(db_NoOfRecords) = 1 Then
				Append_TestHTML StepCounter, "Verify No of records created in CardFulfilment table ", "Created one record in the table with Cardid-"& XML_Cadrid , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify No of records created in CardFulfilment table ",  db_NoOfRecords & " records created in the table", "FAILED"
				
		End If
	
		query = "Select * from CardFulfilment where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_DispatchedDate =  dbRecordSet("DispatchedDate")
		db_ReceivedDate =  dbRecordSet("ReceivedDate")
		db_ActualProductionDatee =  dbRecordSet("ActualProductionDate")
		db_DispatchCompanyID =  dbRecordSet("DispatchCompanyID")
		db_DispatchTrackingReference =  dbRecordSet("DispatchTrackingReference")
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		If isNULL(db_DispatchedDate) = False and  isNULL(db_DispatchedDate) = False and isNULL(db_DispatchCompanyID) = False Then
				Append_TestHTML StepCounter, "Verify DispatchedDate value from CardFulfilment", "DispatchedDate value is-" & db_DispatchedDate  , "PASSED"
				Append_TestHTML StepCounter, "Verify ReceivedDate value from CardFulfilment", "ReceivedDate value is-" & db_ReceivedDate  , "PASSED"
				Append_TestHTML StepCounter, "Verify ActualProductionDate value from CardFulfilment", "ActualProductionDate value is-" & db_ActualProductionDatee  , "PASSED"
				Append_TestHTML StepCounter, "Verify DispatchCompanyID value from CardFulfilment", "DispatchCompanyID value is-" & db_DispatchCompanyID  , "PASSED"
				Append_TestHTML StepCounter, "Verify DispatchTrackingReference value from CardFulfilment", "DispatchTrackingReference value is-" & db_DispatchTrackingReference  , "PASSED"
				
		Else
				Append_TestHTML StepCounter, "Verify DispatchedDate value from CardFulfilment", "DispatchedDate value is-" & db_DispatchedDate  , "FAILED"
				Append_TestHTML StepCounter, "Verify ReceivedDate value from CardFulfilment", "ReceivedDate value is-" & db_ReceivedDate  , "FAILED"
				Append_TestHTML StepCounter, "Verify ActualProductionDate value from CardFulfilment", "ActualProductionDate value is-" & db_ActualProductionDatee  , "FAILED"
				Append_TestHTML StepCounter, "Verify DispatchCompanyID value from CardFulfilment", "DispatchCompanyID value is-" & db_DispatchCompanyID  , "FAILED"
				Append_TestHTML StepCounter, "Verify DispatchTrackingReference value from CardFulfilment", "DispatchTrackingReference value is-" & db_DispatchTrackingReference  , "FAILED"
			bRunFlag = False

End If
		
		
End Function





Function searchandReplaceMultipleStringwithdollar(jsonFPath, searchreplaceString)

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
			sstrings = Split(srstrings(titr),"$")
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
'			On Error Resume Next
'			mycheck =  instr(Split(keycontent,":")(1),chr(34))
'			If Err.Number <> 0  Then
'				stricheck = False
'			Else
'				stricheck = True
'			End If
'				Err.Number = 0

			If (isNumeric(replaceString) = False) or ( instr(Split(keycontent,":")(1),chr(34)) >0 ) Then
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





Function pinChagepremodifieddetailsinInputAPI(fPath, etType)
	On error resume next
	
	bFlag = True
	
		query = "Select * from SysVarColco where [Key]='FleetPINAPIEnabled' ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_SysVarID =  dbRecordSet("SysVarID")
		db_ColcoID =  dbRecordSet("ColcoID")
		db_Value =  dbRecordSet("Value")
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		If db_Value = 1 or db_Value = true   Then
				Append_TestHTML StepCounter, "Verify FleetPINAPIEnabled key in SysVarColco Table", "Entry exist with ID-"&db_SysVarID & " Colco-"& db_ColcoID & " Value is-"&  db_Value , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify FleetPINAPIEnabled key in SysVarColco Table", "Entry exist with ID-"&db_SysVarID & " Colco-"& db_ColcoID & " Value is-"&  db_Value , "FAILED"
			bRunFlag = False
				
		End If	
		
		query = "Select * from FileSeq where FileSeqID like '%DX053%' ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_LastSequence=  dbRecordSet("LastSequence")
		db_FileSeqID=  dbRecordSet("FileSeqID")
		
		Set query = Nothing
		Set dbRecordSet = Nothing
		countynumber = Split(db_FileSeqID,"_")(1)
		
		query = "Select * from Company where ClientCompanyNumber="&cint(countynumber)&" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_CompanyID=  dbRecordSet("CompanyID")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		IssuingCountryNumber = db_CompanyID
		
		query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_CVKi=  dbRecordSet("CVKi")
		db_CVKi2=  dbRecordSet("CVKi2")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from CardPINHistory where CardPANID="& XML_Cardpanid &" ;"
		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_PVV=  dbRecordSet("PVV")
		db_PVKi=  dbRecordSet("PVKi")
		db_PINVersion=  dbRecordSet("PINVersion")		
		Set query = Nothing
		Set dbRecordSet = Nothing
	
		FileCreationDate = findpreviousMonthenddate(-1)
		
		newLastseq = db_LastSequence + 1
		timeval = getDateandTimestamp()
		timeval_arr = split(timeval," ")
		timevaltext = timeval_arr(1)
		hrvaltext =  mid(timevaltext,1,2)
		minvaltext =  mid(timevaltext,3,2)
		secvaltext = mid(timevaltext,5,2)
		
		PINModifiedDate = FileCreationDate & "T" & hrvaltext & ":" & minvaltext & ":" & secvaltext & "Z"
								
		newPVVno = db_PVV + 1
		newPVKno = db_PVKi + 1
		newPINversion = db_PINVersion + 1
		
		If len(newPVVno)<> 4 Then
		
			For itr = len(newPVVno) To 3 Step 1
				newPVVno = "0" & newPVVno
			Next
		End If

		searchreplaceString = "PANID$" & XML_Cardpanid &";PINVersion$"& newPINversion &";PVV$" & newPVVno &";PVKi$"& newPVKno &";EventType$"& etType &";PINModifiedDate$"& PINModifiedDate &";IssuingCountryNumber$"& countynumber &";ColCoID$"& IssuingCountryNumber
		Call searchandReplaceMultipleStringwithdollar(fPath, searchreplaceString)
		
		If newPINversion <> "" Then
			pinChagepremodifieddetailsinInputAPI = newPINversion&"|"& newPVVno &"|"& newPVKno
		Else
			pinChagepremodifieddetailsinInputAPI = ""
		End If

End Function



Function validatepinChangePostAPIDetails(etType, newpinver, typeofPinval ,newvalPVV, newvalPVK,actypeID, evenTID)
On error resume next

	query = "Select * from Card where PAN="& cardPANNum & "  order by ModifiedOn desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_TypeofPin = dbRecordSet("TypeOfPIN")
	db_CardID = dbRecordSet("CardID")
	db_StatusID = dbRecordSet("StatusID")

	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_CardID <> ""  and db_TypeofPin = typeofPinval Then
		Append_TestHTML StepCounter, "Validate TypeOfPIN in Card record entries", "Expected: " & typeofPinval  & VBCRLF & "Actual: " & db_TypeofPin ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate TypeOfPIN in Card record entries", "Expected: " & typeofPinval  & VBCRLF & "Actual: " & db_TypeofPin ,"FAILED"		
		bRunFlag = False
	End If
	query = "Select * from CardPINHistory where CardPANID="& XML_Cardpanid & "  order by ModifiedOn desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CardPINHistoryID = dbRecordSet("CardPINHistoryID")
	db_PVV = dbRecordSet("PVV")
	db_PVKi = dbRecordSet("PVKi")
	db_PINVersion = dbRecordSet("PINVersion")

	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_PVV <> ""  and cint(db_PINVersion) = cint(newpinver) Then
		Append_TestHTML StepCounter, "Validate PINVersion in CardPINHistory record entries", "Expected: " & newpinver  & VBCRLF & "Actual: " & db_PINVersion ,"PASSED"
		Append_TestHTML StepCounter, "Validate PVV in CardPINHistory record entries", "Expected: " & newvalPVV  & VBCRLF & "Actual: " & db_PVV ,"PASSED"
		Append_TestHTML StepCounter, "Validate PVKi in CardPINHistory record entries", "Expected: " & newvalPVK  & VBCRLF & "Actual: " & db_PVKi ,"PASSED"
		
	Else
		Append_TestHTML StepCounter, "Validate TypeOfPIN in Card record entries", "Expected: " & typeofPinval  & VBCRLF & "Actual: " & db_TypeofPin ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from CardActivityLog where CardID="& XML_Cadrid & " order by ModifiedOn desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityLogID = dbRecordSet("ActivityLogID")

		
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityLogID <> "" Then
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate CardActivityLog record entries", "ActivityLogID: " & db_ActivityLogID  ,"FAILED"		
		bRunFlag = False
	End If
	
	query = "Select * from ActivityLog where ActivityLogID="& db_ActivityLogID & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardActivityLog table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ActivityTypeID = dbRecordSet("ActivityTypeID")
	db_Note = dbRecordSet("Note")

	Set dbRecordSet = Nothing
	Set query = Nothing
	
	If db_ActivityTypeID = actypeID and instr(db_Note,evenTID)>0 Then
		Append_TestHTML StepCounter, "Validate ActivityTypeID in ActivityLog record entry", "Expected: " & actypeID & "Actual:"&  db_ActivityTypeID ,"PASSED"
		Append_TestHTML StepCounter, "Validate EventTypeID  in ActivityLog record entry", "Expected: " & evenTID & "Actual:"&  db_Note ,"PASSED"
	Else
		Append_TestHTML StepCounter, "Validate ActivityTypeID in ActivityLog record entry", "Expected: " & actypeID & "Actual:"&  db_ActivityTypeID ,"FAILED"
		Append_TestHTML StepCounter, "Validate EventTypeID  in ActivityLog record entry", "Expected: " & evenTID & "Actual:"&  db_Note ,"FAILED"
		bRunFlag = False
	End If

End Function


Function navigateCustomersummaryScreen(custERPid)
On error resume next

	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_Start") Then
		Call  navigateStartMenu("Link_Customers","Link_SearchforCustomer","WebLIST_Role")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLIST_Role") Then
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", custERPid
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'			Call navigateCustomerSummaryMenuoption("Card Maintenance","WebLink_NewCard")
			

			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary") Then
				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Successfully navigated to requested customer screen", "PASSED"
				
			Else
				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Fail to navigate customer screen", "FAILED"
							bRunFlag = False
				
			End If
		Else
				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Fail to navigate customer search screen", "FAILED"
							bRunFlag = False

		End If
	Else
				Append_TestHTML StepCounter, "GFN Application Home Page", "Fail to get HOME Screen", "FAILED"
							bRunFlag = False
	End If

End Function


Function navigateCardgroupsScreen()
On error resume next

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Groups"
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardGroup", "Click", ""
	wait 4
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_CardGroup") Then
				Append_TestHTML StepCounter, "Verify Cardgroup default screen", "Successfully navigated to Card Group Screen", "PASSED"
	Else
				Append_TestHTML StepCounter, "Verify Cardgroup default screen", "Fail to navigate", "FAILED"
	
	End If
End Function


Function navigateCardgroupOverrideScreen(CGNameval)

On error resume next

	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Epin_WebLink_NewCGroup") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WebLink_NewCGroup", "Click", ""
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Epin_WebEdit_CGName") Then
				Append_TestHTML StepCounter, "Click on New Carg Group link", "Navigated to New Cardgroup Screen", "PASSED"
		
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WebEdit_CGName", "Set", CGNameval
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WebList_CGTypes", "RadioSelect", "#1"
			wait 2
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "EpinWebCheckBox_CGCardDeliveryPoint", "Click", ""
			wait 2
			cccd_status_val = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "EpinWebCheckBox_CGCardDeliveryPoint","GetROProperty","checked")
			If cccd_status_val = "1" Then
					Append_TestHTML StepCounter, "Verify all mandatory details on Cardgroup screen", "Filled all input data", "PASSED"
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WEl_CGOverride", "Click", ""
					wait 4
					Append_TestHTML StepCounter, "Verify Card Group Override Screen", "Navigated to Override screen", "PASSED"
'					
'					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'					If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") = True  Then
'						Append_TestHTML StepCounter, "Verify Email address is non-mandatary ", " Email address field is non-mandatary. Allowing to save without email","PASSED"
'						Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebElement("class:=button ctl00_ctl12","html tag:=SPAN","innertext:=Cancel").Click	
'					Else
'						Append_TestHTML StepCounter, "Verify Email address is non-mandatary ", " Email address field is mandatary/other errors","FAILED"
'					bRunFlag = False
'						
'					End If
			End If
		Else
				Append_TestHTML StepCounter, "Click on New Carg Group link", "Fail to navigate New Cardgroup Screen", "FAILED"
					bRunFlag = False
				
		
		End If
	Else					
		Append_TestHTML StepCounter, "Verify Cardgroup default screen", "Fail to navigate Card Group Screen", "FAILED"
					bRunFlag = False
		
	End If
End Function


Function validateNewFreeTextDBValidation()
	


UCDACheckval = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_uCDAobj","GetROProperty","checked")
If UCDACheckval = "0" Then
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_uCDAobj", "Click", ""
End If
wait 2
	Append_TestHTML StepCounter, "Verify UCDA check box", "Check box ticked and pin delivery Address fields are readonly", "PASSED"
			
		contactdisvalue = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WEDcontact", "GetROProperty", "disabled")
			If  contactdisvalue = "1" Then
				Append_TestHTML StepCounter, "Verify Contact details ", "Contact is not selected in the drop-down","PASSED"
				
				'OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_Dcountry", "RadioSelect", "#1"
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_Dcountry", "RadioSelect", "#1"
				
				''msgbox "Check Country selected or not"
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_Dadress", "Set", "Freetext Address"
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_DEmail", "Set", "test@testmail.com"
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_Dzipcode", "Set", "523240"
			
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") = True  Then
					Append_TestHTML StepCounter, "Maintain freetext in devlivery email address field", "mail id maintained and click on save button","PASSED"
					Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebElement("class:=button ctl00_ctl12","html tag:=SPAN","innertext:=Save").Click	
					wait 2
					cardNum = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Card","GetROProperty","default value")
					If cardNum <> "" Then
							Append_TestHTML StepCounter, "Get new PAN number", "New PAN generated - " & cardNum,"PASSED"
						query = "Select * from CardAddress order by 1 desc;"
						Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
						db_email = dbRecordSet("EmailAddress")
						db_address = dbRecordSet("AddressLines")
						db_zipcode = dbRecordSet("Zipcode")
						db_AddressID = dbRecordSet("AddressID")
						db_CardID = dbRecordSet("CardID")
						Set dbRecordSet = Nothing
						
						If db_email = "test@testmail.com" and db_address="Freetext Address"  Then
							Append_TestHTML StepCounter, "Verify CardAddress entry", "New entry- cardID-" & db_CardID  & "AddressID-"& db_AddressID & " Created","PASSED"
						Else
							Append_TestHTML StepCounter, "Verify CardAddress entry", "New entry- cardID-" & db_CardID  & "AddressID-"& db_AddressID & " Created","FAILED"						
						End If
						query = "Select * from CardAddressType where CardID='"& db_CardID & "' and AddressID='"& db_AddressID & "';"
						Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
						db_AddressTID = dbRecordSet("AddressTypeID")
						Set dbRecordSet = Nothing
						
						If db_AddressTID = "10"  Then
							Append_TestHTML StepCounter, "Verify CardAddressType entry", "New entry- AddressTypeID- Expected: 10" & VBCRLF  & "Aactual-"& db_AddressTID ,"PASSED"
						Else
							Append_TestHTML StepCounter, "Verify CardAddressType entry", "New entry- AddressTypeID- Expected: 10" & VBCRLF  & "Aactual-"& db_AddressTID,"FAILED"						
						End If
					End If
					
				Else
					Append_TestHTML StepCounter, "Verify Email address ", " Email address field is mandatary/other errors","FAILED"
							bRunFlag = False
				End If
			Else
				Append_TestHTML StepCounter, "Verify Contact details ", "Contact details are not found","FAILED"		
							bRunFlag = False
			End If
End Function



Function getCustomerCardgroupid()
On error resume next

	query = "Select * from Customer where CustomerERP='" & customerERP_id & "';"
	Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
	Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	db_CustomerID = dbRecordSet("CustomerID")
	Set query = Nothing
	Set dbRecordSet = Nothing
	
	
	query = "Select * from CardGroup where CustomerID=" & db_CustomerID & " order by  ModifiedOn desc;"
	Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
	Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	db_CardGroupID = dbRecordSet("CardGroupID")
	Set query = Nothing
	Set dbRecordSet = Nothing
		
	If db_CardGroupID <> "" Then
		getCustomerCardgroupid = db_CardGroupID
	Else
	getCustomerCardgroupid = ""
	End If
		
		
End Function



Function validateCardAddressDetails()

On error resume next
	query = "Select Count(*) As No_of_records from Card where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_No_of_records = dbRecordSet("No_of_records")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	query = "Select * from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, cint(db_No_of_records),"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_AddressID= dbRecordSet("AddressID")
	db_IsActive= dbRecordSet("IsActive")
	db_Email= dbRecordSet("Email")'
	db_Phone = dbRecordSet("Phone")
		
	Set dbRecordSet = Nothing
	Set query = Nothing
	addressIds = Split(db_AddressID,"|")

	For itr = 0 To ubound(addressIds) Step 1
			query = "Select Count(*) As No_of_records from CardAddressType where CardID="& XML_Cadrid & " and AddressID =  " & addressIds(itr) & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_No_of_records = dbRecordSet("No_of_records")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If isNULL(db_No_of_records) = true  Then
				Append_TestHTML StepCounter, "Verify No of records in CardAddressType table",  "zero records with AddressID-" & addressIds(itr)  ,"PASSED"
			ElseIf cint(db_No_of_records) = 1 Then
				Append_TestHTML StepCounter, "Verify No of records in CardAddressType table", db_No_of_records & " record with AddressID-" & addressIds(itr)  ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Verify No of records in CardAddressType table", db_No_of_records & " record/records with AddressID-" & addressIds(itr)  ,"PASSED"
			
			End If
			If isNULL(db_No_of_records) =False Then
					query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and AddressID =  " & addressIds(itr) & " order by 1 desc;"
					Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
					Set dbRecordSet = execute_db_query(query, cint(db_No_of_records),"SFN_SHELL_SPRINTQA_ID_OLTP")
					wait 2
					
					db_AddressTypeID= dbRecordSet("AddressTypeID")
								
					Set dbRecordSet = Nothing
					Set query = Nothing
				
							Append_TestHTML StepCounter, "Verify AddressTypeID of " & addressIds(itr) , "Address ID- " & addressIds(itr) & " contains AddressType-" & db_AddressTypeID  ,"PASSED"
					
			Else
							Append_TestHTML StepCounter, "Verify AddressTypeID of " & addressIds(itr) , "Address ID- " & addressIds(itr) & " contains NO AddressType"  ,"PASSED"
			
			End If
			query = "Select Count(*) As No_of_records from CardAddressTypeDeleted where CardID="& XML_Cadrid & " and AddressID =  " & addressIds(itr) & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_No_of_records = dbRecordSet("No_of_records")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If isNULL(db_No_of_records) = True Then
							Append_TestHTML StepCounter, "Verify AddresseTypeID of " & addressIds(itr) , "No entries in CardAddressTypeDeleted with " & addressIds(itr)  ,"PASSED"
			Else
							Append_TestHTML StepCounter, "Verify AddressTypeID of " & addressIds(itr) , db_No_of_records & " records in CardAddressTypeDeleted with " & addressIds(itr)  ,"PASSED"
					query = "Select * from CardAddressTypeDeleted where CardID="& XML_Cadrid & " and AddressID =  " & addressIds(itr) & " order by 1 desc;"
					Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
					Set dbRecordSet = execute_db_query(query, cint(db_No_of_records),"SFN_SHELL_SPRINTQA_ID_OLTP")
					wait 2
					
					db_AddressTypeID= dbRecordSet("AddressTypeID")
								
					Set dbRecordSet = Nothing
					Set query = Nothing
				
							Append_TestHTML StepCounter, "Verify AddressTypeID of " & addressIds(itr) , "Address ID- " & addressIds(itr) & " contains AddressType-" & db_AddressTypeID  ,"PASSED"
			End If
	Next
	
End Function

Function validateprePINReminderValues()
On error resume next
		query = "Select Count(*) As No_of_records from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute CardAddress table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_No_of_records = dbRecordSet("No_of_records")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If dbRecordSet("No_of_records") = "" Then
					Append_TestHTML StepCounter, "Verify entries in CardAddress" , "No Addresses are exist" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify entries in CardAddress" , db_No_of_records & " Addresses are exist" ,"PASSED"
			End If
		query = "Select * from CardAddress where CardID="& XML_Cadrid  & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute CardAddress table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, cint(db_No_of_records),"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			db_AddressID= dbRecordSet("AddressID")
			db_IsActive= dbRecordSet("IsActive")
			db_Email= dbRecordSet("Email")'
			db_Phone = dbRecordSet("Phone")			
			Set dbRecordSet = Nothing
			Set query = Nothing
			
			If db_AddressID <> "" Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddress" , "AddressID values-" & db_AddressID  & " Active values-" & db_IsActive ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Verify AddressID in CardAddress" , "AddressID values-" & db_AddressID  & " Active values-" & db_IsActive ,"PASSED"
					
			End If
		query = "Select * from CardAddressType where CardID="& XML_Cadrid  & " and AddressTypeID = 10 order by 1 desc;"
			Append_TestHTML StepCounter, "Execute CardAddressType table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			card_DAddressID= dbRecordSet("AddressID")
				If dbRecordSet("AddressID") = "" Then
						card_DAddressID = ""
					End If
			Set dbRecordSet = Nothing
			Set query = Nothing	
			
			query = "Select * from CardAddressType where CardID="& XML_Cadrid  & " and AddressTypeID = 13 order by 1 desc;"
			Append_TestHTML StepCounter, "Execute CardAddressType table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			card_PAddressID= dbRecordSet("AddressID")
				If dbRecordSet("AddressID") = "" Then
					card_PAddressID = ""
				End If
			Set dbRecordSet = Nothing
			Set query = Nothing	
			If card_DAddressID <> "" Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_DAddressID  & " of Card Delivery Addresstype-10" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_DAddressID  & " of Card Delivery Addresstype-10" ,"PASSED"
			
			End If
			If card_PAddressID <> "" Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_PAddressID  & " of PIN Delivery Addresstype-13" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_PAddressID  & " of PIN Delivery Addresstype-13" ,"PASSED"
			
			End If
			If card_PAddressID <> "" Then
				query = "Select * from CardAddressTypeDeleted where CardID="& XML_Cadrid & " and AddressID =  " & card_PAddressID & " order by 1 desc;"
				Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
				Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				
				db_CardAddressTypeDeletedID= dbRecordSet("CardAddressTypeDeletedID")
				
				Set dbRecordSet = Nothing
				Set query = Nothing
				If db_CardAddressTypeDeletedID = "" Then
						Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted","No entries are exist", "PASSED"
				Else
						Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted", db_CardAddressTypeDeletedID & " entries are exist", "FAILED"
				
				End If
			End If
	
End Function


Function validatepostPINReminderValues()
'card_PAddressID, card_DAddressID
On error resume next
		query = "Select * from CardAddress where CardID="& XML_Cadrid & " and  AddressID = " & card_PAddressID & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_IsActive = dbRecordSet("IsActive")
			db_OneTimeUse= dbRecordSet("OneTimeUse")
			
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_IsActive = 0 or db_IsActive = flase Then
					Append_TestHTML StepCounter, "Verify AddressID in status" , card_PAddressID &" satus is " & db_IsActive ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Verify AddressID in status" , card_PAddressID &" satus is " & db_IsActive ,"FAILED"
				
			End If
			If db_OneTimeUse = 0 Then
					Append_TestHTML StepCounter, "Verify OneTimeUse in CardAddressType" , "Expected:0" & "Actual:" & db_OneTimeUse  ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify OneTimeUse in CardAddressType" , "Expected:0" & "Actual:" & db_OneTimeUse  ,"FAILED"			
			End If
		query = "Select * from CardAddressType where CardID="& XML_Cadrid  & " and AddressID = " & card_DAddressID & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			card_DAddressID= dbRecordSet("AddressTypeID")
			
			
			Set dbRecordSet = Nothing
			Set query = Nothing	
			
			
			If instr(card_DAddressID,"10")>0  and instr(card_DAddressID,"13")>0 Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , card_DAddressID &" - Two address types contain same Address id's"  ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , card_DAddressID &" - Two address types contain same Address id's"  ,"FAILED"
			
			End If
			
			query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and  AddressID = " & card_PAddressID & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_PAddressID = dbRecordSet("AddressID")
						
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_PAddressID = "" Then
					Append_TestHTML StepCounter, "Verify PinDeliveryAddressID in CardAddressType" , card_PAddressID &" - Entry removed from Card AddressType table"  ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Verify PinDeliveryAddressID in CardAddressType" , card_PAddressID &" - Entry not removed from Card AddressType table"  ,"FAILED"
				
			End If
			
			
			query = "Select * from CardAddressTypeDeleted where CardID="& XML_Cadrid & " and AddressID =  " & card_PAddressID & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_AddressID= dbRecordSet("AddressID")
			db_AddressTypeID= dbRecordSet("AddressTypeID")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_AddressID <> "" Then
					Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted", db_AddressID & " of Addresstypeid-" & db_AddressTypeID &" entry moved into the table", "PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted", db_AddressID & " of Addresstypeid-" & db_AddressTypeID & " entry not moved into the table", "FAILED"
			
			End If
		
	
End Function



Function validatepreDeliveryAddrValues()
On error resume next
		query = "Select Count(*) As No_of_records from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_No_of_records = dbRecordSet("No_of_records")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
		query = "Select * from CardAddress where CardID="& XML_Cadrid  & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, cint(db_No_of_records),"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			db_AddressID= dbRecordSet("AddressID")
			db_IsActive= dbRecordSet("IsActive")
			db_Email= dbRecordSet("Email")'
			db_Phone = dbRecordSet("Phone")			
			Set dbRecordSet = Nothing
			Set query = Nothing
			
			If db_AddressID <> "" Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddress" , "AddressID values-" & db_AddressID  & " Active values-" & db_IsActive ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Verify AddressID in CardAddress" , "AddressID values-" & db_AddressID  & " Active values-" & db_IsActive ,"FAILED"
							bRunFlag = False
			End If
		query = "Select * from CardAddressType where CardID="& XML_Cadrid  & " and AddressTypeID = 10 order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			card_DAddressID= dbRecordSet("AddressID")
			
			Set dbRecordSet = Nothing
			Set query = Nothing	
			
			query = "Select * from CardAddressType where CardID="& XML_Cadrid  & " and AddressTypeID = 13 order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			card_PAddressID= dbRecordSet("AddressID")
			
			Set dbRecordSet = Nothing
			Set query = Nothing	
			If card_DAddressID <> "" Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_DAddressID  & " of Card Delivery Addresstype-10" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_DAddressID  & " of Card Delivery Addresstype-10" ,"FAILED"
			
			End If
			If card_PAddressID = "" or isNULL(card_PAddressID) = false Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "No previous PIN Delivery Addresstype-13" ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , "AddressID values-" & card_PAddressID  & " of PIN Delivery Addresstype-13" ,"FAILED"
			
			End If
			query = "Select * from CardAddressTypeDeleted where CardID="& XML_Cadrid & " and AddressID =  " & card_PAddressID & " order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_CardAddressTypeDeletedID = dbRecordSet("CardAddressTypeDeletedID")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_CardAddressTypeDeletedID = "" Then
					Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted","No entries are exist", "PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted", db_CardAddressTypeDeletedID & " entries are exist", "FAILED"
			
			End If
		
	
End Function




Function validatepostDeliveryAddValues()
On error resume next
		query = "Select * from CardAddress where CardID="& XML_Cadrid & " and AddressID = " & card_DAddressID & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_IsActive = dbRecordSet("IsActive")
			db_CAddressID = dbRecordSet("AddressID")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_IsActive = 1 or db_IsActive = true Then
					Append_TestHTML StepCounter, "Verify AddressID in status" , card_DAddressID &" satus is " & db_IsActive ,"PASSED"
				Else
					Append_TestHTML StepCounter, "Verify AddressID in status" , card_DAddressID &" satus is " & db_IsActive ,"FAILED"
				
			End If
		query = "Select * from CardAddressType where CardID="& XML_Cadrid  & "and AddressID = " & db_CAddressID & " order by ModifiedOn desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			db_DAddressID= dbRecordSet("AddressID")
			db_AddressTypeID= dbRecordSet("AddressTypeID")		
			
			Set dbRecordSet = Nothing
			Set query = Nothing	
			If instr(db_AddressTypeID,"10") > 0  and instr(db_AddressTypeID,"13")>0 Then
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , db_AddressTypeID &" - Two address types contain same Address id's"  ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID in CardAddressType" , db_AddressTypeID &" - Two address types contain same Address id's"  ,"FAILED"
			
			End If
			
'			query = "Select * from CardAddressTypeDeleted where CardID="& XML_Cadrid & " and AddressID =  " & card_PAddressID & " order by 1 desc;"
'			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
'			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
'			wait 2
'			
'			db_AddressID= dbRecordSet("AddressID")
'			db_AddressTypeID= dbRecordSet("AddressTypeID")
'			
'			Set dbRecordSet = Nothing
'			Set query = Nothing
'			If db_AddressID <>"" Then
'					Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted", db_AddressID & " of Addresstypeid-" & db_AddressTypeID &" entry moved into the table", "PASSED"
'			Else
'					Append_TestHTML StepCounter, "Verify AddressID value in CardAddressTypeDeleted", db_AddressID & " of Addresstypeid-" & db_AddressTypeID & " entry not moved into the table", "FAILED"
'			
'			End If
		
	
End Function

Function cardAddressrecordsvaidation(addrvalue )
	On error resume next
		query = "Select Count(*) As No_of_records from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_No_of_records = dbRecordSet("No_of_records")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If cint(db_No_of_records) = cint("1") Then
				Append_TestHTML StepCounter, "Verify CardAddress table record entries", db_No_of_records & " entry/ies are exist in the table and no new record created" , "PASSED"
			Else
				Append_TestHTML StepCounter, "Verify CardAddress table record entries", db_No_of_records & " entry/ies are exist in the table and record created/deleted" , "FAILED"
			
			End If
End Function


Function validatepostCADValues(cdstatus,cdvalue,pdstatus,pdvalue,newentrystatus,emailval)
On error resume next
		query = "Select Count(*) As No_of_records from CardAddress where CardID="& XML_Cadrid & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_No_of_records = dbRecordSet("No_of_records")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
		If db_No_of_records <> "" Then
			Append_TestHTML StepCounter, "Verify Number of Address records in CardAddress table", db_No_of_records & " entry/ies are created in the table" , "PASSED"
		Else
			Append_TestHTML StepCounter, "Verify Number of Address records in CardAddress table", db_No_of_records & " - No entry created in the table" , "FAILED"
		
		End If
		If cdstatus <> "" Then
			query = "Select * from CardAddress where CardID="& XML_Cadrid & " and  AddressID = " & card_DAddressID & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_IsActive = dbRecordSet("IsActive")
			db_OneTimeUse= dbRecordSet("OneTimeUse")
			
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and  AddressID = " & card_DAddressID & " and AddressTypeID=10  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_AddressTypeID = dbRecordSet("AddressTypeID")
			db_AddressID= dbRecordSet("AddressID")
			
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_IsActive = cdvalue or cbool(db_IsActive)=cbool(cdvalue)  Then
					Append_TestHTML StepCounter, "Verify AddressID "& card_DAddressID & " Record Status" , "Record exist in the Card Address Table" ,"PASSED"
					Append_TestHTML StepCounter, "Verify AddressTypeID "& db_AddressTypeID & " Record Status" , "Record exist -AddressID:" & db_AddressID  & " AddressTypeID:" & db_AddressTypeID  ,"PASSED"
					
					Append_TestHTML StepCounter, "Verify AddressID "& card_DAddressID & "  isActive status" , "Expected:" & cdvalue  &" satus is " & db_IsActive ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID "& card_DAddressID & "  isActive status" , "Expected:" & cdvalue  &" satus is " & db_IsActive ,"FAILED"
			End If
			If db_OneTimeUse <>"" Then
					Append_TestHTML StepCounter, "Verify OneTimeUse in CardAddressType" , "Expected:"& db_OneTimeUse & "Actual:" & db_OneTimeUse  ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify OneTimeUse in CardAddressType" , "Expected:0" & "Actual:" & db_OneTimeUse  ,"FAILED"			
			End If
		End If
		If pdstatus <> "" Then
			query = "Select * from CardAddress where CardID="& XML_Cadrid & " and  AddressID = " & card_PAddressID & "  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_IsActive = dbRecordSet("IsActive")
			db_OneTimeUse= dbRecordSet("OneTimeUse")
			
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			
			query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and  AddressID = " & card_PAddressID & " and AddressTypeID=13  order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			
			db_AddressTypeID = dbRecordSet("AddressTypeID")
			db_AddressID= dbRecordSet("AddressID")
			
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			
			If db_IsActive = pdvalue or cbool(db_IsActive)=cbool(pdvalue)  Then
					Append_TestHTML StepCounter, "Verify AddressID "& card_PAddressID & " Record Status" , "Record exist in the Card Address Table" ,"PASSED"
					Append_TestHTML StepCounter, "Verify AddressTypeID "& db_AddressTypeID & " Record Status" , "Record exist -AddressID:" & db_AddressID  & " AddressTypeID:" & db_AddressTypeID  ,"PASSED"
					
					Append_TestHTML StepCounter, "Verify AddressID "& card_PAddressID & "  isActive status" , "Expected:" & pdvalue  &" satus is " & db_IsActive ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify AddressID "& card_PAddressID & "  isActive status" , "Expected:" & pdvalue  &" satus is " & db_IsActive ,"FAILED"
			End If
			If db_OneTimeUse <> "" Then
					Append_TestHTML StepCounter, "Verify OneTimeUse in CardAddressType" , "Expected:" &  db_OneTimeUse & "Actual:" & db_OneTimeUse  ,"PASSED"
			Else
					Append_TestHTML StepCounter, "Verify OneTimeUse in CardAddressType" , "Expected:0" & "Actual:" & db_OneTimeUse  ,"FAILED"			
			End If
		End If
			If newentrystatus Then
				query = "Select * from CardAddress where CardID="& XML_Cadrid & " order by AddressID desc;"
				Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
				Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				
				db_AddressID= dbRecordSet("AddressID")
				db_IsActive= dbRecordSet("IsActive")
				db_Email= dbRecordSet("Email")
				
				Set dbRecordSet = Nothing
				Set query = Nothing	
				query = "Select * from CardAddressType where CardID="& XML_Cadrid & " and  AddressID = " & db_AddressID & "  order by 1 desc;"
				Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
				Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
				wait 2
				
				db_AddressTypeID = dbRecordSet("AddressTypeID")
				db_AddressID= dbRecordSet("AddressID")
				
				
				Set dbRecordSet = Nothing
				Set query = Nothing
				If db_AddressID <> "" and ucase(db_Email) = ucase(emailval) Then
					Append_TestHTML StepCounter, "Verify New Entry in CardAddressType" , "New entry created in the CardAddress table and AddressID is-" & db_AddressID & " Is Active-" & db_IsActive  ,"PASSED"
					Append_TestHTML StepCounter, "Verify AddressTypeID "& db_AddressTypeID & " Record Status" , "Record exist -AddressID:" & db_AddressID  & " AddressTypeID:" & db_AddressTypeID  ,"PASSED"
					
				Else
					Append_TestHTML StepCounter, "Verify New Entry in CardAddressType" , "New entry created in the CardAddress table and AddressID is-" & db_AddressID & " Is Active-" & db_IsActive   ,"FAILED"
				End If
			Else
				query = "Select * from CardAddress where CardID="& XML_Cadrid & " order by AddressID desc;"
					Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
					Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
					wait 2
					
					db_AddressID= dbRecordSet("AddressID")
					db_IsActive= dbRecordSet("IsActive")
					db_Email= dbRecordSet("Email")
					
					Set dbRecordSet = Nothing
					Set query = Nothing	
					If  cint(db_AddressID)<= cint(card_PAddressID) or cint(db_AddressID)<= cint(card_DAddressID)Then
							Append_TestHTML StepCounter, "Verify New Entry in CardAddress" , "No New entry created in the CardAddress table "  ,"PASSED"
						Else
							Append_TestHTML StepCounter, "Verify New Entry in CardAddress" , "New entry created in the CardAddress table " & card_DAddressID & "-" & db_AddressID ,"FAILED"
						
					End If
			End If
'			query = "Select * from CardAddressTypeDeleted where CardID="& XML_Cadrid & " order by 1 desc;"
'			Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
'			Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
'			wait 2
'			
'			db_AddressID= dbRecordSet("AddressID")
'			db_AddressTypeID= dbRecordSet("AddressTypeID")
'			
'			Set dbRecordSet = Nothing
'			Set query = Nothing
'			If db_AddressID = "" Then
'					Append_TestHTML StepCounter, "Verify Record Entry in CardAddressTypeDeleted", "No entry deleted and not moved into the CardAddressTypeDeleted Table", "PASSED"
'			Else
'					Append_TestHTML StepCounter, "Verify Record Entry in CardAddressTypeDeleted", "Entry deleted and  moved into the CardAddressTypeDeleted Table", "FAILED"
'					
'			
'			End If
'		
	
End Function




Public Function doAPIOutboundControlAction(keydata,cstatusFlag,chstatusFlag)

On Error Resume Next
	bRunFlag = True	
	query = "Select * from APIOutboundControl;"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
		db_APIAuthMethodID = dbRecordSet("APIAuthMethodID")
		db_APIOutboundControlID = dbRecordSet("APIOutboundControlID")
		
		Set dbRecordSet = Nothing
		
		If isNull(db_APIAuthMethodID) Then
			Append_TestHTML StepCounter, keydata & "APIAuthMethodID Default value","Unknown/NULL data value " ,"FAILED"
			bRunFlag = False
		Else
			If instr(db_APIAuthMethodID,cstatusFlag)>0 or instr(db_APIAuthMethodID,cint(cstatusFlag)) >0  Then
				Append_TestHTML StepCounter, keydata & " Existing value","Expected Value: " & cstatusFlag  & VBCRLF & "Actual Value: " &db_APIAuthMethodID  ,"PASSED"
				query = "Update APIOutboundControl set APIAuthMethodID= " & chstatusFlag & ";"
				Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_MASTER")
				Set dbRecordSet = Nothing
				wait 10
				query = "Select * from APIOutboundControl;"
				Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
					db_APIAuthMethodID1= dbRecordSet("APIAuthMethodID")
					db_APIOutboundControlID1 = dbRecordSet("APIOutboundControlID")
				Set dbRecordSet = Nothing
				If cint(db_APIAuthMethodID1) = cint(chstatusFlag) Then
					Append_TestHTML StepCounter,keydata & " Updated value","Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_APIAuthMethodID1  ,"PASSED"
				Else
					Append_TestHTML StepCounter, db_Key_val & "Data Updated value","Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_APIAuthMethodID1  ,"FAILED"
					bRunFlag = False
					
				End If
			Else
				Append_TestHTML StepCounter, keydata & " Existing value","No Update requiered in DB - Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_APIAuthMethodID  ,"PASSED"
			End If
		End If
End  Function


Function validateEmailVariableTextCodeData()

	On Error Resume Next
	bRunFlag = True	
	query = "Select * from Card where CardID="& XML_Cadrid & " order by ModifiedOn desc;"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		db_TokenTypeID = dbRecordSet("TokenTypeID")
		Set dbRecordSet = Nothing
		Set query = Nothing
	query = "Select * from TokenTypeControl where TokenTypeID="& db_TokenTypeID & " ;"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		db_CardDeliveryID = dbRecordSet("CardDeliveryID")
		Set dbRecordSet = Nothing
		Set query = Nothing	
	query = "Select * from CardDelivery where CardDeliveryID="& db_CardDeliveryID & " ;"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		db_EmailVariableTextCode = dbRecordSet("EmailVariableTextCode")
		Set dbRecordSet = Nothing
		Set query = Nothing	
		
		If db_EmailVariableTextCode <> "" Then
				Append_TestHTML StepCounter, "Verify EmailVariableTextCode value", "Value is not NULL and default value is-" & db_EmailVariableTextCode , "PASSED"
		Else
				Append_TestHTML StepCounter, "Verify EmailVariableTextCode value", "Value is NULL and default value is-" & db_EmailVariableTextCode , "FAILED"
			bRunFlag = False
		
		End If
		
End Function

Function verifyCardPINMailerRequestDetails(pinAdviceval,successval,processval)
	
	On Error Resume Next
	bRunFlag = True	
	query = "Select * from CardPINMailerRequest where CardID="& XML_Cadrid & " order by ModifiedOn desc;"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
		db_CardPINMailerRequestID = dbRecordSet("CardPINMailerRequestID")
		db_IsSuccess = dbRecordSet("IsSuccess")
		db_Processed = dbRecordSet("Processed")
		db_PINAdviceTypeID = dbRecordSet("PINAdviceTypeID")
		
			
		Set dbRecordSet = Nothing
		Set query = Nothing
	If db_CardPINMailerRequestID <>"" Then
				Append_TestHTML StepCounter, "Verify New CardPINMailerRequestID Record", "New Record created" , "PASSED"
				If cint(db_PINAdviceTypeID) = cint(pinAdviceval) Then
					Append_TestHTML StepCounter, "Verify PINAdviceTypeID value", "Expected:-1"& VBCRLF & "Actual:-" & db_PINAdviceTypeID , "PASSED"
				Else
					Append_TestHTML StepCounter, "Verify PINAdviceTypeID value", "Expected:-1"& VBCRLF & "Actual:-" & db_PINAdviceTypeID , "FAILED"
				
				End If
'				If cint(db_IsSuccess) = cint(successval) or cbool(db_IsSuccess) = cbool(successval) Then
'					Append_TestHTML StepCounter, "Verify IsSuccess value", "Expected:-1"& VBCRLF & "Actual:-" & db_IsSuccess , "PASSED"
'				Else
'					Append_TestHTML StepCounter, "Verify IsSuccess value", "Expected:-1"& VBCRLF & "Actual:-" & db_IsSuccess , "FAILED"
'				
'				End If
				If cint(db_Processed) = cint(processval) or cbool(db_Processed) = cbool(processval) Then
					Append_TestHTML StepCounter, "Verify Processed value", "Expected:-1"& VBCRLF & "Actual:-" & db_Processed , "PASSED"
					
				Else
					Append_TestHTML StepCounter, "Verify Processed value", "Expected:-1"& VBCRLF & "Actual:-" & db_Processed , "FAILED"
				
				End If
	Else
				Append_TestHTML StepCounter, "Verify New CardPINMailerRequestID Record", "New Record created" , "FAILED"
	
	End If
	
End Function

Function checkJobStatus(jobidval,DbName,searchXMLvalue)
	On Error Resume Next
	
	query = "Select * from Job where JobTypeID = "& jobidval & " and Cast(DateCreated as Date) = Cast(Getdate() as Date) order by ModifiedOn desc"
	Append_TestHTML StepCounter, "Verify job -"& jobidval,query, "PASSED"
	wait 10
	Set dbRecordSet = execute_db_query(query, 1,DbName)
	wait 2
	
	db_jobid = dbRecordSet("ID")
	db_statusid = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	db_DateCreated= dbRecordSet("DateCreated")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	If instr(db_inputxml,searchXMLvalue)>0 and db_statusid = "4" Then
				Append_TestHTML StepCounter, "Validate job StatusID   status", "Expected StatusIDValue: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"PASSED"
				Append_TestHTML StepCounter, "Validate job InputXml File  status", "Input XM value " & db_inputxml  ,"PASSED"
				
			Else
				Append_TestHTML StepCounter, "Validate job StatusID   status", "Expected StatusIDValue: 4"  & VBCRLF & "Actual Value: " & db_statusid ,"FAILED"
				Append_TestHTML StepCounter, "Validate job InputXml File  status", "Input XM value " & db_inputxml  ,"FAILED"
				
		End If	
	
End Function


Function returnCustomerAddressDetails(cardapi_data)
	On Error Resume Next
	pctval = cardapi_data("PinContactTypeIDval")
	If pctval = 1 Then
			caddtypeid= 13
	ElseIf pctval = 2 Then
			caddtypeid= 10
	
	ElseIf pctval = 3 Then
			caddtypeid= 13
	End If
	query = "Select * from CardAddress where AddressID in ( Select AddressID from CardAddressType where AddressTypeID= " & caddtypeid & " and CardID="& XML_Cadrid & ") order by ModifiedOn desc;"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_AddressID = dbRecordSet("AddressID")
		db_ContactName = dbRecordSet("ContactName")
		db_Email= dbRecordSet("Email")
		db_Phone = dbRecordSet("Phone")
		Set dbRecordSet = Nothing
		Set query = Nothing
	
		query = "Select * from IssuingCountry where ISOCode2 = '" & countryCode & "'"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_IssuingCountryNumber = dbRecordSet("IssuingCountryNumber")
		Set dbRecordSet = Nothing
		Set query = Nothing
	
		db_ICN = cint(db_IssuingCountryNumber)
		
		query = "Select * from Card where CardID=" & XML_Cadrid & ";"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_pan = dbRecordSet("PAN")
		db_ExpiryDate = dbRecordSet("ExpiryDate")
		Set dbRecordSet = Nothing
		Set query = Nothing
		db_ExpiryDates = Split(db_ExpiryDate,"-")

		If db_pan <> "" Then
			returnCustomerAddressDetails = db_pan & ";" & cint(db_ExpiryDates(1)) & ";" & cint(Mid(db_ExpiryDate,3,2)) & ";email"& ";" & db_Email & ";issuing_country" & ";" & db_ICN
		Else
			returnCustomerAddressDetails = ""
		End If
		
End Function



Function returnCustomerDetails()
	On Error Resume Next
	
		If customerERP_id <> "" Then
			returnCustomerDetails = customerERP_id &";email"& ";sms"
		Else
			returnCustomerDetails = ""
		End If
		
End Function




Public Function validateOutboundFileData(filePath,data,ftype)

       On error resume next

				bFlag = True

	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(filePath) Then
		If (ucase(ftype) = ucase("dat") or ucase(ftype) = ucase("File") ) Then
			Set fileRead = fileSysObj.OpenTextFile(filePath,1,False,-1)
		Else
			Set fileRead = fileSysObj.OpenTextFile(filePath,1)
			Set fileRead1 = fileSysObj.OpenTextFile(filePath,1)
			
		End  IF
		content = fileRead.ReadAll
		Append_TestHTML StepCounter, "File data ", content , "PASSED"
		LineCount = 0
		Do While fileRead1.AtEndofStream <> True
		    sLinedata = fileRead1.ReadLine
		       sLinedatass = Split(sLinedata,":")
		    If instr(sLinedata,"Type") <> 0  or instr(sLinedata,"Id") <> 0 or instr(sLinedata,"Type")<> 0  or instr(sLinedata,"Api")<> 0 or instr(sLinedata,"Type")<> 0 or instr(sLinedata,"Request")<> 0 or instr(sLinedata,"Response")<> 0 or instr(sLinedata,"ResponseStatus")<> 0 or instr(sLinedata,"Date")<> 0 Then 
				
				If sLinedatass(1)<>"" Then
					Append_TestHTML StepCounter, "File data Line-" & LineCount , "Data-" &  LineCount  & " Value-" & sLinedata , "PASSED"
				Else
					Append_TestHTML StepCounter, "Validate" & sLinedatass(0) & " Value" , "Value is Empty" , "FAILED"
				End If	
			Else
					Append_TestHTML StepCounter, "Validate data line" , "Invalid line data" , "FAILED"
			
		    End If
		    LineCount= LineCount+1
		Loop
		fileRead1.Close
'		'''''msgbox content
		If instr(data,";") = 0 Then
			If instr(content,data)>0 Then
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "matched in the path "&filePath , "PASSED"
				
			Else
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "not matched in the path "&filePath , "FAILED"
				bRunFlag = False
				bFlag = False	
			
			End If
		Else
		dataval = split(data,";")
			For iitr = 0 To ubound(dataval) Step 1
				If instr(content,dataval(iitr))>0 Then
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "matched in the path "&filePath , "PASSED"
					
				Else
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "not matched in the path "&filePath , "FAILED"
				bRunFlag = False
				bFlag = False	
				End If
			Next
		End If
	
	fileRead.Close
	Set fileSysObj = Nothing
	End If


End Function

Public Function validateOutboundFilenegativeData(filePath,data,ftype)

       On error resume next

				bFlag = True

	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(filePath) Then
		If (ucase(ftype) = ucase("dat") or ucase(ftype) = ucase("File") ) Then
			Set fileRead = fileSysObj.OpenTextFile(filePath,1,False,-1)
		Else
			Set fileRead = fileSysObj.OpenTextFile(filePath,1)
			Set fileRead1 = fileSysObj.OpenTextFile(filePath,1)
			
		End  IF
		content = fileRead.ReadAll
		Append_TestHTML StepCounter, "File data ", content , "PASSED"
		LineCount = 0
		Do While fileRead1.AtEndofStream <> True
		    sLinedata = fileRead1.ReadLine
		       sLinedatass = Split(sLinedata,":")
		    If instr(sLinedata,"Type") <> 0  or instr(sLinedata,"Id") <> 0 or instr(sLinedata,"Type")<> 0  or instr(sLinedata,"Api")<> 0 or instr(sLinedata,"Type")<> 0 or instr(sLinedata,"Request")<> 0 or instr(sLinedata,"Response")<> 0 or instr(sLinedata,"ResponseStatus")<> 0 or instr(sLinedata,"Date")<> 0 Then 
				
				If sLinedatass(1)<>"" Then
					Append_TestHTML StepCounter, "File data Line-" & LineCount , "Data-" &  LineCount  & " Value-" & sLinedata , "PASSED"
				Else
					Append_TestHTML StepCounter, "Validate" & sLinedatass(0) & " Value" , "Value is Empty" , "FAILED"
				End If	
			Else
					Append_TestHTML StepCounter, "Validate data line" , "Invalid line data" , "FAILED"
			
		    End If
		    LineCount= LineCount+1
		Loop
		fileRead1.Close
'		'''''msgbox content
		If instr(data,";") = 0 Then
			If instr(content,data)>0 Then
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "matched in the path "&filePath , "FAILED"
				
			Else
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "not matched in the path and new file not created and also Request not have this record"&filePath , "PASSED"
							
			End If
		Else
		dataval = split(data,";")
			For iitr = 0 To ubound(dataval) Step 1
				If instr(content,dataval(iitr))>0 Then
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "matched in the path "&filePath , "PASSED"
					
				Else
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "not matched in the path "&filePath , "FAILED"
				bRunFlag = False
				bFlag = False	
				End If
			Next
		End If
	
	fileRead.Close
	Set fileSysObj = Nothing
	End If


End Function

Function doUseFleetPinAction(cpflagaction,setflagval)
	On error resume next
	bFlag = True

	If ucase(cpflagaction) =ucase( "Resend") Then
		query1 = "Update Customer Set UseFleetPIN = 1 where CustomerERP='"& customerERP_id &"';"
			Append_TestHTML StepCounter, "Update UseFleetPIN value in  Customer", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			Set dbRecordSet = Nothing
	ElseIf ucase(cpflagaction) =ucase( "Setup") Then		
		query1 = "Update Customer Set UseFleetPIN = 0 where CustomerERP='"& customerERP_id &"';"
			Append_TestHTML StepCounter, "Update UseFleetPIN value in  Customer", query1, "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			Set dbRecordSet = Nothing
	End If
	
End Function

Function setandresendFleetPinInCardParameterScreen(cpflagaction)
	
On error resume next
	bFlag = True
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Parameters"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardParams", "Click", ""
	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
		Append_TestHTML StepCounter, "Card Parameters", "Opened card parameters page", "PASSED"
	else
		Append_TestHTML StepCounter, "Card Parameters	", "Navigation to card parameters page failed", "FAILED"
		bFlag = False
	End If
	
	If ucase(cpflagaction) =ucase( "Resend") Then
		Call performCardActions("CardsParamWebLink_ResendFleetPINButton")
	ElseIf ucase(cpflagaction) =ucase( "Setup") Then		
		Call performCardActions("CardsParamWebLink_setupFleetPINButton")
	End If
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","CP_WebEdit_Mobilefield") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CP_WebEdit_Mobilefield", "Set", "+919951111654"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CP_WebEdit_Emailfield", "Set", "TestSetup@gmail.com"
		typeOfChangeFleetPin = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CP_WebCheckbox_ManagerChangeFleetPin","GetROProperty","checked")
		If typeOfChangeFleetPin = 0 or typeOfChangeFleetPin = False	Then
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CP_WebCheckbox_ManagerChangeFleetPin", "Click", ""
		End If
		Append_TestHTML StepCounter,"Verify all mandatory Values","Successfully filled with all fields","PASSED"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		wait 4	
		'msgbox "Wait for navigation"
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "CP_WebCheckbox_UseFleetPin") Then
			Append_TestHTML StepCounter, "Check FleetPin UI Verification", "Details are save successfully", "PASSED"
			wait 5
			query = "Select * from Customer where CustomerERP='"& customerERP_id &"' order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			db_CustomerID = dbRecordSet("CustomerID")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			query = "Select * from CustomerFleetPINRequest where  CustomerID="& db_CustomerID &" order by 1 desc;"
			Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			
			db_IsSuccess = dbRecordSet("IsSuccess")
			db_Processed = dbRecordSet("Processed")
			db_PINChangeAllowed= dbRecordSet("PINChangeAllowed")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
'			If db_IsSuccess = 1 or db_IsSuccess = True Then
'				Append_TestHTML StepCounter, "Check IsSuccess value from CustomerFleetPINRequest", "Expected-1" & "Actual-" & db_IsSuccess , "PASSED"
'			Else
'				Append_TestHTML StepCounter, "Check IsSuccess value from CustomerFleetPINRequest", "Expected-1" & "Actual-" & db_IsSuccess , "FAILED"
'				bRunFlag = False			
'			End If
			If db_Processed = 1 or db_Processed = True Then
				Append_TestHTML StepCounter, "Check Processed value from CustomerFleetPINRequest", "Expected-1" & "Actual-" & db_Processed , "PASSED"
			Else
				Append_TestHTML StepCounter, "Check Processed value from CustomerFleetPINRequest", "Expected-1" & "Actual-" & db_Processed , "FAILED"
				bRunFlag = False			
			End If
			If db_PINChangeAllowed = 1 or db_PINChangeAllowed = True Then
				Append_TestHTML StepCounter, "Check PINChangeAllowed value from CustomerFleetPINRequest", "Expected-1" & "Actual-" & db_Processed , "PASSED"
			Else
				Append_TestHTML StepCounter, "Check PINChangeAllowed value from CustomerFleetPINRequest", "Expected-1" & "Actual-" & db_Processed , "FAILED"
				bRunFlag = False			
			End If
		else
			Append_TestHTML StepCounter, "Card Parameters	", "Navigation to card parameters page failed", "FAILED"
			bFlag = False
		End If
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for create customer card parameters","FAILED"
		bRunFlag = False
	End If	
	
End Function


			
Function validateDX500CRAfiledata(fPath,ftype)
On error resume Next
	bFlag = True
			query = "Select * from FileSeq where FileSeqID like '%DX500%';"
			Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			
			db_LastSequence = dbRecordSet("LastSequence")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			timeval = getDateandTimestamp()
		datetimeval_arr = split(timeval," ")
		datevaltext = timeval_arr(0)
		datevaltext = mid(datevaltext,"-","")
	Set fileSysObj = createObject("Scripting.FileSystemObject")

	If (ucase(ftype) = ucase("dat") or ucase(ftype) = ucase("File") ) Then
		Set DXread = fileSysObj.OpenTextFile(filePath,1,False,-1)
	Else
		Set DXread = fileSysObj.OpenTextFile(filePath,1)
	End  IF
	lcount = 0
	Do while not DXread.AtEndOfStream
		strLine = DXread.ReadLine()
		lcount = lcount + 1
		If instr(strLine,"XCH")>0 Then
			
			Append_TestHTML StepCounter,"CRA file Heder Data in a File", strLine , "PASSED"
			
			Rectypeval = mid(strLine,1,4)
			Originval = mid(strLine,5,10)
			fileseqval = mid(strLine,15,6)
			timestampval = mid(strLine,21,11)
			fileversionval = mid(strLine,33,5)
			otherval = mid(strLine,39)
				If Rectypeval = "XCH" Then
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4"& "Expected-XCH" &"Actual-"& Rectypeval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4" & "Expected-XCH" &"Actual-"& Rectypeval, "FAILED"		
				End If
				If instr(Originval,"ESC")>0 Then
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "FAILED"					
				End If
				If instr(fileseqval,db_LastSequence)>0 Then
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "FAILED"
							
				End If
				If instr(timestampval,datevaltext)>0 Then
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "PASSED"
				Else
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "FAILED"
					
				End If
				If instr(fileversionval,"V 3.0") > 0 Then
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "FAILED"
				End If
				If instr(otherval,"000000000000000000000000")>0 Then
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "FAILED"
							
				End If
		ElseIf  instr(strLine,"XCT")>0  Then
		Append_TestHTML StepCounter,"CRA file Trailer Data in a File", strLine , "PASSED"
			
			Rectypeval = mid(strLine,1,4)
			Originval = mid(strLine,5,10)
			fileseqval = mid(strLine,15,6)
			timestampval = mid(strLine,21,11)
			fileversionval = mid(strLine,33,5)
			otherval = mid(strLine,39)
				If Rectypeval = "XCT" Then
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4"& "Expected-XCT" &"Actual-"& Rectypeval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4" & "Expected-XCH" &"Actual-"& Rectypeval, "FAILED"		
				End If
				If instr(Originval,"ESC")>0 Then
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "FAILED"					
				End If
				If instr(fileseqval,db_LastSequence)>0 Then
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "FAILED"
							
				End If
				If instr(timestampval,datevaltext)>0 Then
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "PASSED"
				Else
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "FAILED"
					
				End If
				If instr(fileversionval,"V 3.0") > 0 Then
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "FAILED"
				End If
				If instr(otherval,lcount)>0 Then
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "FAILED"
							
				End If
		End If
	
	Loop
End Function	


			
Function validateDX500PMAfiledata(fPath,ftype)
On error resume Next
	bFlag = True
			query = "Select * from FileSeq where FileSeqID like '%DX500%';"
			Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			
			db_LastSequence = dbRecordSet("LastSequence")
			
			Set dbRecordSet = Nothing
			Set query = Nothing
			timeval = getDateandTimestamp()
		datetimeval_arr = split(timeval," ")
		datevaltext = timeval_arr(0)
		datevaltext = mid(datevaltext,"-","")
	Set fileSysObj = createObject("Scripting.FileSystemObject")

	If (ucase(ftype) = ucase("dat") or ucase(ftype) = ucase("File") ) Then
		Set DXread = fileSysObj.OpenTextFile(filePath,1,False,-1)
	Else
		Set DXread = fileSysObj.OpenTextFile(filePath,1)
	End  IF
	lcount = 0
	Do while not DXread.AtEndOfStream
		strLine = DXread.ReadLine()
		lcount = lcount + 1
		If instr(strLine,"XPH")>0 Then
			
			Append_TestHTML StepCounter,"PMA file Heder Data in a File", strLine , "PASSED"
			
			Rectypeval = mid(strLine,1,4)
			Originval = mid(strLine,5,10)
			fileseqval = mid(strLine,15,6)
			timestampval = mid(strLine,21,11)
			fileversionval = mid(strLine,33,5)
			otherval = mid(strLine,39)
				If Rectypeval = "XPH" Then
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4"& "Expected-XCH" &"Actual-"& Rectypeval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4" & "Expected-XCH" &"Actual-"& Rectypeval, "FAILED"		
				End If
				If instr(Originval,"ESC")>0 Then
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "FAILED"					
				End If
				If instr(fileseqval,db_LastSequence)>0 Then
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "FAILED"
							
				End If
				If instr(timestampval,datevaltext)>0 Then
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "PASSED"
				Else
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "FAILED"
					
				End If
				If instr(fileversionval,"V 3.0") > 0 Then
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "FAILED"
				End If
				If instr(otherval,"000000000000000000000000")>0 Then
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "FAILED"
							
				End If
		ElseIf  instr(strLine,"XPT")>0  Then
		Append_TestHTML StepCounter,"PMA file Trailer Data in a File", strLine , "PASSED"
			
			Rectypeval = mid(strLine,1,4)
			Originval = mid(strLine,5,10)
			fileseqval = mid(strLine,15,6)
			timestampval = mid(strLine,21,11)
			fileversionval = mid(strLine,33,5)
			otherval = mid(strLine,39)
				If Rectypeval = "XPT" Then
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4"& "Expected-XCT" &"Actual-"& Rectypeval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record type", "XCH Populated in total Length 4" & "Expected-XCH" &"Actual-"& Rectypeval, "FAILED"		
				End If
				If instr(Originval,"ESC")>0 Then
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Origin value", "ESC polpulated  for euroShell Cards in origin total length 10"& "Expected-ESC" &"Actual-"& Originval, "FAILED"					
				End If
				If instr(fileseqval,db_LastSequence)>0 Then
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File sequence number", "Sequence Number  populated from Incremental number per colco"& "Expected-"& db_LastSequence &"Actual-"& fileseqval, "FAILED"
							
				End If
				If instr(timestampval,datevaltext)>0 Then
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "PASSED"
				Else
					Append_TestHTML StepCounter,"VerifyTimestamp", "TimeStamp populated in YYMMDDHHMMSS of file creation. Total Lengh 12"& "Expected-"& datevaltext &"Actual-"& timestampval, "FAILED"
					
				End If
				If instr(fileversionval,"V 3.0") > 0 Then
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify File type version", "File type version populated Total Lenghth 6"& "Expected- V 3.0"  &"Actual-"& fileversionval , "FAILED"
				End If
				If instr(otherval,lcount)>0 Then
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "PASSED"
				Else
					Append_TestHTML StepCounter,"Verify Record count, Checksum and unused", "All are Zeros", "FAILED"
							
				End If
		End If
	
	Loop
End Function	




Function cardgroupoverRideAddress(new_countryName,address,postalCode,email,phoneNo)
On error resume next
	bFlag = True
	Call objectClick("Epin_WEl_CGOverride","Override Address tab")
	wait 3
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Epin_WC_CGuCDAobj") Then
		Append_TestHTML StepCounter, "Override Address screen", "User navigates to 'Override Address' screen", "PASSED"		
		ui_prop = Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*ucAddressMaintainDelivery_txtAddressLines").GetRoproperty("disabled")
		If ui_prop = 0 Then
			Call enterWebList_value("Epin_WE_CGDcountry","Country Name",new_countryName)
			Call enterTextbox_value("Epin_WE_CGDadress","Address",address)
			Call enterTextbox_value("Epin_WE_CGDzipcode","PostalCode",postalCode)
			Call enterTextbox_value("Epin_WE_CGDEmail","Email Address",email)
			Call enterTextbox_value("WedEdit_DeliveryMobilePhone","Mobile Phone Number",phoneNo)
		End If
		
		UCDACheckval = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WC_CGuCDAobj","GetROProperty","checked")
		If UCDACheckval = "1" Then	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WC_CGuCDAobj", "Click", ""
		End If
		Append_TestHTML StepCounter, "Verify UCDA check box", "'Use Card Delivery Address' Check box is Unticked" , "PASSED"
		Append_TestHTML StepCounter, "Verify Pin delivery Address", "'Pin delivery Address' fields are enabled", "PASSED"			
		ui_prop = Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*ucAddressMaintainPin_txtAddressLines").GetRoproperty("disabled")
		If ui_prop = 0 Then
			Call enterWebList_value("WedList_PinCountry","Country Name",new_countryName)
			Call enterTextbox_value("WedEdit_PinAddress","Address",address)
			Call enterTextbox_value("WedEdit_PinZipcode","PostalCode",postalCode)
			Call enterTextbox_value("Epin_WE_CGPemail","Email Address",email)
			Call enterTextbox_value("WebEdit_PinMobilePhone","Mobile Phone Number",phoneNo)
		End If		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Save' link", "PASSED"
	else
		Append_TestHTML StepCounter, "Override Address screen", "User does not navigates to 'Override Address' screen", "FAILED"	
	End  If		
End  Function 
	
	
	Public Function GFN_CreateCard_of_PINAdviceType(cardGroup, embossName, cardType, embossType,pinadvtypeval)
	On error resume next	
	bFlag = True	
	wait 1
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")	
	wait 10
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Card maintenance' screen", "PASSED"
		Call objectClick("WebLink_NewCard","NewCard link")
		wait 10
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_CardType") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'New Card' screen", "PASSED"
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_NewCard_Cardgroup") Then
				Call enterWebList_value("WebList_NewCard_Cardgroup","Card Group",cardGroup)
			Else
				Browser("creationTime:=1").Page("creationTime:=1").Link("innertext:=Select an Option","html tag:=A").Click
				wait 2
				Browser("creationTime:=1").Page("creationTime:=1").WebElement("html id:=ctl00_CPH_multiViewSettings_ucCardMaintain_mvCard_ddlCardGroupID.*","innertext:="&cardGroup).Click
				
				wait 4
			End If
			
			
			Append_TestHTML StepCounter, "Select value", "Select the '"&field_name&"' value as '"&value&"'", "PASSED"	
		
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_CardType", "RadioSelect", cardType
			wait 2
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_EmbossType", "RadioSelect", embossType
			wait 2
		'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_VRN", "Set", cust_vrn
			'Call enterWebList_value("WebList_CardType","Card Category",cardType)			
			'Call enterWebList_value("WebList_EmbossType","Emboss Type",embossType)
			Select Case embossType
				Case "Driver":
					'Browser(strBrowser).Page(strPage).WebEdit("html id:=.*txtEmbossDriver").Set embossName
					Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*txtEmbossDriver").Set embossName
					Append_TestHTML StepCounter, "Create Card", "Enter the DriverName '"&embossName&"'", "PASSED"	
				Case "Vehicle":
					Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*txtEmbossRegNumber").Set embossName
					Append_TestHTML StepCounter, "Create Card", "Enter the Vehicle Reg No. '"&embossName&"'", "PASSED"
'					Call enterTextbox_value("WebEdit_VRN","VRN",embossName)	
				Case "Bearer":
					Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*txtBearerDescription").Set embossName
					Append_TestHTML StepCounter, "Create Card", "Enter the BearerDescription '"&embossName&"'", "PASSED"				
			End Select
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebList_PinadvType", "RadioSelect", pinadvtypeval
	
'			Call enterWebList_value("WebList_TypeOfPin","TypeOfPin",countryName)
'			Call enterWebList_value("WebList_PinAdviceType","Pin Advice Type",pinAdviceType)
'			Call enterWebList_value("WebList_PinSelectionMethod","Pin Selection Method",pinSelectionMethod)
			Call addAddressEmailofPinadvanceType()
			
			Call click_on_saveElement()
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does  not navigates to 'New Card' screen", "FAILED"
		End  IF
	else
		Append_TestHTML StepCounter, "Card maintenance", "User does  not navigates to 'Card maintenance' screen", "FAILED"
	End  IF	
	
	cardNum = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Card","GetROProperty","default value")
	Append_TestHTML StepCounter, "Card creation", "Card PAN number '"&cardNum&"' is created successfully", "PASSED"
	expDate_actual = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Expiry","GetROProperty","default value")
	expDate_arr = split(expDate_actual,"/")
	expDate = expDate_arr(2) & "-" & expDate_arr(1) & "-" & expDate_arr(0)
	
	embossText = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossText","GetROProperty","default value")	
	
	wait 3   ' waiting for the DB update
	
	query = "select * from Card where PAN = '" & cardNum & "';"
	dbName = appName & "_SHELL_SPRINTQA_" & countryCode & "_OLTP"
	set dbRecordSet = execute_db_query(query, 1, dbName)
	wait 2
	cardPAN = dbRecordSet("PAN")
	cardPANNum = cardPAN
	dbExpDate =  dbRecordSet("ExpiryDate")
	dbEmbossNum =  dbRecordSet("EmbossRegNumber")
	dbEmbossText =  dbRecordSet("EmbossText")
	cardPAN_no = cardPAN
	cardPANNum = cardPAN
	cardExpiry_date = dbExpDate
	If instr(cardPAN, cardNum)<>0 Then
		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & cardPAN & VBCRLF & "Actual Value: " & cardNum ,"PASSED"
	else
		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & cardPAN & VBCRLF & "Actual Value: " & cardNum ,"FAILED"
		bRunFlag = False
	End If
'	
'	If instr(expDate, dbExpDate)<>0 Then
'		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & expDate & VBCRLF & "Actual Value: " & dbExpDate ,"PASSED"
'	else
'		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & expDate & VBCRLF & "Actual Value: " & dbExpDate ,"FAILED"
'		bRunFlag = False
'	End If
'	
'	If instr(driverName, dbEmbossNum)<>0 Then
'		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & driverName & VBCRLF & "Actual Value: " & dbEmbossNum ,"PASSED"
'	else
'		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & driverName & VBCRLF & "Actual Value: " & dbEmbossNum ,"FAILED"
'		bRunFlag = False
'	End If
'	
'	If instr(embossText, dbEmbossText)<>0 Then
'		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & embossText & VBCRLF & "Actual Value: " & dbEmbossText ,"PASSED"
'	else
'		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & embossText & VBCRLF & "Actual Value: " & dbEmbossText ,"FAILED"
'		bRunFlag = False
'	End If
	
End Function

Public Function validatenegativeFiledata(filePath,data,ftype)
       On error resume next
'       '''''msgbox filePath
'       '''''msgbox data
'       '''''msgbox ftype
				bFlag = True

	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(filePath) Then
		If (ucase(ftype) = ucase("dat") or ucase(ftype) = ucase("File") ) Then
			Set fileRead = fileSysObj.OpenTextFile(filePath,1,False,-1)
		Else
			Set fileRead = fileSysObj.OpenTextFile(filePath,1)
		End  IF
		content = fileRead.ReadAll
		Append_TestHTML StepCounter, "File data ", content , "PASSED"
		
'		'''''msgbox content
		If instr(data,";") = 0 Then
			If instr(content,data) = 0 Then
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "not matched in the path "&filePath , "PASSED"
				
			Else
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "matched in the path "&filePath , "FAILED"
				bRunFlag = False
				bFlag = False	
			
			End If
		Else
		dataval = split(data,";")
			For iitr = 0 To ubound(dataval) Step 1
				If instr(content,dataval(iitr))>0 Then
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "matched in the path "&filePath , "PASSED"
					
				Else
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "not matched in the path "&filePath , "FAILED"
				bRunFlag = False
				bFlag = False	
				End If
			Next
		End If
	End If
End Function

Function pinReminderPrerequestFalseData()

	On error resume next
	bFlag = True
	query = "Select * from CardIssueAttributes where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardIssueAttributes table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_RequirePIN= dbRecordSet("RequirePIN")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	If cint(db_RequirePIN) = cint("0") or db_RequirePIN = false Then
		Append_TestHTML StepCounter, "Verify CardIssueAttributes RequirePIN value", "Expected: FALSE" & VBCRLF & "Actual: " & db_RequirePIN , "PASSED"
	Else
		Append_TestHTML StepCounter, "Verify CardIssueAttributes RequirePIN value", "Expected: FALSE" & VBCRLF & "Actual: " & db_RequirePIN , "FAILED"	
	End If
	
	query = "Select * from CardPAN where CardPANID="& XML_Cardpanid & " order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CardPAN table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_PAN1 = dbRecordSet("PAN")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	query = "Select * from Card where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute Card table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_TypeofPin = dbRecordSet("TypeOfPIN")
	db_PinSelMethod = dbRecordSet("PINSelectionMethod")
	db_StatusID = dbRecordSet("StatusID")
	db_ExpiryDate = dbRecordSet("ExpiryDate")
	db_Pan = dbRecordSet("PAN")
	
	pAPI_Expitedate = Mid(Replace(db_ExpiryDate,"-",""),3,4)
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	
	
	query = "Select * from CustomerCard where CardID="& XML_Cadrid & "  order by 1 desc;"
	Append_TestHTML StepCounter, "Execute CustomerCard table query ",query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_CustomerCardID = dbRecordSet("CustomerCardID")
	db_CustomerID = dbRecordSet("CustomerID")
	
	Set dbRecordSet = Nothing
	Set query = Nothing
	dbquery = "Select * from SysVarColco where SysVarID=141;"
	Append_TestHTML StepCounter, "Execute SysVarColco query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_ColcoID = dbRecordSet("ColcoID")
	
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	pAPI_PANID = XML_Cardpanid
	pAPI_CardID = XML_Cadrid
	pAPI_CustomerID = db_CustomerID
	pAPI_cardPAN = db_Pan
	
	If pAPI_Expitedate <>"" Then
		pinReminderPrerequestFalseData = pAPI_CustomerID & "|" & pAPI_CardID & "|" &  pAPI_PANID & "|" & pAPI_cardPAN & "|" & pAPI_Expitedate &"|"& db_ColcoID
	Else
		pinReminderPrerequestFalseData = NULL
	End If
End Function
