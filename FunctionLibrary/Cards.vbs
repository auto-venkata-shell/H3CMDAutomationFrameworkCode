Public cards_embossingName



'******************************* HEADER ******************************************
' Description : The function to open CardMaintenance screen of specific customer
' Creator :  Venkata Srinivasa Rao. K
' Date : 3rd June, 2022
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Function fillCardTypeDetails(cardT_data)

	On error resume next
	
	bFlag = True
	wait 1
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Types"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardTypes", "Click", ""
	
	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCardType") Then
	
		textsearchval = ValidateWebTableCellValue("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "No_Frame","Cards_WebTable_CTListdetails",cardT_data("CardType"))
	If textsearchval Then
		Append_TestHTML StepCounter, "Verify Card Type Details", "Card type entry already exist under customer", "PASSED"
	
	Else
	
		Append_TestHTML StepCounter, "Card Types", "New Card Type screen navigatiion", "PASSED"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCardType", "Click", ""
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebCheckBox_CardReissue") Then
			Append_TestHTML StepCounter, "Create New Card", "Navigated to New card screen", "PASSED"
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_CardReissue", "Click", ""
			wait 2
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CardType", "RadioSelect", cardT_data("CardType")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_TokenType", "RadioSelect", cardT_data("TokenType")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_EmbossingType", "RadioSelect", cardT_data("EmbossingType")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_PurchaseCat", "RadioSelect", cardT_data("PurchaseCategory")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_MinCardLife", "Set", cardT_data("MinCardLifeval")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_MinReissuePeriod", "Set", cardT_data("MinReissuePeriodval")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ExpiryPeriod", "Set", cardT_data("ExpiryPeriodval")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossingName", "Set", cardT_data("EmbossingNameval")
			minCardLife = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_MinCardLife","GetROProperty","value")
			
			minReissuePeriod = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_MinReissuePeriod","GetROProperty","value")
			
			expiryPeriod = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ExpiryPeriod","GetROProperty","value")
			
			cards_embossingName = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossingName","GetROProperty","value")
			
			isActive = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_CTActive","GetROProperty","checked")
			If isActive = 1 or isActive = True Then
				isActive = True
			else
				isActive = False
			End If
			
			isReissue = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_Reissue","GetROProperty","checked")
			If isReissue = 1 or isReissue = True Then
				isReissue = True
			else
				isReissue = False
			End If
				Append_TestHTML StepCounter,"Customer card type details ","Maintained all input values" ,"PASSED"
			
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			
			
		else
			Append_TestHTML StepCounter, "Create New Card", "Navigation to create card page failed", "FAILED"
			bFlag = False
		End If
	
	End If
	
	else
		Append_TestHTML StepCounter, "Card Types", "Navigation to card types page failed", "FAILED"
		bFlag = False
	End If
	
	
	
	
	wait 3   ' waiting for the DB update
	
	query = "select * from CustomerCardType where CustomerID = (select customerid from Customer where CustomerERP = '" & customerERP_id & "') order by 1 desc;"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	If instr(dbRecordSet("MinCardLife"),minCardLife) <> 0 Then
		Append_TestHTML StepCounter,"MinCardLife value matches","Expected Value: " & minCardLife & VBCRLF & "Actual Value: " & dbRecordSet("MinCardLife") ,"PASSED"
	else
		Append_TestHTML StepCounter,"MinCardLife value mismatch","Expected Value: " & minCardLife & VBCRLF & "Actual Value: " & dbRecordSet("MinCardLife") ,"FAILED"
		bFlag = False
	End If
	
	If instr(dbRecordSet("MinReissueCardPeriod"),minReissuePeriod) <> 0 Then
		Append_TestHTML StepCounter,"MinReissueCardPeriod value matches","Expected Value: " & minReissuePeriod & VBCRLF & "Actual Value: " & dbRecordSet("MinReissueCardPeriod") ,"PASSED"
	else
		Append_TestHTML StepCounter,"MinReissueCardPeriod value mismatch","Expected Value: " & minReissuePeriod & VBCRLF & "Actual Value: " & dbRecordSet("MinReissueCardPeriod") ,"FAILED"
		bFlag = False
	End If
	
	If instr(dbRecordSet("ExpiryPeriod"),expiryPeriod) <> 0 Then
		Append_TestHTML StepCounter,"ExpiryPeriod value matches","Expected Value: " & expiryPeriod & VBCRLF & "Actual Value: " & dbRecordSet("ExpiryPeriod") ,"PASSED"
	else
		Append_TestHTML StepCounter,"ExpiryPeriod value mismatch","Expected Value: " & expiryPeriod & VBCRLF & "Actual Value: " & dbRecordSet("ExpiryPeriod") ,"FAILED"
		bFlag = False
	End If
	
	If instr(dbRecordSet("EmbossName"),embossingName) <> 0 Then
		Append_TestHTML StepCounter,"EmbossName value matches","Expected Value: " & embossingName & VBCRLF & "Actual Value: " & dbRecordSet("EmbossName") ,"PASSED"
	else
		Append_TestHTML StepCounter,"EmbossName value mismatch","Expected Value: " & embossingName & VBCRLF & "Actual Value: " & dbRecordSet("EmbossName") ,"FAILED"
		bFlag = False
	End If
	''msgbox dbRecordSet("IsActive") 
	''msgbox isActive
	If dbRecordSet("IsActive") = isActive Then
		Append_TestHTML StepCounter,"IsActive value matches","Expected Value: " & isActive & VBCRLF & "Actual Value: " & dbRecordSet("IsActive") ,"PASSED"
	else
		Append_TestHTML StepCounter,"IsActive value mismatch","Expected Value: " & isActive & VBCRLF & "Actual Value: " & dbRecordSet("IsActive") ,"FAILED"
		bFlag = False
	End If
	
	If dbRecordSet("IsFixedReissue") = isReissue Then
		Append_TestHTML StepCounter,"IsFixedReissue value matches","Expected Value: " & isReissue & VBCRLF & "Actual Value: " & dbRecordSet("IsFixedReissue") ,"PASSED"
	else
		Append_TestHTML StepCounter,"IsFixedReissue value mismatch","Expected Value: " & isReissue & VBCRLF & "Actual Value: " & dbRecordSet("IsFixedReissue") ,"FAILED"
		bFlag = False
	End If
	
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for create customer card types","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for create customer card types","FAILED"
		bRunFlag = False
	End If
	
	set dbRecordSet = Nothing
End Function


'******************************* HEADER ******************************************
' Description : The function to create card in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function CreateandSaveCard( cardT_data)

	'On error resume next
	
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Maintenance"
	wait 5
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardMaintenance", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter, "Create card", "Navigated to create card page", "PASSED"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCard", "Click", ""
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_CardType") Then
			Append_TestHTML StepCounter, "Create card", "Navigated to create new card page", "PASSED"
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_CardTypeIDVal", "RadioSelect", cardT_data("CardType")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CMTokenType", "RadioSelect", cardT_data("CardType")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CMPurchaseCat", "RadioSelect", cardT_data("PurchaseCategory")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_CardType", "RadioSelect", cardT_data("Cardcategory")
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_EmbossType", "RadioSelect", cardT_data("EmbossingType")
			If instr(cardT_data("Cardcategory"),"Vehicle")>0 Then
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_VRN", "Set", cardT_data("CustVRNval")
			ElseIf instr(cardT_data("Cardcategory"),"Driver")>0 Then
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebEdit_DriverName", "Set", cardT_data("CustVRNval")
			
			End If
			
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
				Append_TestHTML StepCounter, "Create card", "Added details on create new card page", "PASSED"
					
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
				Call insertDataintoCardOverrideScreen()
				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WebElement_Failedheader") Then
					Append_TestHTML StepCounter, "Create card", "Getting Error as Failed Validation", "FAILED"
				Else
						If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") Then
							Append_TestHTML StepCounter, "Create card", "Navigated to create new card page", "PASSED"
							OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
							wait 5
							cardNum = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Card","GetROProperty","default value")
							
							expDate_actual = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Expiry","GetROProperty","default value")
							expDate_arr = split(expDate_actual,"/")
							expDate = expDate_arr(2) & "-" & expDate_arr(1) & "-" & expDate_arr(0)
							
							embossText = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossText","GetROProperty","default value")	
							
							wait 3   ' waiting for the DB update
							
							query = "select * from Card where PAN = '" & cardNum & "';"
							set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
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
							
							If instr(expDate, dbExpDate)<>0 Then
								Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & expDate & VBCRLF & "Actual Value: " & dbExpDate ,"PASSED"
							else
								Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & expDate & VBCRLF & "Actual Value: " & dbExpDate ,"FAILED"
								bRunFlag = False
							End If
							
							If instr(cardT_data("CustVRNval"), dbEmbossNum)<>0 Then
								Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & cardT_data("CustVRNval") & VBCRLF & "Actual Value: " & dbEmbossNum ,"PASSED"
							else
								Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & cardT_data("CustVRNval") & VBCRLF & "Actual Value: " & dbEmbossNum ,"FAILED"
								bRunFlag = False
							End If
							
							If instr(cards_embossingName, dbEmbossText)<>0 Then
								Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & cards_embossingName & VBCRLF & "Actual Value: " & dbEmbossText ,"PASSED"
							else
								Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & cards_embossingName & VBCRLF & "Actual Value: " & dbEmbossText ,"FAILED"
								bRunFlag = False
							End If
				
						Else
							Append_TestHTML StepCounter, "Create card", "UI error and validation data missed", "FAILED"
						End If
						
				End  If	
					
			Else
				Append_TestHTML StepCounter, "Create card", "Adding details on create card page failed", "FAILED"
			End If
			
		
		else
			Append_TestHTML StepCounter, "Create card", "Navigation to create card page failed", "FAILED"
		End If
		
		
	else
		Append_TestHTML StepCounter, "Create card", "Navigation to create card page failed", "FAILED"
	End If
	
	
End Function

Function navigateCustomerSummaryscreenforOtherScreens()
	On error resume next
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
	wait 1
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", customerERP_id
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_CustSummary") = True  Then
			Append_TestHTML StepCounter, "Navigate to Customer Summary Screen", "Successfully navigated", "PASSED"
	Else
			Append_TestHTML StepCounter, "Navigate to Customer Summary Screen", "Successfully navigated", "FAILED"
	End If
	
End Function


'******************************* HEADER ******************************************
' Description : The function to change the value for specific Key data in SysVarColco value
' Creator : Venkata Srinivasa Rao. K
' Date : 26th October, 2022
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function doSysVarColcoKeydataFlagAction(keydata,cstatusFlag,chstatusFlag)

On Error Resume Next
	bRunFlag = True	
	query = "Select * from SysVarColco where [Key]='" & keydata & "';"
		Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		db_Key_val = dbRecordSet("Key")
		db_sysvar_Id = dbRecordSet("SysVarID")
		db_def_Vaue = dbRecordSet("Value")
		Set dbRecordSet = Nothing
		
		If isNull(db_def_Vaue) Then
			Append_TestHTML StepCounter, keydata & "Default value","Unknown/NULL data value "& VBCRLF & "Actual - Key: " &db_Key_val & VBCRLF & "Actual - SysVarID: " &db_sysvar_Id & VBCRLF & "Actual - Value: " &db_def_Vaue  ,"FAILED"
			bRunFlag = False
		Else
			If cBool(db_def_Vaue) = cBool(cstatusFlag) Then
				Append_TestHTML StepCounter, keydata & " Existing value","Expected Value: " & cstatusFlag  & VBCRLF & "Actual Value: " &db_def_Vaue  ,"PASSED"
				query = "Update SysVarColco set Value= " & chstatusFlag & " where SysVarID=" & db_sysvar_Id & ";"
				Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
				Set dbRecordSet = Nothing
	
				query = "Select * from SysVarColco where [Key]='" & keydata & "';"
				Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
				db_Key_val1 = dbRecordSet("Key")
				db_def_Vaue1 = dbRecordSet("Value")
				Set dbRecordSet = Nothing
				If cBool(db_def_Vaue1) = cBool(chstatusFlag) Then
					Append_TestHTML StepCounter,keydata & " Updated value","Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_def_Vaue1  ,"PASSED"
				Else
					Append_TestHTML StepCounter, db_Key_val & "Data Updated value","Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_def_Vaue1  ,"FAILED"
				End If
			Else
				Append_TestHTML StepCounter, keydata & " Existing value","No Update requiered in DB - Expected Value: " & chstatusFlag & VBCRLF & "Actual Value: " &db_def_Vaue  ,"PASSED"
			End If
		End If
End  Function

Function cardsPreConditionDBconfigcheck(keydata,valcheck)
	
	On Error Resume Next
			query = "Select * from SysVarColco where [Key]='" & keydata & "';"
			Append_TestHTML StepCounter, "Execute Query", query , "PASSED"
			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			db_Key_val1 = dbRecordSet("Key")
			db_def_Vaue1 = dbRecordSet("Value")
			Set dbRecordSet = Nothing
			Set query = Nothing
			If db_def_Vaue1 = valcheck Then
					Append_TestHTML StepCounter,keydata & " Verify " & keydata & " value","Expected Value: " & valcheck & VBCRLF & "Actual Value: " &db_def_Vaue1  ,"PASSED"
				Else
					Append_TestHTML StepCounter,keydata & " Verify " & keydata & " value","Expected Value: " & valcheck & VBCRLF & "Actual Value: " &db_def_Vaue1  ,"FAILED"			
			End If
			
End Function

Function navigateCardMaintanceScreen()
	
	On error resume next
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Maintenance"
	wait 5
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardMaintenance", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
			Append_TestHTML StepCounter,"CardMaintenance Screen navigation", " Successfully navigated"  ,"PASSED"
		Else
			Append_TestHTML StepCounter,"CardMaintenance Screen navigation", " Fail to navigate"  ,"FAILED"
			bFlag = False
			bRunFlag = False
	End  If
	
End Function




Function searchCardsinCardMaitenance(statusType)

On Error Resume Next
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Card_WebElement_Searchtab") Then
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebElement_Searchtab", "Click", ""
	wait 5
End If
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WE_CM_card", "Set", cardPANNum
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CM_status", "RadioSelect", statusType
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebElement_Searchtab", "Click", ""
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	wait 2
	cardexist = ValidateWebTableCellValue("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "No_Frame","Cards_WebTable_CM_searchresults",cardPANNum)
	wait 2
	If cardexist Then
		Call ClickObjectInWebTbl("WebBrowser_PTShell_INDO","NoWindow", "WebPage_PTShell_INDO","NoFrame","Cards_WebTable_CM_searchresults","Cards_WC_rowcheck","html tag:INPUT","")
		wait 2
				Append_TestHTML StepCounter,"Verify card "& cardPANNum & "in the Search " & statusType & " Status Result", "Data Verified and selected successfully"  ,"PASSED"
	Else
				Append_TestHTML StepCounter,"Verify card "& cardPANNum & "in the Search " & statusType & " Status Result", "Data validation Failed"  ,"FAILED"
	
	End If
End  If
End Function


'type:checkbox

Function insertDataintoCardOverrideScreen()
On Error Resume Next
If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=hidden error stick","html tag:=LI","innertext:=.*There are errors on your page.*").Exist Then
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_overridetab", "Click", ""
	
	If countryCode = "PH" Then
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_OR_WE_mobile") Then
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_OR_WE_mobile", "Set", "+919876543211"
				Append_TestHTML StepCounter,"Enter Phone number in Card Override screen", "Data inserted successfully"  ,"PASSED"
			Else
				Append_TestHTML StepCounter,"Enter Phone number in Card Override screen", "Fail to navigate or field validation failed"  ,"FAILED"
			
		End If
	
	ElseIf countryCode = "NL" Then	
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Epin_WE_APDEmail") Then
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_APDEmail", "Set", "testnewemail@shell.com"
				Append_TestHTML StepCounter,"Enter Pin delivery email in Card Override screen", "Data inserted successfully"  ,"PASSED"
			Else
				Append_TestHTML StepCounter,"Enter Pin delivery email in Card Override screen", "Fail to navigate or field validation failed"  ,"FAILED"
			
		End If
		
	ElseIf countryCode = "ID" Then
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Epin_WE_ACDEmail") Then
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_ACDEmail", "Set", "testnewemail@shell.com"
				Append_TestHTML StepCounter,"Enter Card delivery email in Card Override screen", "Data inserted successfully"  ,"PASSED"
			Else
				Append_TestHTML StepCounter,"Enter Card delivery email in Card Override screen", "Fail to navigate or field validation failed"  ,"FAILED"
			
		End If
		
	End If
OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
wait 2
End If

End Function

Function performCardActions(actionName)

On Error Resume Next
wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Drop")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Drop", "Mover", ""
'				Append_TestHTML StepCounter, "Verify options under selected card Actions","Successfully clicked on Action and Verified", "PASSED"
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Drop", "Click", ""
		wait 5
		Append_TestHTML StepCounter, "Verify options under selected card Actions","Successfully clicked on Action and Verified", "PASSED"
		
	End If
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", actionName, "Click", ""
End Function

Function copyCardsValidation()
On Error Resume Next
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WE_CM_numberofcards")  Then
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WE_CM_numberofcards", "Set", "1"
	Append_TestHTML StepCounter, "Enter Value into the Number of Cards Field","Successfully entered value 1", "PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebElement_numCardsSave", "Click", ""
	wait 4
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WebTable_copyRestricteddata") Then
				Append_TestHTML StepCounter, "Verify Restricted Data screen","Successfully verified", "PASSED"
	
				Call ClickObjectInWebTbl("WebBrowser_PTShell_INDO","NoWindow", "WebPage_PTShell_INDO","NoFrame","Cards_WebTable_copyRestricteddata","Cards_WebElement_ErrRegistration","outerhtml:",".*EmbossRegNumber.*")
				wait 2
				Call OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebEdit",3,1,"Set","WEB1002")
				Call OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebEdit",3,2,"Set","NewCDriver")
				Call OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebEdit",3,3,"Set","NewBDesc")
				Call OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebEdit",3,4,"Set","NewAddText")
				wait 2
				Append_TestHTML StepCounter, "Maintain Mandatory Fields Data","Successfully data maintained", "PASSED"
				wait 4
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
				wait 2
				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_Card") Then
					Append_TestHTML StepCounter, "Verify screen navigated back to Card Maintan tabl","Successfully naviagated", "PASSED"
				

				Else
					Append_TestHTML StepCounter, "Verify screen navigated back to Card Maintan tabl","Fail to naviagate", "FAILED"

				End If
				
	End If
Else
	Append_TestHTML StepCounter, "Verify Confimartion Pop-up","Fail to appear", "FAILED"

End If
End Function




Function cardCancellAction()
On error resume next
	
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Cards_WebElement_label") Then
		sel_value = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CM_Newstatus", "GetROProperty", "selection")
			Append_TestHTML StepCounter, "Navigate to Change Status Screen","Successfully navigated", "PASSED"
		If ucase(sel_value) = ucase("Cancelled") Then
			textsearchval = ValidateWebTableCellValue("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "No_Frame","Cards_WebTable_CM_searchresultstable",cardPANNum)
			If textsearchval Then
					Append_TestHTML StepCounter, "Verify new status drop down value","Selected " & sel_value & "as default value", "PASSED"
			
				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
					If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") Then
						Append_TestHTML StepCounter, "Click on Save button to cancel the card","Clicked and getting Confirmation pop-up", "PASSED"
						OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
							wait 5
					Else
					Append_TestHTML StepCounter, "Verify Confirmation Save Pop-up","No confirmation Save pop-up ", "FAILED"
					
					End If
				Else
					Append_TestHTML StepCounter, "Click on Save button to cancel the card","No Save Link button", "FAILED"
				End If
			Else
					Append_TestHTML StepCounter, "Verify new status drop down value","Not selected " & sel_value & "as default value", "FAILED"
			
			End If
		End If
	Else
			Append_TestHTML StepCounter, "Navigate to Change Status Screen","Fail to  navigate", "FAILED"
		
	End If
End Function

Function navigateCMscreenviaSearchforCards(CardNo, statusT )
	On Error Resume Next
	nav_status =  navigateStartMenu("Link_Cards","Cards_Link_SearchForCards","WebEdit_CardNo")
	If nav_status Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CardNo", "Set", CardNo
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search",  "Click", ""
		
			textsearchval = ValidateWebTableCellValue("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "No_Frame","WebTable_SearchForCards",CardNo)
			textsearchTval = ValidateWebTableCellValue("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "No_Frame","WebTable_SearchForCards",statusT)
			
			If textsearchval and textsearchTval Then
				Append_TestHTML StepCounter, "Verify Card search results","Successfully Identified with " & statusT  &  " status", "PASSED"
				
'				Cards_WebTable_CM_searchresults
				
				Call ClickObjectInWebTbl("WebBrowser_PTShell_INDO","NoWindow", "WebPage_PTShell_INDO","NoFrame","WebTable_SearchForCards","Cards_WebElement_searchText","innertext:",statusT)
				wait 4
				If  VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Card_WebElement_Searchtab")  Then
					Append_TestHTML StepCounter, "Navigate to Card Maintance screen by clicking search entry","Successfully navigated", "PASSED"
				Else
				Append_TestHTML StepCounter, "Navigate to Card Maintance screen by clicking search entry","Fail to navigate", "FAILED"
				
				End If
				
			Else
				Append_TestHTML StepCounter, "Verify Card search results","Fail to Identify with " & statusT  &  " status", "FAILED"
			
			End If
	Else
		Append_TestHTML StepCounter, "Navigate to Cards - Search for Cards","Fail to navigate the Search for Cards screen", "FAILED"
	End If
	
End Function
	

Function navigateToCMSearchTab(searchobj,searchdata,verifydata)
	
On Error Resume Next
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Card_WebElement_Searchtab") Then
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebElement_Searchtab", "Click", ""
	wait 2
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter,"Verify to navigate Search Tab", "Successfully navigated'", "PASSED"
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Refresh", "Click", ""
		defaultstatus_val = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CM_status", "GetROProperty", "selected item index")
				
		If cint(defaultstatus_val) = cint("0") Then
			Append_TestHTML StepCounter, "Clear Search Filters", "Cleared all default search values", "PASSED"
		' Do dynamic below data as function arguments		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", searchobj, "Set", searchdata
			
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
			
			
			cardissuevalexist = ValidateWebTableCellValue("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "No_Frame","Cards_WebTable_CM_searchresults",verifydata)
			 cardissuevalexist1 = OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebElement",2,11,"GetCellData","")
			 newcardPANval = OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebElement",2,3,"GetCellData","")
			 newcardStatus = OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebElement",2,4,"GetCellData","")
			 newcarddetails = OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebElement",2,7,"GetCellData","")
			
			 newcarddetails1 = OperateOnWebtable("WebBrowser_PTShell_INDO","No_Window", "WebPage_PTShell_INDO","No_Frame","Cards_WebTable_copyRestricteddata","WebElement",2,7,"GetRowWithCellText","NewCDriver")
	
			If cardissuevalexist or instr(newcarddetails,"NewCDriver")>0 Then
					Append_TestHTML StepCounter, "Verify Re issued card Details ", "Details are - " & newcardPANval & "-" & newcardStatus & "-" & cardissuevalexist  , "PASSED"
			Else
					Append_TestHTML StepCounter, "Verify Re issued card Details ", "Details are - " & newcardPANval & "-" & newcardStatus & "-" & cardissuevalexist  , "FAILED"
			
			End If
		Else 
			Append_TestHTML StepCounter, "Clear Search Filters", "Not Cleared all default search values", "FAILED"
		
		End If
	Else
		Append_TestHTML StepCounter,"Verify to navigate Search Tab", "Fail to navigate'", "FAILED"
	
	End If
Else
		Append_TestHTML StepCounter,"Verify to Card Maintance Tab", "Fail to navigate Maintanance Tab'", "FAILED"

End If
End Function





		
	
	

'
'
'
'Function navigate_CardStatus_in_GFN(URL,cardNo)
'Set cardNoobj =  Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebEdit("html id:=ctl00_CPH_csdCardSearch_mvMain_txtPAN")
'Set cardstausobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_csdCardSearch_mvMain_ddlStatusID")
'Set cardresultobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebTable("html id:=ctl00_CPH_csdCardSearch_grdResults")
'Set blockcardobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").Link("class:=level2","html tag:=A","innertext:=Block / Unblock Card ")
'Set cardMaintananceobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").Link("class:=popout level1","html tag:=A","innertext:=Card Maintenance ")
'	On Error resume Next
'bSingleSignOnFlag = True 
'bSignOnFlag = False
'	Call launchApplication()
'	Call pageNavigation("Search for Cards","Link_SearchForCards")
'	If cardNoobj.exist Then
'		Append_TestHTML StepCounter, "Search for Cards","User navigates successfully to 'Search For Cards' page", "PASSED"
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CardNo", "Set", cardNo
'		cardstausobj.Select "Active"
'		wait 2
'		Append_TestHTML StepCounter, "Search for Cards with active status","Enter the Card number(PAN) '"&cardNo&"'", "PASSED"
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'		Append_TestHTML StepCounter, "Search for Cards","Click on Search link", "PASSED"
'		If cardresultobj.Exist Then
'			Call clickOntableElement(cardresultobj)
'					
'		Else
'					Append_TestHTML StepCounter, "Verify Card Search results","No results are displayed", "FAILED"
'							bRunFlag = False			
'					
'		End If
'	else
'		Append_TestHTML StepCounter, "Verify Search for Cards screen","Fail to navigate screen", "FAILED"
'							bRunFlag = False			
'		
'	End  IF
'Set cardNoobj =  Nothing
'Set cardstausobj = Nothing
'Set cardresultobj = Nothing
'Set blockcardobj = Nothing
'Set cardMaintananceobj = Nothing
'		
'End Function
'
'Function blockCardAction()
'	Set stauschobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_ddlStatus")
'Set reasonvalobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_ddlReason")
'Set notestxtobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebEdit("html id:=ctl00_CPH_txtNotes","html tag:=TEXTAREA")
'	On Error resume Next
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_Cardmaintenance")  Then
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Cardmaintenance", "Mover", ""
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_block", "Click", ""
'			If stauschobj.Exist Then
'				Append_TestHTML StepCounter, "Verify Block/Unblock Card Screen","Navigated successfully", "PASSED"
'				stauschobj.Select "Blocked Card"
'				reasonvalobj.Select "Lost"
'				notestxtobj.Set "block test"
'				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'				Append_TestHTML StepCounter, "Change Card status","Successfully changed to Blocked status", "PASSED"
'			Else
'				Append_TestHTML StepCounter, "Verify Block/Unblock Card Screen","Fail to Navigate", "FAILED"
'							bRunFlag = False			
'			End If
'	Else
'				Append_TestHTML StepCounter, "Verify CardMaintanance Screen","Fail to Navigate", "FAILED"
'							bRunFlag = False		
'	End If
'		Set stauschobj = Nothing
'Set reasonvalobj = Nothing
'Set notestxtobj = Nothing
'			
'End Function
'
''******************************* HEADER ******************************************
'' Description : The function to open CardMaintenance screen of specific customer
'' Creator :  Venkata Srinivasa Rao. K
'' Date : 3rd June, 2022
'' Last Modified On : 
'' Last Modified By :
'' Input Parameter : 
'' Output Parameter : None
''******************************************** HEADER ******************************************
''Function navigateExistingcustomerSummaryScreen(custERPid)
''
''	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_Start") Then
''		Call  navigateStartMenu("Link_Customers","Link_SearchforCustomer","WebLIST_Role")
''		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLIST_Role") Then
''			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
''			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", custERPid
''			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'''			Call navigateCustomerSummaryMenuoption("Card Maintenance","WebLink_NewCard")
''			
''
''			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary") Then
''				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Successfully navigated to requested customer screen", "PASSED"
'''				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Maintenance"
'''				
'''				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardMaintenance", "Click", ""
''				Set cussumobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").Link("innertext:=Customer Summary ","class:=popout level1","visible:=True")
''				Set cardmaichildobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").Link("innertext:=Card Maintenance ","class:=level2")
''				cussumobj.Highlight
''				Setting.webPackage("ReplayType")=2
''						wait 2
''						cussumobj.FireEvent "onmouseover"
''						wait 2
''						Setting.webPackage("ReplayType")=1
''						wait 2
''					cussumobj.DoubleClick
''					wait 5
''					cardmaichildobj.Highlight
''				 cardmaichildobj.Click
''				 wait 5
''				 OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCard", "Mover", ""
''			Else
''				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Fail to navigate customer screen", "FAILED"
''							bRunFlag = False
''				
''			End If
''		Else
''				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Fail to navigate customer search screen", "FAILED"
''							bRunFlag = False
''
''		End If
''	Else
''				Append_TestHTML StepCounter, "GFN Application Home Page", "Fail to get HOME Screen", "FAILED"
''							bRunFlag = False
''	End If
''
''End Function
''
''
'
''******************************* HEADER ******************************************
'' Description : The function to open any option under Start Menu 
'' Creator :  Venkata Srinivasa Rao. K
'' Date : 3rd June, 2022
'' Last Modified On : 
'' Last Modified By :
'' Input Parameter : 
'' Output Parameter : None
''******************************************** HEADER ******************************************
'Public Function navigateStartMenu(mainoptionobj,suboptionobj,verifyoptionobj)
'	On error resume next
'	bFlag = True
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Start", "Mover", ""
'	If suboptionobj = "" Then
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", mainoptionobj, "Click", ""
'	Else
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "",mainoptionobj, "Mover", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", suboptionobj, "Click", ""
'	End If
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", verifyoptionobj)  Then
'		Print "Navigated to"&verifyoptionobj
'	Else
'		bFlag = False
'	End IF		
'	navigateStartMenu = bFlag
'End Function
'
'Function moveCardsAction()
'	Set cardmaresultobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebTable("html id:=ctl00_CPH_multiViewSettings_mvActions_ucCardSearch_grdResults")
'	Set customerToobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_ddlCustomerTo")
'	Set cardgroupobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_ddlCardGroupTo")
'	Set effDateobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebEdit("html id:=ctl00_CPH_sdEffectiveDate_Start")
'On Error resume Next
'If cardmaresultobj.Exist Then
'	cardmaresultobj.ChildItem(2,1,"WebCheckBox",0).set "On"
'	wait 5
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Actions", "Mover", ""
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_MMCards", "Click", ""
'			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_MMTargetlabel")  Then
'				Append_TestHTML StepCounter, "Click on MoveMultipleCards Action","Successfully navigated", "PASSED"
'				customerToobj.Select "TT"
'				cardgrpvalues = cardgroupobj.GetROProperty("all items")
'				If instr(cardgrpvalues,";") >0 Then
'					cardgroupobj.Select "#1"
'				End If
'				Append_TestHTML StepCounter, "Fill Target section customerid,cardgroup and Effective date","Successfully selected", "PASSED"
'				
'				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'			Else
'				Append_TestHTML StepCounter, "Click on MoveMultipleCards Action","Fail to  navigate", "FAILED"
'							bRunFlag = False		
'				End If
'		
'Else
'				Append_TestHTML StepCounter, "Click on CardMaintanance Screen","Fail to  navigate", "FAILED"
'				
'							bRunFlag = False		
'End If
'
'
'End Function
'
'
'Function clickOntableElement(gridCTobj)
'	
'	On Error resume Next
'		If gridCTobj.RowCount > 1 Then
'			cell_value = gridCTobj.GetCellData(2,3)
'			Set myObject = description.Create
'			myObject("micclass").value = "WebElement"
'			myObject("innertext").value = cell_value
'			Set ccobj = gridCTobj.ChildObjects(myObject)
'			ccobj(0).Highlight
'			wait 5
'			Append_TestHTML StepCounter, "Select row from the Transaction result","Successfully selected and clicking", "PASSED"
'			ccobj(0).click
'			wait 5
'		End  If
'End Function 
'
'
'Function searchsubmenuUnderCustomer(custERPid,searchitem,objtypeval,identifyobj)
'On error resume next
'	bFlag = True
'			
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_Start") Then
'		Call  navigateStartMenu("Link_Customers","Link_SearchforCustomer","WebLIST_Role")
'		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLIST_Role") Then
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", custERPid
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary") Then
'				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Successfully navigated to requested customer screen", "PASSED"
'				
'					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'					
'					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", searchitem
'					
'					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", objtypeval, "Click", ""
'					wait 4
'				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",identifyobj) Then
'						Append_TestHTML StepCounter, "Open"& searchitem &" Page", "Successfully navigated to "& searchitem & " screen", "PASSED"
'					Else
'						Append_TestHTML StepCounter, "Open"& searchitem &" Page", "Fail to  navigat "& searchitem & " screen", "FAILED"	
'								bRunFlag = False						
'				End If
'			Else
'						Append_TestHTML StepCounter, "Open Customer Summary Screen Page", "Fail to navigate Customer Summary screen", "FAILED"	
'								bRunFlag = False						
'			
'			End  If	
'		Else
'						Append_TestHTML StepCounter, "Open Customer Search Screen Page", "Fail to navigate Customer Search screen", "FAILED"	
'								bRunFlag = False				
'		End  IF	
'	Else
'						Append_TestHTML StepCounter, "Open Home Page", "Fail to navigate Home screen", "FAILED"	
'								bRunFlag = False			
'	End  IF				
'End Function
'
'Function changeCustomerStatus()
'	Set custstausddobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_mvCustomerStatus_ddlStatus")
'	Set reasonstatusddobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebList("html id:=ctl00_CPH_mvCustomerStatus_ddlReason")
'	On Error Resume Next
'	If custstausddobj.Exist Then
'		custstausddobj.Select "Blocked"
'		wait 3
'		reasonstatusddobj.Select "Credit"
'		Append_TestHTML StepCounter, "Change Customer Status to Blocked", "status selected and reason maintained", "PASSED"	
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'		Append_TestHTML StepCounter, "Verify Confirmation pop-up", "Verified and click on OK", "PASSED"	
'		wait 2
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebElement_OK", "Click", ""
'						
'	Else
'		Append_TestHTML StepCounter, "Change Customer Status to Blocked", "Fail to navigate the screen", "FAILED"	
'	
'	End If
'End Function
'
'Function getActiveCardfromCardDBTable()
'	On Error Resume Next
'	bRunFlag = True	
'		query = "Select Card.CardID,Card.PAN from CustomerCard, Card where CustomerCard.CardID = Card.CardID and Card.StatusID = 1 and Card.ExpiryDate > CURRENT_TIMESTAMP;"
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		active_card_id = dbRecordSet("CardID")
'		active_PAN = dbRecordSet("CardID")
'		Set dbRecordSet = Nothing
'		If active_card_id <> "" Then
'			getActiveCardfromCardDBTable = active_card_id & "|" & active_PAN
'		Else
'			getActiveCardfromCardDBTable = ""
'			Append_TestHTML StepCounter,"Get active card from Card Table" ,"No active cards"  ,"FAILED"
'							bRunFlag = False			
'		End If
'End Function
'
'Function getCustomeridFromCustomerTable(cERP)
'	On Error Resume Next
'	bRunFlag = True	
'		query = "Select * from Customer where CustomerERP='" & cERP & "';"
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		new_cust_id = dbRecordSet("CustomerID")
'		Set dbRecordSet = Nothing
'		If new_cust_id <> "" Then
'			getCustomeridFromCustomerTable = new_cust_id
'		Else
'			getCustomeridFromCustomerTable = ""
'			Append_TestHTML StepCounter,"Create and validate Creation of Customer table" ,"Customer entry not created int Customer table "  ,"FAILED"
'							bRunFlag = False			
'		End If
'End Function
'
'Function getCardOfCustomer(CardPanNum)
'	On Error Resume Next
'	bRunFlag = True	
'		query = "Select * from Card where PAN='" & CardPanNum & "';"
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		new_card_id = dbRecordSet("CardID")
'		Set dbRecordSet = Nothing
'		If new_cust_id <> "" Then
'			getCustomeridFromCustomerTable = new_card_id
'		Else
'			getCustomeridFromCustomerTable = ""
'			Append_TestHTML StepCounter,"Create and validate Creation of Customer table" ,"Customer entry not created int Customer table "  ,"FAILED"
'							bRunFlag = False			
'		End If
'End Function
'
'Function checkAuthoriserQueueTable(erID,eTypeid)
'	On Error Resume Next
'	bRunFlag = True	
'		query = "Select * from AuthoriserQueue where EntityRowID='" & erID & "' and AuthoriserEntityTypeID ='" & eTypeid & "' order by 1 desc;"
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'		inauthQueueID = dbRecordSet("AuthoriserQueueID")
'		inqueueStatusID = dbRecordSet("StatusID")
'		inprocessedDate = dbRecordSet("ProcessedOn")
'		Set dbRecordSet = Nothing
'		If inqueueStatusID = 1  Then
'			Append_TestHTML StepCounter,"Replication script populate AuthoriserQueue table with  AuthoriserEntityTypeID- "&erID ,"Acknowledgement received Expected: 1" & "  Actual: " & inqueueStatusID ,"PASSED"
'		ElseIf inqueueStatusID = "" Then
'			Append_TestHTML StepCounter,"Replication script populate AuthoriserQueue table with  AuthoriserEntityTypeID- "&erID ,"Acknowledgement received Expected: 1" & "  Actual: Empty - move to Processed Table "  ,"PASSED"
'							bRunFlag = False
'		End If
'End Function
'
'Function checkAuthoriserQueueProcessedTable(rID,entityTypeid)
'	On Error Resume Next
'	bRunFlag = True	
'		query = "Select * from AuthoriserQueueProcessed where EntityRowID='" & rID & "' and AuthoriserEntityTypeID ='" & entityTypeid & "' order by 1 desc;"
'		Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'		
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'		authQueueID = dbRecordSet("AuthoriserQueueID")
'		queueStatusID = dbRecordSet("StatusID")
'		processedDate = dbRecordSet("ProcessedOn")
'		
'		Set dbRecordSet = Nothing
'		If queueStatusID = 3 Then
'			Append_TestHTML StepCounter,"Verify Customer record is sync up with CFGW AuthoriserEntityTypeID- "&entityTypeid ,"Acknowledgement received Expected: 3" & "  Actual: " & queueStatusID ,"PASSED"
'		Else
'			Append_TestHTML StepCounter,"Verify Customer record is sync up with CFGW AuthoriserEntityTypeID- "&entityTypeid,"Acknowledgement received Expected: 3" & "  Actual: " & queueStatusID ,"FAILED"
'							bRunFlag = False
'		End If
'End Function
'
'
'Function insertEnryintoOverrideTable(userName)
'	On Error Resume Next
'	bRunFlag = True	
'		query =" Select * from UserAccount where LoginName like ''" & userName & "%' "
'		Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		userID = dbRecordSet("UserID")
'		uName = dbRecordSet("LoginName")
'		Set dbRecordSet = Nothing
'
'		If userID <> "" Then
'			Append_TestHTML StepCounter,"Get UserID from UserAccount Table" ,"User id-"& userID & "User Name-"& uName ,"PASSED"
'			query =" insert into _UserGroupBindingOverride(UserID,GroupID,CompanyID,ModifiedBy,ModifiedByAPP,ModifiedOn) Values("& userID &",10,85,'YM','YM','2021-07-05 12:34:35.683'); "
'			Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_MASTER")
'			Set dbRecordSet = Nothing
'			wait 10
'			query ="Select * from _UserGroupBindingOverride where UserID ="& userID & ";"
'			Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'			grpID = dbRecordSet("GroupID")
'			Set dbRecordSet = Nothing
'			If grpID = 10 Then
'					Append_TestHTML StepCounter,"Verify User entry",userID & "inserted with group id-"& grpID ,"PASSED"
'			Else
'				query =" insert into _UserGroupBindingOverride(UserID,GroupID,CompanyID,ModifiedBy,ModifiedByAPP,ModifiedOn) Values("& userID &",10,85,'YM','YM','2021-07-05 12:34:35.683'); "
'				Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'				Set dbRecordSet = Nothing
'			End If
'		End If
'		
'		Call navigatetoCustomerSummary("ID00000004")
'		Call submenucheckList("Customer Summary","Addresses",False)
'		query =" delete from _UserGroupBindingOverride where UserID="& userID & " ; "
'		Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_MASTER")
'		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		Set dbRecordSet = Nothing
'		Call OpenApplication(url)
'		Call navigatetoCustomerSummary("ID00000004")
'		Call submenucheckList("Customer Summary","Addresses",True)	
'				
'End Function
'
'Function navigatetoCustomerSummary(custERPid)
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_Start") Then
'		Call  navigateStartMenu("Link_Customers","Link_SearchforCustomer","WebLIST_Role")
'		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLIST_Role") Then
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", custERPid
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary") Then
'				Append_TestHTML StepCounter, "Open Customer-"& custERPid, "Successfully navigated to requested customer screen", "PASSED"
'			Else
'						Append_TestHTML StepCounter, "Open Customer Summary Screen Page", "Fail to navigate Customer Summary screen", "FAILED"	
'								bRunFlag = False						
'			
'			End  If	
'		Else
'						Append_TestHTML StepCounter, "Open Customer Search Screen Page", "Fail to navigate Customer Search screen", "FAILED"	
'								bRunFlag = False				
'		End  IF	
'	Else
'						Append_TestHTML StepCounter, "Open Home Page", "Fail to navigate Home screen", "FAILED"	
'								bRunFlag = False			
'	End  IF	
'End Function
'
'
'Function submenucheckList(mainlink,sublink,sublinkflag)
'	On error resume next
'	bFlag = True
'	Set linkobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").Link("html tag:=A","innertext:="&mainlink)
'	Set linkobj1 = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").Link("html tag:=A","innertext:="&sublink)
'	linkobj.Highlight
'	wait 2
'	Setting.webPackage("ReplayType")=2
'	wait 2
'	linkobj.FireEvent "onmouseover"
'	wait 2
'	linkobj.Click
'	wait 2
'	If sublinkflag Then
'		If linkobj1.Exist = False Then
'			Append_TestHTML StepCounter,"Navigate "& sublink & "screen","Expected : Option should disappear Actual : disappeared under customer-"& customerERP_id ,"PASSED"
'		Else
'			Append_TestHTML StepCounter,"Navigate "& sublink & "screen","Expected : Option should disappear Actual : Appeared under customer-"& customerERP_id ,"FAILED"
'		End If
'	Else
'		If linkobj1.Exist Then
'			linkobj1.Highlight
'			wait 2
'			Append_TestHTML StepCounter,"Navigate "& sublink & "screen","Options are displayed correctly under customer-"& customerERP_id ,"PASSED"
'			wait 5
'			linkobj1.Click
'			wait 5
'		End If
'	End If
'
'	Setting.webPackage("ReplayType")=1
'	If Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
'		Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
'		wait 2
'	End If
'	wait 2
'Set linkobj = Nothing
'Set linkobj1 = Nothing
'End Function
'
'
'Function performGroupBindingAction()
'On error resume next
'	bFlag = True
'	Set cardmaresultobj = Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebTable("html id:=ctl00_CPH_multiViewSettings_mvActions_ucCardSearch_grdResults")
'
'			query ="Select * from _Pages where PageID=32584;"
'			Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'			Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'			pName = dbRecordSet("PageName")
'			Set dbRecordSet = Nothing
'			If pName <> "" Then
'				Append_TestHTML StepCounter,"Verify 32584 page Access","PageName is -"& pName ,"PASSED"
'				query ="delete from _GroupPageBinding where GroupID =1 and PageID=32584;"
'				Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_MASTER")
'				wait 10
'				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'				Set dbRecordSet = Nothing
'				Call searchsubmenuUnderCustomer("ID00000004","Card Maintenance","Link_Cardmaintenance","Link_Cardmaintenance")
'				Call clickOntableElement(cardmaresultobj)
'				Call submenucheckList("Card Maintenance","View Card Authorisations and Denial Activity ",False)
'				
'				query = "insert into [dbo].[_GroupPageBinding](GroupID,PageID,RightIdentifier,ModifiedBy,ModifiedByApp,ModifiedOn) Values(1,32584,3,'Config','M3.19q19_ConfigTooli','2021-07-02 10:40:35.493')"
'				Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"
'				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_MASTER")
'				wait 10
'				Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'				
'				Call OpenApplication(url)
'				Call searchsubmenuUnderCustomer("ID00000004","Card Maintenance","Link_Cardmaintenance","Link_Cardmaintenance")
'				Call clickOntableElement(cardmaresultobj)
'				Call submenucheckList("Card Maintenance","View Card Authorisations and Denial Activity ",True)
'			Else
'				Append_TestHTML StepCounter,"Verify 32584 page Access","PageName not found"& pName ,"FAILED"
'								bRunFlag = False			
'			End If
'			
'End Function



