Function dbChanges_navigate_cardParameter(useFleetPIN,selectedPIN)
On error resume next
	bFlag = True
	Call pageNavigation("Card Parameters","Link_CardParameters")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebCheckbox_SelectedPIN")  Then
		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Card Parameters' screen", "PASSED"
		Call operateOnCheckBox("WebCheckbox_UseFleetPIN","UseFleetPIN",useFleetPIN)
		Call operateOnCheckBox("WebCheckbox_SelectedPIN","Selected PIN",selectedPIN)
		Call validate_textbox_enabled("WedEdit_CustomerFleetPIN","Customer Fleet PIN")
		Call validate_textbox_enabled("WedEdit_SSPSFormNumber","SSPS Form Number")		
	else
		Append_TestHTML StepCounter, "Fee Rule Tier", "User does not navigates to 'Card Parameters' screen", "FAILED"
	End  If
End  Function

Function dbChanges_DB_queries(cardPANNum)
On error resume next
	bFlag = True
	query_CardID = "Select * from Card where PAN = '"&cardPANNum&"'"
	Append_TestHTML StepCounter, "validate Full name",query_CardID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardID, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_CardID = dictDbResultSet("CardID")
	query_CardPANID = "Select * from CardPAN where PAN = '"&cardPANNum&"'"
	Append_TestHTML StepCounter, "validate Full name",query_CardPANID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardPANID, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_CardPANID = dictDbResultSet("CardPANID")
	If cdbl(db_CardPANID) = (cdbl(db_CardID) + 16777216) Then
		Append_TestHTML StepCounter, "verify CARD PAN ID and Card ID","As expected, Card Pan ID '"&db_CardPANID&"' is equal to summation of Card ID '"&db_CardID&"' and 2^24", "PASSED"
	else
		Append_TestHTML StepCounter, "verify CARD PAN ID and Card ID","As expected, Card Pan ID '"&db_CardPANID&"' is not equal to summation of Card ID '"&db_CardID&"' and 2^24", "FAILED"
	End If
End Function

Function db_changes_verify_cardPAN_activeStatus(strStatus,strDriverName)
On error resume next
	bFlag = True
	Call enterWebList_value("WebList_SearchCardStatus","Card Status",strStatus)
'	Call enterTextbox_value("WebEdit_CardNo","Card PAN Number",cardPANNum)
	Call objectClick("WebLink_Search","Search link")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_SearchForCards")  Then
		Call operateOnTblElement(".*grdResults","Card" ,strStatus)
		wait 5
		If Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*ctl00_CPH_multiViewSettings_ucCardMaintain_mvCard_txtEmbossDriver").exist  Then
			cardPANNum = Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*mvCard_txtPAN").GetROProperty("value")
			Append_TestHTML StepCounter, "Search for Cards", "User navigates to 'Card maintenance' screen", "PASSED"
			Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=ctl00_CPH_multiViewSettings_ucCardMaintain_mvCard_txtEmbossDriver").Set strDriverName		
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			wait 3
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_NewCard_ConfReq")  Then
				conf_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_NewCard_ConfReq","GetROProperty","innertext")	
				Append_TestHTML StepCounter, "Click on save", "User is displayed with '"&conf_msg&"' popup message", "PASSED"
				Call enterWebList_value("WebList_CardReasonReplace","Reason","Replaced")
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
				Append_TestHTML StepCounter, "Click on save", "Click on save button in 'Confirmation Required' popup", "PASSED"
				wait 5
			End  If
		else
			Append_TestHTML StepCounter, "Search for Cards", "User does not navigates to 'Card maintenance' screen", "FAILED"
		End  If
	else
		Append_TestHTML StepCounter, "Search for Cards", "No values displayed for the entered search criteria in Card table", "FAILED"
	End  If	
End Function

Function dbChanges_verifyDB_cardID_activeStatus(cardPANNum)
On error resume next
	bFlag = True
	query_CardID = "Select * from Card where PAN = '"&cardPANNum&"' and StatusID = '31'"
	Append_TestHTML StepCounter, "Execute Query",query_CardID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardID, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_CardID = dictDbResultSet("CardID")
	If db_CardID <> empty Then
		Append_TestHTML StepCounter, "verify Card status","The Card status is changed from '1' to '31' for the existing CardID '"&db_CardID&"'", "PASSED"
	else
		Append_TestHTML StepCounter, "verify Card status","The Card status is not changed from '1' to '31' for the existing CardID '"&db_CardID&"'", "FAILED"
	End If
	
	query_CardID = "Select * from Card where PAN = '"&cardPANNum&"' and StatusID = '10'"
	Append_TestHTML StepCounter, "Execute Query",query_CardID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardID, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_newCardID = dictDbResultSet("CardID")
	If  db_newCardID <> empty Then
		Append_TestHTML StepCounter, "verify New Card","New card with CardID '"&db_newCardID&"' is generated for the PAN no.'"&cardPANNum&"' in Card table", "PASSED"
	else
		Append_TestHTML StepCounter, "verify New Card","New card is not generated for the PAN no.'"&cardPANNum&"' in Card table", "FAILED"
	End If
	dbChanges_verifyDB_cardID_activeStatus = db_CardID
End Function

Function dbChanges_verifyDB_CardPANID_activeStatus(cardPANNum,db_CardID)
On error resume next
	bFlag = True
	sum_db_CardPANID = cdbl(db_CardID) + 16777216
	query_CardPANID = "Select * from CardPAN where PAN = '"&cardPANNum&"'"
	Append_TestHTML StepCounter, "Execute Query",query_CardPANID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardPANID, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_CardPANID = dictDbResultSet("CardPANID")
	If db_CardPANID <> empty Then
		Append_TestHTML StepCounter, "verify Card PAN ID","Card PAN ID '"&db_CardPANID&"' is generated for the PAN no. '"&cardPANNum&"' in CardPAN table", "PASSED"
		If sum_db_CardPANID = db_CardPANID Then
			Append_TestHTML StepCounter, "verify Card PAN ID","The generated Card PAN ID '"&db_CardPANID&"' is equal to summation of Card ID '"&db_CardID&"' and 2^24", "PASSED"
		else
			Append_TestHTML StepCounter, "verify Card PAN ID","The generated Card PAN ID '"&db_CardPANID&"' is not equal to summation of Card ID '"&db_CardID&"' and 2^24", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "verify Card PAN ID","Card PAN ID is not generated for the PAN no. '"&cardPANNum&"' in CardPAN table", "FAILED"
	End If
End Function

Function db_changes_verify_cardPAN_blockedStatus(strStatus,noOfCopies)
On error resume next
	bFlag = True
	Call enterWebList_value("WebList_SearchCardStatus","Card Status",strStatus)
'	Call enterTextbox_value("WebEdit_CardNo","Card PAN Number",cardPANNum)
	Call objectClick("WebLink_Search","Search link")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_SearchForCards")  Then
		Call operateOnTblElement(".*grdResults","Card" ,strStatus)
		cardPANNum = Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*mvCard_txtPAN").GetROProperty("value")
		Call PageNavigation("Copy Card","Link_CopyCard")		
		If Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//FORM[@id='aspnetForm']/DIV[12]/UL[1]/LI[1]/SPAN[1]").exist  Then

			Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//FORM[@id='aspnetForm']/DIV[12]/UL[1]/LI[1]/SPAN[1]").Click
		
		End  if
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_NumberOfCopies")  Then
			Append_TestHTML StepCounter, "Click on save", "User is displayed with 'To copy cards enter the number required below.' popup with 'Number of Cards' text box", "PASSED"
			Call enterTextbox_value("WedEdit_NumberOfCopies","Number Of Copies",noOfCopies)
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CopyCardSave_ConfMsg", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save button in popup", "PASSED"
		End  If
	else
		Append_TestHTML StepCounter, "Search for Cards", "No values displayed for the 'Blocked Card' status", "FAILED"
		Append_TestHTML StepCounter, "Search for Cards", "Unable to proceed further", "FAILED"
	End  If	
End Function

Function click_on_copyCard_tableItem(strNewDriverName)
On error resume next
	bFlag = True
	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=.*fielderror").Exist(6) then
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=.*fielderror").click
		Set desc = Description.Create
		desc("micclass").value = "WebEdit"
		Set childObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_CardEditList_Table").childObjects(desc)
		chdcount = childObj.count
		For i = 1 To chdcount Step 1
			childObj(i).click
			childObj(i).Set strNewDriverName
			Exit For
		Next
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	End  If
	
End  Function

Function dbChanges_verifyDB_cardID_blockedstatus(cardPANNum)
On error resume next
	bFlag = True
	query_CardID = "Select * from Card where PAN = '"&cardPANNum&"' and StatusID = '7'"
	Append_TestHTML StepCounter, "Execute Query",query_CardID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardID, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_CardID = dictDbResultSet("CardID")
	If db_CardID <> empty Then
		Append_TestHTML StepCounter, "verify Card status","Existing CardID '"&db_CardID&"' is found in the Card table with StatusID = '7'", "PASSED"
	else
		Append_TestHTML StepCounter, "verify Card status","Existing CardID '"&db_CardID&"' is not found in the Card table", "FAILED"
	End If
	
	query_CardIDNew = "Select * from Card where PAN = '"&cardPANNum&"' and StatusID = '31'"
	Append_TestHTML StepCounter, "Execute Query",query_CardIDNew, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardIDNew, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_PAN = dictDbResultSet("PAN")
	db_CardIDNew = dictDbResultSet("CardID")
	If  db_PAN <> empty Then
		Append_TestHTML StepCounter, "verify New Card","New Card ID '"&db_CardIDNew&"' is generated for the PAN no. '"&db_PAN&"' with status ID as '31'", "PASSED"
	else
		Append_TestHTML StepCounter, "verify New Card","New Card ID '"&db_CardIDNew&"' is not generated for the PAN no. '"&db_PAN&"' with status ID as '31'",  "FAILED"
	End If
End Function

Function dbChanges_verifyDB_CardPANID_blockedstatus(db_PAN)
On error resume next
	bFlag = True
	CardPANID_val = cdbl(db_CardIDNew) + 16777216
	query_CardPANID = "Select * from CardPAN where CardPANID = '"&CardPANID_val&"'"
	Append_TestHTML StepCounter, "Execute Query",query_CardPANID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardPANID, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_CardPANID = dictDbResultSet("CardPANID")
	db_PAN = dictDbResultSet("PAN")
	If db_CardPANID <> empty Then
		Append_TestHTML StepCounter, "verify Card PAN ID","Card PAN ID '"&db_CardPANID&"' is generated for the PAN no. '"&db_PAN&"' in CardPAN table", "PASSED"
		If cdbl(db_CardPANID) = CardPANID_val Then
			Append_TestHTML StepCounter, "verify CARD PAN ID and Card ID","As expected, Card Pan ID '"&db_CardPANID&"' is equal to summation of Card ID '"&db_CardIDNew&"' and 2^24", "PASSED"
		else
			Append_TestHTML StepCounter, "verify CARD PAN ID and Card ID","As expected, Card Pan ID '"&db_CardPANID&"' is not equal to summation of Card ID '"&db_CardIDNew&"' and 2^24", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "verify Card PAN ID","Card PAN ID is not generated for the PAN no. '"&db_PAN&"' in CardPAN table", "FAILED"
	End If
End Function

Function dbChanges_UIauditCustomerTable_check(fullName)
On error resume next
	bFlag = True
	Call pageNavigation("Customer Details","Link_CustomerDetails")
	wait 4
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebEdit_FullName")  Then
			Append_TestHTML StepCounter, "Click on save", "User navigates to 'Customer Details' screen", "PASSED"
			ui_beforeFullName = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_FullName", "GetROProperty", "value")
			Append_TestHTML StepCounter, "Customer Details screen", "Get the existing Full Name '"&ui_beforeFullName&"' from 'Customer Details' screen", "PASSED"
			Call enterTextbox_value("WebEdit_FullName","Full Name",fullName)			
			Call click_on_save()
			wait 3
	else
		Append_TestHTML StepCounter, "Click on save", "User does not navigates to 'Customer Details' screen", "FAILED"
	End  if
	dbChanges_UIauditCustomerTable_check = ui_beforeFullName
End Function

Function dbChanges_DBauditCustomerTable_DBcheck(ui_beforeFullName,fullName,customerERP)
On error resume next
	bFlag = True
	query_customer = "Select * from Customer where CustomerERP = '"&customerERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_FullName = dictDbResultSet("FullName")
	If db_FullName = fullName Then
		Append_TestHTML StepCounter, "verify customer table","Entered Full Name '"&fullName&"' is updated in Customer table", "PASSED"
	else
		Append_TestHTML StepCounter, "verify customer table","Entered Full Name '"&fullName&"' is not updated in Customer table", "FAILED"
	End If
	
	query_auditCustomer = "Select * from audit.customer WHERE FullName = '"&fullName&"'"
	Append_TestHTML StepCounter, "Execute Query",query_auditCustomer, "PASSED"
	set dictDbResultSet = execute_db_query(query_auditCustomer, 1, "GFN_SHELL_SPRINTQA_ID_REPORTS")
	wait 2
	db_FullName = dictDbResultSet("FullName")
	db_Before_FullName = dictDbResultSet("Before_FullName")
'	db_advType = dictDbResultSet("DefaultPINAdviceType")
	
	If db_Before_FullName = ui_beforeFullName Then
		Append_TestHTML StepCounter, "verify audit customer table","Before Full Name '"&db_Before_FullName&"' is exist in audit.Customer table", "PASSED"
	else
		Append_TestHTML StepCounter, "verify audit customer table","Before Full Name '"&db_Before_FullName&"' is not exist in audit.Customer table",  "FAILED"
	End If
	
	If db_FullName = fullName Then
		Append_TestHTML StepCounter, "verify audit customer table","Full Name '"&db_FullName&"' is exist in audit.Customer table", "PASSED"
	else
		Append_TestHTML StepCounter, "verify audit customer table","Full Name '"&db_FullName&"' is not exist in audit.Customer table",  "FAILED"
	End If
	
'	update_query_customer = "Update customer set DefaultPINAdviceType = '1' where CustomerERP = '"&customerERP&"'"
'	Append_TestHTML StepCounter, "Execute Query",update_query_audit, "PASSED"
'	Call update_db_query(update_query_audit, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	wait 2
	
'	query_auditCustomer = "Select * from audit.customer order by ModifiedOn desc"
'	Append_TestHTML StepCounter, "Execute Query",query_auditCustomer, "PASSED"
'	set dictDbResultSet = execute_db_query(query_auditCustomer, 1, "GFN_SHELL_SPRINTQA_ID_REPORTS")
'	wait 2
'	db_defaultPINadvType = dictDbResultSet("DefaultPINAdviceType")
'	db_Before_defaultPINadvType = dictDbResultSet("Before_DefaultPINAdviceType")
'	If db_Before_defaultPINadvType = db_advType Then
'		Append_TestHTML StepCounter, "verify audit.customer table","Before DefaultPINAdviceType '"&db_Before_defaultPINadvType&"' is exist in audit.Customer table", "PASSED"
'	else
'		Append_TestHTML StepCounter, "verify audit.customer table","Before DefaultPINAdviceType '"&db_Before_defaultPINadvType&"' is not exist in audit.Customer table",  "FAILED"
'	End If
'	If db_defaultPINadvType = db_advType Then
'		Append_TestHTML StepCounter, "verify audit.customer table","DefaultPINAdviceType '"&db_defaultPINadvType&"' is updated in audit.Customer table", "PASSED"
'	else
'		Append_TestHTML StepCounter, "verify audit.customer table","DefaultPINAdviceType '"&db_defaultPINadvType&"' is not updated in audit.Customer table",  "FAILED"
'	End If
	
End Function

Function dbChanges_replication_DB_CardPANtable(cardPANNum,str_dbName)
On error resume next
	bFlag = True
	query_CardID = "Select * from Card where PAN = '"&cardPANNum&"'"
	Append_TestHTML StepCounter, "validate table",query_CardID, "PASSED"
	set dictDbResultSet = execute_db_query(query_CardID, 1, str_dbName)
	wait 2
	db_CardID = dictDbResultSet("CardID")

	
	If db_CardID <> empty  Then
		Append_TestHTML StepCounter, "validate table","As expected, the created Card id '"&db_CardID&"' is found in '"&str_dbName&"' table", "PASSED"
	else
		Append_TestHTML StepCounter, "validate table","The created Card id '"&db_CardID&"' is not found in '"&str_dbName&"' table",  "FAILED"
	End If
End Function

Function navigate_cardMaintenance()
On error resume next
	bFlag = True
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")
'	Call enterTextbox_value("WedEdit_EnterCard","Card PAN Number",cardPANNum)
	'Call objectClick("WebLink_Search","Search link")
	wait 5
End Function

Function navigate_searchForCards(strStatus,cardPANNum)
On error resume next
	bFlag = True
	Call PageNavigation("Search For Cards","Link_SearchForCards")	
	Call enterWebList_value("WebList_SearchCardStatus","Card Status",strStatus)
	Call enterTextbox_value("WebEdit_CardNo","Card PAN Number",cardPANNum)
	Call objectClick("WebLink_Search","Search link")
End Function

Function precondition_blockCard(strStatus,reason)
On error resume next
	bFlag = True
	Call enterWebList_value("WebList_SearchCardStatus","Card Status",strStatus)
	'Call enterTextbox_value("WebEdit_CardNo","Card PAN Number",cardPANNum)
	Call objectClick("WebLink_Search","Search link")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_SearchForCards")  Then
		Call operateOnTblElement(".*grdResults","Card" ,strStatus)
'		Call click_on_save_element()
		Call PageNavigation("Block / Unblock Card","WebLink_Block_UnblockCard")	
'		If VerifyWebObjectExist ("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebELement_Yes") Then
''			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_Yes", "Highlight", ""
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_Yes", "Click", ""
'			Append_TestHTML StepCounter, "Click on save", "Click on 'Yes' in popup ", "PASSED"
'		End  If
		Call enterWebList_value("WebList_Block_UnblockCard","Status","Blocked Card")	
		Call enterWebList_value("WebList_BlockCard_Reason","Reason",reason)
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
'		Call objectClick("WebElement_Search","Search tab")
	End  if	
End Function