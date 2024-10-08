'Function preCondition_colcoPinAdvice()
'	updatequery_ColcoPINAdvice = "Update ColcoPINAdvice Set DefaultForNewTopLevelCustomer = 1 where PINAdviceTypeID =2 and ColcoID = 1"
'	Append_TestHTML StepCounter, "Update 'Value' column 'SysVarColco' table",updatequery_ColcoPINAdvice, "PASSED"
'	set dictDbResultSet = update_db_query(updatequery_ColcoPINAdvice, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	wait 2
'	updatequery_sysvarcolco = "Update SysVarColco Set PINChangeAllowedByCardholder = 1 , PINChangeAllowedFromFleetPIN = 1 where ColcoID = 1"
'	Append_TestHTML StepCounter, "Update 'Value' column 'SysVarColco' table",updatequery_sysvarcolco, "PASSED"
'	set dictDbResultSet = update_db_query(updatequery_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	wait 2
'End Function

Function preCondition_colcoPinAdvice(ColcoID)
On error resume next
	bFlag = True
	updatequery_ColcoPINAdvice = "Update ColcoPINAdvice Set DefaultForNewTopLevelCustomer = 1 where PINAdviceTypeID =2 and ColcoID = "&ColcoID
	Append_TestHTML StepCounter, "Update 'Value' column 'SysVarColco' table",updatequery_ColcoPINAdvice, "PASSED"
	Call update_db_query(updatequery_ColcoPINAdvice, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	db_value = dictDbResultSet("Value")
	wait 2
	updatequery_SysVarColco = "Update SysVarColco Set value = 1 where SysvarID = 145"
	Append_TestHTML StepCounter, "Update 'Value' column in 'SysVarColco' table",updatequery_SysVarColco, "PASSED"
	Call update_db_query(updatequery_SysVarColco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	db_value = dictDbResultSet("Value")
	wait 2
	updatequery_SysVarColco = "Update SysVarColco Set value = 1 where SysvarID = 146"
	Append_TestHTML StepCounter, "Update 'Value' column in 'SysVarColco' table",updatequery_SysVarColco, "PASSED"
	Call update_db_query(updatequery_SysVarColco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	db_value = dictDbResultSet("Value")
	wait 2
End Function

Function checkCustomer_DBTable(customerERP)
On error resume next
	bFlag = True
	query_customer = "Select * from Customer where CustomerERP = '"&customerERP&"'"
	Append_TestHTML StepCounter, "validate Customer table",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_CustomerID = dictDbResultSet("CustomerID")
	db_DefaultPINAdviceType = dictDbResultSet("DefaultPINAdviceType")
	
	If db_DefaultPINAdviceType = 2 Then
		Append_TestHTML StepCounter, "verify PinAdviceType","Default PIN Advice Type value is '"&db_DefaultPINAdviceType&"', as expected", "PASSED"
	else
		Append_TestHTML StepCounter, "verify PinAdviceType","Default PIN Advice Type value is '"&db_DefaultPINAdviceType&"'", "FAILED"
	End If
	query_custPINChange = "Select * from CustomerPINChange where CustomerID = '"&db_CustomerID&"'"
	Append_TestHTML StepCounter, "validate CustomerPINChange table",query_custPINChange, "PASSED"
	set dictDbResultSet = execute_db_query(query_custPINChange, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_PINChangeAllowedByCardholder = dictDbResultSet("PINChangeAllowedByCardholder")
	db_PINChangeAllowedFromFleetPIN = dictDbResultSet("PINChangeAllowedFromFleetPIN")
	If db_PINChangeAllowedByCardholder = True Then
		Append_TestHTML StepCounter, "verify CustomerPINChange","PINChangeAllowedByCardholder value is '"&db_PINChangeAllowedByCardholder&"', as expected", "PASSED"
	else
		Append_TestHTML StepCounter, "verify CustomerPINChange","PINChangeAllowedByCardholder value is '"&db_PINChangeAllowedByCardholder&"'", "FAILED"
	End If
	If db_PINChangeAllowedFromFleetPIN = True Then
		Append_TestHTML StepCounter, "verify CustomerPINChange","PINChangeAllowedFromFleetPIN value is '"&db_PINChangeAllowedFromFleetPIN&"', as expected", "PASSED"
	else
		Append_TestHTML StepCounter, "verify CustomerPINChange","PINChangeAllowedFromFleetPIN value is '"&db_PINChangeAllowedFromFleetPIN&"'", "FAILED"
	End If
End Function

Function checkCustomer_newSubAccountDBTable(customerERP_id,sub_customerERP_id,cust_erp_sub1)
On error resume next
	bFlag = True
	query_customer = "Select * from Customer where CustomerERP = '"&cust_erp_sub1&"'"
	Append_TestHTML StepCounter, "validate newly created subAccount",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_TopLevelCustomerID = dictDbResultSet("TopLevelCustomerID")
	Append_TestHTML StepCounter, "validate newly created subAccount","TopLevelCustomerID of newly created subAccount '"&db_TopLevelCustomerID&"'", "PASSED"
	
	db_ParentCustomerID = dictDbResultSet("ParentCustomerID")
	Append_TestHTML StepCounter, "validate newly created subAccount","ParentCustomerID of newly created subAccount '"&db_ParentCustomerID&"'", "PASSED"
	
	
	query_customerERP = "Select * from Customer where CustomerERP = '"&customerERP_id&"'"
	Append_TestHTML StepCounter, "validate Customer A",query_customerERP, "PASSED"
	set dictDbResultSet = execute_db_query(query_customerERP, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_CustomerID = dictDbResultSet("CustomerID")
	Append_TestHTML StepCounter, "validate Customer A","Customer ID of Customer A is '"&db_CustomerID&"'", "PASSED"
	query_customerERP_sub = "Select * from Customer where CustomerERP = '"&sub_customerERP_id&"'"
	Append_TestHTML StepCounter, "validate Customer table",query_customerERP_sub, "PASSED"
	set dictDbResultSet = execute_db_query(query_customerERP_sub, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_CustomerID_sub = dictDbResultSet("CustomerID")
	Append_TestHTML StepCounter, "validate Customer B","Customer ID of Customer B is '"&db_CustomerID_sub&"'", "PASSED"
	
	If db_CustomerID = db_TopLevelCustomerID Then
		Append_TestHTML StepCounter, "validate data", "TopLevel Customer ID of newly created sub account '"&db_TopLevelCustomerID& VBCRLF & "is equal to Customer ID of Customer A is '"&db_CustomerID, "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "TopLevel Customer ID of newly created sub account '"&db_TopLevelCustomerID& VBCRLF & "is not equal to Customer ID of Customer A is '"&db_CustomerID,"FAILED"
	End If
	If db_CustomerID_sub = db_ParentCustomerID Then
		Append_TestHTML StepCounter, "validate data", "Parent Customer ID of newly created sub account '"&db_TopLevelCustomerID& VBCRLF & "is equal to Customer ID of Customer B is '"&db_CustomerID, "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "Parent Customer ID of newly created sub account '"&db_TopLevelCustomerID& VBCRLF & "is not equal to Customer ID of Customer B is '"&db_CustomerID,"FAILED"
	End If
End Function

Function verify_PINAdviceType_blankValue()
On error resume next
	bFlag = True
	Flag = True
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Card maintenance' screen", "PASSED"
		Call objectClick("WebLink_NewCard","NewCard link")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_PinAdviceType") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'New Card' screen", "PASSED"
			ui_pinAdvTypeValues = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_PinAdviceType", "GetROProperty", "all items")
			arr_pinAdvTypeValues = split(ui_pinAdvTypeValues,";")
			For i = 0 To ubound(arr_pinAdvTypeValues) Step 1
				If arr_pinAdvTypeValues(i) <> empty Then
					Append_TestHTML StepCounter, "verify 'PinAdviceType' drop down list", "'PinAdviceType' drop down list has the value as '"&arr_pinAdvTypeValues(i)&"'  in the list", "PASSED"
				else
					Flag = false
				End If
			Next
			If Flag = True Then
				Append_TestHTML StepCounter, "verify 'PinAdviceType' drop down list", "'PinAdviceType' drop down does not have an empty value in the list", "PASSED"
			else
				Append_TestHTML StepCounter, "verify 'PinAdviceType' drop down list", "'PinAdviceType' drop down have an empty value in the list", "FAILED"
			End If
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'New Card' screen", "FAILED"
		End  If	
	else
		Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'Card maintenance' screen", "FAILED"	
	End  If
End Function

Function validate_pinAdviceType_dropdown_createCard(customerERP_id)
On error resume next
	bFlag = True
	Flag = True
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")	
	wait 10
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Card maintenance' screen", "PASSED"
		Call objectClick("WebLink_NewCard","NewCard link")
		wait 10
		' WebList_PinAdviceType
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_CardType") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'New Card' screen", "PASSED"
			ui_pinAdvType = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_PinAdviceType", "GetROProperty", "default value")
			Append_TestHTML StepCounter, "Card maintenance", "'PIN Advice Type' field value in UI is '"&ui_pinAdvType&"'", "PASSED"
			query_PINAdvType = "Select * from Customer where CustomerERP = '"&customerERP_id&"'"
			Append_TestHTML StepCounter, "Customer table",query_PINAdvType, "PASSED"
			set dictDbResultSet = execute_db_query(query_PINAdvType, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
			wait 2
			db_defaultPINAdvType = dictDbResultSet("DefaultPINAdviceType")
			Select Case db_defaultPINAdvType
				Case 1:
					ui_defaultPINAdvType = "PAPER"
				Case 2:
					ui_defaultPINAdvType = "EMAIL"
				Case 3:
					ui_defaultPINAdvType = "SMS"
			End Select
			If ui_defaultPINAdvType = Ucase(ui_pinAdvType) Then
				Append_TestHTML StepCounter, "Customer table","The 'DefaultPINAdviceType' value in Customer table is '"&db_defaultPINAdvType&"' ,as expected", "PASSED"
			else
				Append_TestHTML StepCounter, "Customer table","The 'DefaultPINAdviceType' value in Customer table is '"&db_defaultPINAdvType, "FAILED"
			end if 
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'New Card' screen", "FAILED"
		End  If	
	else
		Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'Card maintenance' screen", "FAILED"	
	End  If
End Function

Function pinAdviceType_emptyValue()
On error resume next
	bFlag = True
	Flag = True
	ui_pinAdvTypeAll = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_PinAdviceType", "GetROProperty", "all items")
	Append_TestHTML StepCounter, "Pin Advice Type", "'Fetch all the drop down values from Pin Advice Type field", "PASSED"	
	Append_TestHTML StepCounter, "Pin Advice Type", "'Verify any blank value found in 'Pin Advice Type' drop down", "PASSED"		
	arr_ui_pinAdvTypeAll = split(ui_pinAdvTypeAll,";")
	For i = 0 To ubound(arr_ui_pinAdvTypeAll) Step 1
		If arr_ui_pinAdvTypeAll(i) <> empty Then
			Append_TestHTML StepCounter, "Pin Advice Type", "'PinAdviceType' drop down value in the list '"&arr_ui_pinAdvTypeAll(i)&"'", "PASSED"
		else
			Flag = false
		End If
	Next
	If Flag = True Then
	Append_TestHTML StepCounter, "Pin Advice Type", "'PinAdviceType' drop down does not have an empty value in the list", "PASSED"
	else
	Append_TestHTML StepCounter, "Pin Advice Type", "'PinAdviceType' drop down have an empty value in the list", "FAILED"
	End If
			
End  Function

Function createNewCardGroup(cardGrpName,CardType,cardDelPoint)
On error resume next
	bFlag = True
	Call pageNavigation("Card Groups","WebLink_CardGroup")
	Call objectClick("Epin_WebLink_NewCGroup","New Card Group link")
	Call enterTextbox_value("Epin_WebEdit_CGName","Card Group Name",cardGrpName)
'	Call enterWebList_value("Epin_WebList_CGTypes","Card Type",CardType)
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WebList_CGTypes", "RadioSelect", CardType
'	listval = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WebList_CGTypes","GetROProperty","default value")
'	Append_TestHTML StepCounter, "Select value", "Select the 'Card Type' value as '"&listval&"'", "PASSED"	
	Call operateOnCheckBox("EpinWebCheckBox_CGCardDeliveryPoint","Card Delivery Point",cardDelPoint)		
End Function

Function select_existingCardGroup(tblElement,CardType,cardDelPoint)
On error resume next
	bFlag = True
	Call pageNavigation("Card Groups","WebLink_CardGroup")
	wait 10
	If Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_mvMain_gvCardGroups").Exist then			
		Call getTblEmement(".*gvCardGroups","Card Group" ,tblElement)	
		Call operateOnCheckBox("EpinWebCheckBox_CGCardDeliveryPoint","Card Delivery Point",cardDelPoint)
	else
		Append_TestHTML StepCounter, "Card Groups", "User is not displayed with 'Card Group' table", "FAILED"
	end if 		
End Function
Function getColcoLevelRegionName()
	On error resume next
	bFlag = True
		query = "Select * from SysVarColCo where SysvarID = 145"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		Db_ColcoID =  dbRecordSet("ColcoID")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from Company where CompanyID=" & Db_ColcoID & ";"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		Db_CountryID =  dbRecordSet("CountryID")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		query = "Select * from Region where CountryID="& Db_CountryID & ";"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		Db_RegionName =  dbRecordSet("RegionName")
				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
		If Db_RegionName <> "" Then
			getColcoLevelRegionName = Db_RegionName
		Else
			getColcoLevelRegionName = ""
		End If
End Function
Function cardDelivery_overRideAddress(new_countryName,address,postalCode,email,phoneNo)
On error resume next
	bFlag = True
		
		
	Call objectClick("Epin_WEl_CGOverride","Override Address tab")
	wait 3
	regval = getColcoLevelRegionName()
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Epin_WC_CGuCDAobj") Then
		Append_TestHTML StepCounter, "Override Address screen", "User navigates to 'Override Address' screen", "PASSED"		
		ui_prop = Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*ucAddressMaintainDelivery_txtAddressLines").GetRoproperty("disabled")
		If ui_prop = 0 Then
			Call enterWebList_value("Epin_WE_CGDcountry","Country Name",new_countryName)
			Call enterTextbox_value("Epin_WE_CGDadress","Address",address)
			Call enterTextbox_value("Epin_WE_CCityval","City",countryName)
			Call enterWebList_value("Epin_WL_CRegionval","Region",regval)
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
			Call enterTextbox_value("Epin_WE_PCityval","City",countryName)
			Call enterWebList_value("Epin_WL_PRegionval","Region",regval)
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
	
Function DBvalidation_cardGroup(cardGrpName,AddressLines,Zipcode,EmailAddress,MobilePhone)
On error resume next
	bFlag = True
	query_cardGroup = "Select * from CardGroup where CardGroupName = '"&cardGrpName&"' order by 1 desc"
	Append_TestHTML StepCounter, "CardGroup table",query_cardGroup, "PASSED"
	set dictDbResultSet = execute_db_query(query_cardGroup, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_CardGroupID = dictDbResultSet("CardGroupID")
	If Instr(db_CardGroupID,"|") > 1 Then
		arr_db_CardGroupID = split(db_CardGroupID,"|")
		db_CardGroupID = arr_db_CardGroupID(0)
	End If
	query_cardGroupAddr = "Select * from CardGroupAddress where CardGroupID = '"&db_CardGroupID&"'"
	Append_TestHTML StepCounter, "CardGroupAddress table",query_cardGroupAddr, "PASSED"
	set dictDbResultSet = execute_db_query(query_cardGroupAddr, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_AddressLines = dictDbResultSet("AddressLines")
	db_Zipcode = dictDbResultSet("Zipcode")
	db_EmailAddress = dictDbResultSet("EmailAddress")
	db_MobilePhone = dictDbResultSet("MobilePhone")
	
	If AddressLines = db_AddressLines Then
		Append_TestHTML StepCounter, "validate data", "UI - Address Lines: "&AddressLines& VBCRLF & "DB - Address Lines: "&db_AddressLines, "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "UI - Address Lines: "&AddressLines& VBCRLF & "DB - Address Lines: "&db_AddressLines, "FAILED"
	End If
	If Zipcode = db_Zipcode Then
		Append_TestHTML StepCounter, "validate data", "UI - Zipcode: "&Zipcode& VBCRLF & "DB - Zipcode: "&db_Zipcode, "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "UI - Zipcode: "&Zipcode& VBCRLF & "DB - Zipcode: "&db_Zipcode, "FAILED"
	End If
	If EmailAddress = db_EmailAddress Then
		Append_TestHTML StepCounter, "validate data", "UI - EmailAddress: "&EmailAddress& VBCRLF & "DB - EmailAddress: "&db_EmailAddress, "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "UI - EmailAddress: "&EmailAddress& VBCRLF & "DB - EmailAddress: "&db_EmailAddress, "FAILED"
	End If
	If MobilePhone = db_MobilePhone Then
		Append_TestHTML StepCounter, "validate data", "UI - MobilePhone: "&MobilePhone& VBCRLF & "DB - MobilePhone: "&db_MobilePhone, "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "UI - MobilePhone: "&MobilePhone& VBCRLF & "DB - MobilePhone: "&db_MobilePhone, "FAILED"
	End If
End  function

Function createCard(cardGroup,cardCategory,embossType,driverName,countryName,pinAdviceType,pinSelectionMethod)
On error resume next
	bFlag = True
	Flag = True
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Card maintenance' screen", "PASSED"
		Call objectClick("WebLink_NewCard","NewCard link")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_TypeOfPin") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'New Card' screen", "PASSED"
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_NewCard_Cardgroup") Then
				Call enterWebList_value("WebList_NewCard_Cardgroup","Card Group",cardGroup)
			Else
				Browser("creationTime:=1").Page("creationTime:=1").Link("innertext:=Select an Option","html tag:=A").Click
				wait 2
				Browser("creationTime:=1").Page("creationTime:=1").WebElement("html id:=ctl00_CPH_multiViewSettings_ucCardMaintain_mvCard_ddlCardGroupID.*","innertext:="&cardGroup).Click
				wait 4
			End If
			
			Call enterWebList_value("WebList_CardType","Card Category",cardCategory)			
			Call enterWebList_value("WebList_EmbossType","Emboss Type",embossType)
			Call enterTextbox_value("WebEdit_DriverName","DriverName",driverName)			
			Call enterWebList_value("WebList_TypeOfPin","TypeOfPin",countryName)
			Call enterWebList_value("WebList_PinAdviceType","Pin Advice Type",pinAdviceType)
			Call enterWebList_value("WebList_PinSelectionMethod","Pin Selection Method",pinSelectionMethod)
			Call click_on_save()
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does  not navigates to 'New Card' screen", "FAILED"
		End  IF
	else
		Append_TestHTML StepCounter, "Card maintenance", "User does  not navigates to 'Card maintenance' screen", "FAILED"
	End  IF
End  Function

Function verify_pinDeliveryAddress_enabled()
On error resume next
	bFlag = True
	Call objectClick("Epin_overridetab","Override Address tab")
	Browser("creationTime:=1").Page("creationTime:=1").webElement("xpath:=//LI/A[normalize-space()='Override Address']").click
	UCDACheckval = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WC_CGuCDAobj","GetROProperty","checked")
	If UCDACheckval = "0" Then	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WC_CGuCDAobj", "Click", ""
	End If
	wait 2
	Append_TestHTML StepCounter, "Verify UCDA check box", "'Use Card Delivery Address' Check box is ticked" , "PASSED"
	Append_TestHTML StepCounter, "Verify Pin delivery Address", "'Pin delivery Address' fields are readonly", "PASSED"
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WC_CGuCDAobj", "Click", ""
	Append_TestHTML StepCounter, "Verify UCDA check box", "Uncheck 'Use Card Delivery Address' Check box" , "PASSED"
	Append_TestHTML StepCounter, "Verify Pin delivery Address", "'Pin delivery Address' fields are enabled", "PASSED"		
End Function

Function createCard_overRideAddress_DeliveryCard(DelOneTimeUse,email,phoneNo,new_countryName,postalCode)
On error resume next
	bFlag = True
	Call objectClick("Epin_overridetab","Override Address tab")
	Append_TestHTML StepCounter, "Page Navigation", "User navigates to 'Override Address' screen", "PASSED"
	wait 3
	ui_countryProp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_Dcountry", "GetROProperty", "disabled")
	
	If ui_countryProp <> 1 Then		
		Append_TestHTML StepCounter, "Page Navigation", "Enter 'Card Delivery' detals", "PASSED"
		Call enterWebList_value("Epin_WE_Dcountry","Card Delivery - Country Name",new_countryName)	
		Call enterTextbox_value("WebEdit_overrideAddress","Card Delivery - Address",address)
		Call enterTextbox_value("Epin_WE_Dzipcode","Card Delivery - Postal Code",postalCode)		
		Call enterTextbox_value("WedEdit_DeliveryEmail","Card Delivery - Emai lAddress",email)
		Call enterTextbox_value("WedEdit_DeliveryMobilePhone","Card Delivery - Mobile Phone",phoneNo)
		Call operateOnCheckBox("WebCheckbox_DeliveryOneTimeUse","Card Delivery - OneTimeUse",DelOneTimeUse)
	else
		Call objectClick("WebElement_CardDelAddrSearch","Card Delivery Address Search element")
		wait 5
		Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").Webtable("html id:=ctl00_CPH_grdAddress")
		If tableObj.Exist then
			Append_TestHTML StepCounter, "Override Address screen", "User is displayed with 'Card Delivery Address' table", "PASSED"	
			Set desc = Description.Create
			desc("micclass").value = "WebElement"
			Set chdObj = tableObj.ChildObjects(desc)			
			For i = 1 To chdObj.count Step 1
				If Ucase(trim(chdObj(i).GetRoProperty("innertext"))) = Ucase(new_countryName) Then
					chdObj(i).click
					Append_TestHTML StepCounter, "Override Address screen", "click on particular address in 'Card Delivery Address' table", "PASSED"	
					Exit For
				End  If
			Next
		else
			Append_TestHTML StepCounter, "Override Address screen", "User default address is displayed with 'Card Delivery Address' table", "PASSED"	
		End  If
	End  If
End  Function

Function createCard_overRideAddress_DeliveryPin(PinOneTimeUse,email,phoneNo,new_countryName,address,postalCode,city,region)

On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebCheckbox_UseCardDeliveryAddress") Then
		Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress","0")
		Append_TestHTML StepCounter, "Card Delivery", "Enter 'Pin Delivery' details", "PASSED"
		Call enterWebList_value("Epin_WLPcountry","Pin Delivery - Country Name",new_countryName)	
		Call enterTextbox_value("Epin_WC_Paddress","Pin Delivery - ContactName",address)
		Call enterTextbox_value("Epin_WE_Pzipcode","Pin Delivery - postalCode",postalCode)
		Call enterWebList_value("WedList_PinRegion","Pin Delivery - Region",region)
		Call enterTextbox_value("WedEdit_PinCity","Pin Delivery - City",city)
		
		'Call enterTextbox_value("WebEdit_PinDeliveryEmailAddress","Pin Delivery - EmailAddress",email)
		Call enterTextbox_value("WebEdit_PinDeliveryPhoneno","Pin Delivery - MobilePhone",phoneNo)
		Call operateOnCheckBox("WebCheckbox_PinOneTimeUse","OneTimeUse - Pin Delivery",PinOneTimeUse)	
	else
		Append_TestHTML StepCounter, "Page Navigation", "Pin Delivery fields are not enabled", "FAILED"
	End  If
End  Function

Function  perform_ReprintPIN()
On error resume next
	bFlag = True
	Browser("creationTime:=1").Page("creationTime:=1").Link("innertext:=Reprint PIN.*").Click
	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//FORM[@id='aspnetForm']/DIV[16]/UL[1]/LI[1]/SPAN[1]").exist Then
		Append_TestHTML StepCounter, "PIN reminder pop up", "User is displayed with 'Are you sure you want to send a paper PIN reminder?' pop up", "PASSED"		
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//FORM[@id='aspnetForm']/DIV[16]/UL[1]/LI[1]/SPAN[1]").Click
		Append_TestHTML StepCounter, "PIN reminder pop up", "Click on 'Yes' button in pop up", "PASSED"
	End  If
	
	
End  function

Function  DBvalidation_ReprintPIN()
On error resume next
	bFlag = True
	query_customer = "Select * from Card where PAN = '"&cardPANNum&"'"
	Append_TestHTML StepCounter, "Card table",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_CardID = dictDbResultSet("CardID")
	query_customer = "Select * from CardPINMailerRequest where CardID = '"&db_CardID&"'"
	Append_TestHTML StepCounter, "Card table",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	db_PINAdviceTypeID = dictDbResultSet("PINAdviceTypeID")
	If db_PINAdviceTypeID = 1 Then
		Append_TestHTML StepCounter, "validate data", "As expected, PINAdviceTypeID value is 1 in CardPINMailerRequest table", "PASSED"
	else
		Append_TestHTML StepCounter, "validate data", "As expected, PINAdviceTypeID value is not 1 in CardPINMailerRequest table", "FAILED"
	End If	
End  function

Function verifyFields_cardTypeScreen()
On error resume next
	bFlag = True
	Call  navigateStartMenu("UILink_CardTypes","UILink_SearchCardTypes","UI_WebEdit_CardTypes")
'	Call navigateWithStartMenu("WebLink_CardTypes","WebLink_SearchCardTypes","WebLink_Search","Card Types","Search Card Types")
	Call objectClick("WebLink_Search","Search link")
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*grdResults")
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Search Card Types", "User is displayed with 'Search Card Types' table", "PASSED"
		Call operateOnTblElement(".*grdResults","Search Card Types","Yes")
		Call validate_checkBox_disabled("WebCheckbox_EMVContactless","EMVContactless")
		Call validate_checkBox_disabled("WebCheckbox_RFID","RFID")
		Call validate_checkBox_disabled("WebCheckbox_VirtualCard","VirtualCard")
'		Call objectClick("WebElement_CardSettings","CardSettings tab")
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//LI/A[normalize-space()='Card Settings']").Click
		Append_TestHTML StepCounter, "Search Card Types", "Click on 'CardSettings tab'", "PASSED"
		wait 5
		Call validate_checkBox_disabled("WebCheckbox_IsPINChangeSupported","IsPINChangeSupported")
'		a = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*grdResults").GetCellData(2,3)
'		Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*grdResults").ChildItem(2,3,"WebElement",0).click
	else
		Append_TestHTML StepCounter, "Search Card Types", "User is not displayed with 'Search Card Types' table", "FAILED"
	End If
	
End Function

Public Function ui_changes_CreateSubCust(cust_erp, cust_erp_sub,fullName,shortName,tradingName)
	On error resume next	
	bFlag = True	
	wait 1
	filePath = sCurrentDirectory & "Test Data\customer_"& countryCode & ".txt"
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set fileRead = fileSysObj.OpenTextFile(filePath,1)
	content = fileRead.ReadAll
	fileRead.Close
	Set DXwrite = fileSysObj.OpenTextFile(filePath,2)
	Newcontent = content + 1
	registrationnumber1 = Newcontent
	customername = cust_data("fullName") & registrationnumber1
	DXwrite.Write Newcontent
	DXwrite.Close
	wait 1
	Call pageNavigation("Customer Details","Link_CustomerDetails")
'	Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").Link("New Sub Account").Click
	Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").Link("New Sub Account").Click
	If appEnvName = "RD" Then
		Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("New Sub Account").Click
	Else
		Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("New Sub Account").Click
	End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_No", "Click", ""
	End If
'	Call objectClick("WebLink_NewSubAccnt","New Sub Account link")	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_SubFullName") Then
		Append_TestHTML StepCounter, "Create sub Customer", "User navigates to 'create sub customer' page", "PASSED"
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SubFullName", "Set", fullName	' cust_name & registrationnumber1
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SubShortName", "Set",	shortName		' cust_name & registrationnumber1
		
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SubTradeName", "Set", tradingName				'cust_name & registrationnumber1
	
	cust_reg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RegNum", "GetROProperty", "value")
	
	cust_reg2 = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Reg2Num", "GetROProperty", "value")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RegNum", "Set", registrationnumber1 'cust_reg+"0"
		
		
		
'		Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").WebEdit("FullName").Set fullName
'		Append_TestHTML StepCounter, "Enter value", "Enter the 'Full Name' value as 'FullName"&fullName&"'", "PASSED"	
'		Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").WebEdit("ShortName").Set shortName
'		Append_TestHTML StepCounter, "Enter value", "Enter the 'Short Name' value as 'ShortName"&shortName&"'", "PASSED"		
'		Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").WebEdit("TradingName").Set tradingName
		Append_TestHTML StepCounter, "Enter value", "Enter the 'Trading Name' value as 'TradingName"&tradingName&"'", "PASSED"		
'		Call enterTextbox_value("WebEdit_SubFullName","SubFullName",cust_name)
'		Call enterTextbox_value("WebEdit_SubShortName","SubShortName",cust_name)
'		Call enterTextbox_value("WebEdit_SubTradeName","SubTradeName",cust_name)
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'		Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").Link("Save").Click
		If appEnvName = "RD" Then
			Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("Save").Click
			wait 5
			If Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").WebElement("Save").Exist Then
				Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").WebElement("Save").Click
			End If
		Else
			Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Save").Click
			wait 5
			If Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").WebElement("Save").Exist Then
				Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").WebElement("Save").Click
			End If
		End If


		Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
		
'		If Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").WebElement("Save").exist Then
'			Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").WebElement("Save").Click
'			Append_TestHTML StepCounter, "Click on save", "Click on save element in popup", "PASSED"
'		End If	
		wait 3
'		cust_erp_sub1 = Browser("Shell Nederland Verkoopmij. BV").Page("Shell Nederland Verkoopmij. BV").WebEdit("CustomerERP").GetRoProperty("value")
		cust_erp_sub1 = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CustERP", "GetROProperty", "value")
		'cust_erp_sub = "ID00000109"	
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CustERP") Then
			Append_TestHTML StepCounter, "Save sub Customer details", "Sub customer '"&cust_erp_sub&"' is created", "PASSED"
		else
			Append_TestHTML StepCounter, "Save sub Customer details", "Saving sub customer details failed", "FAILED"
		End If
		wait 4
		
		If appEnvName = "RD"  Then
			Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("Customer Structure").Click
		Else
			Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Customer Structure").Click
		End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_No", "Click", ""
	End If
	'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustDetails", "Click", ""
	'	wait 2
	'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustStructure", "Click", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Customer Structure"
'		wait 2
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustomerStructure", "Click", ""
'		wait 2
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_CustMain") Then
			Append_TestHTML StepCounter, "Open customer hierarchy", "Opened customer hierarchy page", "PASSED"
		else
			Append_TestHTML StepCounter, "Open customer hierarchy", "Navigation to customer hierarchy page failed", "FAILED"
		End If
		If appEnvName = "RD"  Then
			Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("Customer Structure").Click
		Else
			Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Customer Structure").Click
		End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_No", "Click", ""
	End If
		cust_main = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustMain", "GetROProperty", "innertext")
		
		cust_sub = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustSub", "GetROProperty", "innertext")
		

		If instr(cust_main, cust_erp) <> 0 Then
			Append_TestHTML StepCounter, "Validate main customer", "Validated main customer successfully. Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & cust_main, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate main customer", "Validation of main customer failed. Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & cust_main, "PASSED"
		End If
		
		If instr(cust_sub, cust_erp_sub) <> 0 Then
			Append_TestHTML StepCounter, "Validate sub customer", "Validated sub customer successfully. Expected Value: " & cust_erp_sub & VBCRLF & "Actual Value: " & cust_sub, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate sub customewr", "Validation of sub customer failed. Expected Value: " & cust_erp_sub & VBCRLF & "Actual Value: " & cust_sub, "PASSED"
'			bRunFlag = False
		End If
	else
		Append_TestHTML StepCounter, "Create sub Customer", "Navigation to create sub customer page failed", "FAILED"
	End If

End Function

Function validate_newCard_email_field(invalidEmail1,useCardDelAddress)
On error resume next
	bFlag = True
	Call objectClick("WebElement_OverrideAddress","Override Address tab")
	Call enterTextbox_value_desc("txtEmailAddressCardDelivery","CardDeliveryEmailAddress",invalidEmail1)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtEmailAddressCardDelivery")
	
	Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
	
	Call enterTextbox_value_desc("txtEmailAddressPinDelivery","PinDeliveryEmailAddress",invalidEmail1)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtEmailAddressCardDelivery")
End Function

Function validate_newCard_PhoneNo_field(invalidPhoneNo1,useCardDelAddress)
On error resume next
	bFlag = True
	Call objectClick("WebElement_OverrideAddress","Override Address tab")
	Call enterTextbox_value_desc("txtMobilePhoneDelivery","CardDelivery Phone number",invalidPhoneNo1)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtMobilePhoneDelivery")
	
	Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
	
	Call enterTextbox_value_desc("txtMobilePhonePinDelivery","PinDelivery Phone number",invalidPhoneNo1)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtMobilePhonePinDelivery")
End Function

Function validate_newCard_PhoneNo_field(invalidPhoneNo1,invalidPhoneNo2,useCardDelAddress)
On error resume next
	bFlag = True
	Call objectClick("WebElement_OverrideAddress","Override Address tab")
	Call enterTextbox_value_desc("txtMobilePhoneDelivery","CardDelivery Phone number",invalidPhoneNo1)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtMobilePhoneDelivery")
	
	Append_TestHTML StepCounter, "Invalid data", "Check for another invalid Phone number", "PASSED"
	Call enterTextbox_value_desc("txtMobilePhoneDelivery","CardDelivery Phone number",invalidPhoneNo2)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtMobilePhoneDelivery")
	
	Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
	
	Call enterTextbox_value_desc("txtMobilePhonePinDelivery","PinDelivery Phone number",invalidPhoneNo1)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtMobilePhonePinDelivery")
	
	Call enterTextbox_value_desc("txtMobilePhonePinDelivery","PinDelivery Phone number",invalidPhoneNo2)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
	Call validate_errorMsg()	
	Call verify_fieldMandatory("txtMobilePhonePinDelivery")
End Function


Public Function validate_activeCard(invalidEmail1,status,useCardDelAddress)
	On error resume next	
	bFlag = True	
	wait 1
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")
	Call enterWebList_value("WebList_SearchCardStatus","Status",status)
	Call objectClick("WebLink_Search","Search link")
	Call operateOnTblElement(".*grdResults","Card Table",status)
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_OverrideAddress") Then
		Append_TestHTML StepCounter, "Card maintenance", "User is displayed with 'Card details'", "PASSED"
		Call objectClick("WebElement_OverrideAddress","Override Address tab")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CardDeliveryEmailAddress") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Override Address' screen", "PASSED"
			Call enterTextbox_value("WebEdit_CardDeliveryEmailAddress","CardDelivery EmailAddress",invalidEmail1)
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtEmailAddressCardDelivery")
			
			Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
			
			Call enterTextbox_value("WebEdit_PinDeliveryEmailAddress","PinDelivery EmailAddress",invalidEmail1)	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtEmailAddressPinDelivery")	
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'Override Address' screen", "FAILED"
		end if
	else
		Append_TestHTML StepCounter, "Card maintenance", "User is not displayed with 'Card details'", "FAILED"
	end if		
End Function

Public Function validate_activeCard_phoneNo(invalidPhoneNo1,invalidPhoneNo2,status,useCardDelAddress)
	On error resume next	
	bFlag = True	
	wait 1
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")
	wait 4
	Call enterWebList_value("WebList_SearchCardStatus","Status",status)
	wait 2
	Call objectClick("WebLink_Search","Search link")
	wait 2
	Call operateOnTblElement(".*grdResults","Card Table",status)
	wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_OverrideAddress") Then
		Append_TestHTML StepCounter, "Card maintenance", "User is displayed with 'Card details'", "PASSED"
		Call objectClick("WebElement_OverrideAddress","Override Address tab")
		wait 2
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CardDeliveryPhoneno") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Override Address' screen", "PASSED"
			Call enterTextbox_value("WebEdit_CardDeliveryPhoneno","CardDelivery Phone number",invalidPhoneNo1)
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhoneDelivery")
			
			Call enterTextbox_value("WebEdit_CardDeliveryPhoneno","CardDelivery Phone number",invalidPhoneNo2)
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhoneDelivery")
			
			Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
			
			Call enterTextbox_value("WebEdit_PinDeliveryPhoneno","PinDelivery Phone number",invalidPhoneNo1)	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhonePinDelivery")	
			
			Call enterTextbox_value("WebEdit_PinDeliveryPhoneno","PinDelivery Phone number",invalidPhoneNo2)	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhonePinDelivery")
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'Override Address' screen", "FAILED"
		end if
	else
		Append_TestHTML StepCounter, "Card maintenance", "User is not displayed with 'Card details'", "FAILED"
	end if		
End Function

Function validate_errorMsg()
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_errorMsg") Then
		ui_errorMsg = Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=hidden error stick").GetROProperty("innertext")
		Append_TestHTML StepCounter, "Error message", "User is displayed with error message '"&ui_errorMsg&"'", "PASSED"
	else
		Append_TestHTML StepCounter, "Error message", "User is not displayed with error message '"&ui_errorMsg&"'", "FAILED"
	End  if
End Function

Function verify_fieldMandatory(fieldID)
On error resume next
	bFlag = True
	ui_emailProp = Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*"&fieldID).GetROProperty("class")
	if Instr(ui_emailProp,"error") > 1 Then
		Append_TestHTML StepCounter, "EmailAddress field", "CardDelivery - EmailAddress field is displayed with red mark", "PASSED"
	else
		Append_TestHTML StepCounter, "EmailAddress field", "CardDelivery - EmailAddress field is not displayed with red mark", "FAILED"
	End  If
End Function

Function cardGroup_emailAddressField(invalidEmail1,useCardDelAddress)
On error resume next
	bFlag = True
	Call objectClick("Epin_WEl_CGOverride","Override Address tab")
	wait 3
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Epin_WC_CGuCDAobj") Then
		Append_TestHTML StepCounter, "Override Address screen", "User navigates to 'Override Address' screen", "PASSED"	
		Call enterTextbox_value("Epin_WE_CGDEmail","Card Delivery - Email Address",invalidEmail1)		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Save' link", "PASSED"
		Call validate_errorMsg()	
		Call verify_fieldMandatory("txtEmailDelivery")
		Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
		Call enterTextbox_value("Epin_WE_CGPemail","Pin Delivery - Email Address",invalidEmail1)
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Save' link", "PASSED"
		Call validate_errorMsg()	
		Call verify_fieldMandatory("txtEmailPin")		
	else
		Append_TestHTML StepCounter, "Override Address screen", "User does not navigates to 'Override Address' screen", "FAILED"	
	End  If		
End  Function 

Function cardGroup_phoneNoField(invalidPhone1,invalidPhone2,useCardDelAddress)
On error resume next
	bFlag = True
	Call objectClick("Epin_WEl_CGOverride","Override Address tab")
	wait 3
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Epin_WC_CGuCDAobj") Then
		Append_TestHTML StepCounter, "Override Address screen", "User navigates to 'Override Address' screen", "PASSED"	
		Call enterTextbox_value("WebEdit_CardDeliveryPhoneno","Card Delivery - Phone number",invalidPhone1)		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Save' link", "PASSED"
		Call validate_errorMsg()	
		Call verify_fieldMandatory("txtMobilePhoneDelivery")
		Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
		Call enterTextbox_value("WedEdit_PinMobilePhone","Pin Delivery - Email Address",invalidPhone1)
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Save' link", "PASSED"
		Call validate_errorMsg()	
		Call verify_fieldMandatory("txtMobilePhonePin")		
	else
		Append_TestHTML StepCounter, "Override Address screen", "User does not navigates to 'Override Address' screen", "FAILED"	
	End  If		
End  Function 

Public Function validate_cardGroup_existingCard(invalidEmail1,cardDelPoint)
	On error resume next	
	bFlag = True	
	wait 1
	Call pageNavigation("Card Groups","WebLink_CardGroup")
	wait 10
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*gvCardGroups")
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Card Groups", "User is displayed with 'Card group table'", "PASSED"
		Rcount = tableObj.RowCount
		If Rcount > 1 Then
			Append_TestHTML StepCounter, "Card Groups", "Card group details found in 'Card group table'", "PASSED"
			mycgval=tableObj.GetCellData(2,1)
			Call operateOnTblElement(".*gvCardGroups","Card Group Table",mycgval)
			Call operateOnCheckBox("EpinWebCheckBox_CGCardDeliveryPoint","Card Delivery Point",cardDelPoint)	
			If Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//LI/A[normalize-space()='Override Address']").exist Then
				Append_TestHTML StepCounter, "Card Groups", "User is displayed with 'Card details'", "PASSED"
				Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//LI/A[normalize-space()='Override Address']").click
				Append_TestHTML StepCounter, "Card Groups", "Click on 'Override Address' tab", "PASSED"
				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CardDeliveryEmail") Then
					Append_TestHTML StepCounter, "Card Groups", "User navigates to 'Override Address' screen", "PASSED"
					Call enterTextbox_value_desc("txtEmailDelivery","CardDelivery EmailAddress",invalidEmail1)	
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
					Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
					Call validate_errorMsg()	
					Call verify_fieldMandatory("txtEmailDelivery")					
					Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress","0")
					Call enterTextbox_value_desc("txtEmailPin","PinDelivery EmailAddress",invalidEmail1)	
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
					Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
					Call validate_errorMsg()	
					Call verify_fieldMandatory("txtEmailPin")
				else
					Append_TestHTML StepCounter, "Card Groups", "User does not navigates to 'Override Address' screen", "FAILED"
				End  If
			else
				Append_TestHTML StepCounter, "Card Groups", "User doe not navigates to 'Override Address' screen", "FAILED"
			End  If
		else
			Append_TestHTML StepCounter, "Card Groups", "No details found in 'Card group table'", "FAILED"			
		End If
	else
		Append_TestHTML StepCounter, "Card Groups", "User is not displayed with 'Card group table'", "FAILED"			
	End If			
End  Function

Public Function validate_cardGroup_existingCard_phoneNo(invalidPhone1,invalidPhone2,cardDelPoint)
	On error resume next	
	bFlag = True	
	wait 1
	Call pageNavigation("Card Groups","WebLink_CardGroup")
	wait 10
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*gvCardGroups")
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Card Groups", "User is displayed with 'Card group table'", "PASSED"
		Rcount = tableObj.RowCount
		If Rcount > 1 Then
			Append_TestHTML StepCounter, "Card Groups", "Card group details found in 'Card group table'", "PASSED"
			currentcelldata = tableObj.GetCellData(2,1)
			Call operateOnTblElement(".*gvCardGroups","Card Group Table",currentcelldata)
			Call operateOnCheckBox("EpinWebCheckBox_CGCardDeliveryPoint","Card Delivery Point",cardDelPoint)	
			If Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//LI/A[normalize-space()='Override Address']").exist Then
				Append_TestHTML StepCounter, "Card Groups", "User is displayed with 'Card details'", "PASSED"
				Browser("creationTime:=1").Page("creationTime:=1").WebElement("xpath:=//LI/A[normalize-space()='Override Address']").click
				Append_TestHTML StepCounter, "Card Groups", "Click on 'Override Address' tab", "PASSED"
				If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CardDeliveryPhoneno") Then
					Append_TestHTML StepCounter, "Card Groups", "User navigates to 'Override Address' screen", "PASSED"
					Call enterTextbox_value_desc("txtMobilePhoneDelivery","CardDelivery Phone number",invalidPhone1)	
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
					Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
					Call validate_errorMsg()	
					Call verify_fieldMandatory("txtMobilePhoneDelivery")					
					Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress","0")
					Call enterTextbox_value_desc("txtMobilePhonePin","PinDelivery Phone number",invalidPhone1)	
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
					Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
					Call validate_errorMsg()	
					Call verify_fieldMandatory("txtMobilePhonePin")
				else
					Append_TestHTML StepCounter, "Card Groups", "User does not navigates to 'Override Address' screen", "FAILED"
				End  If
			else
				Append_TestHTML StepCounter, "Card Groups", "User doe not navigates to 'Override Address' screen", "FAILED"
			End  If
		else
			Append_TestHTML StepCounter, "Card Groups", "No details found in 'Card group table'", "FAILED"			
		End If
	else
		Append_TestHTML StepCounter, "Card Groups", "User is not displayed with 'Card group table'", "FAILED"			
	End If			
End  Function

Function  precondition_subLevelCustomer(customerERP_id,sub_customerERP_id)	
On error resume next
	bFlag = True
	query_customer = "Select * from Customer where CustomerERP = '"&customerERP_id&"'"
	Append_TestHTML StepCounter, "Customer table",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	db_CustomerID = dictDbResultSet("CustomerID")
	wait 2
	
	query_customer = "Select * from Customer where CustomerERP = '"&sub_customerERP_id&"'"
	Append_TestHTML StepCounter, "Customer table",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	db_sub_CustomerID = dictDbResultSet("CustomerID")
	wait 2
	
	db_PINAdviceTypeID = dictDbResultSet("PINAdviceTypeID")
	query_defaultPINAdvType = "Update Customer Set DefaultPINAdviceType = 2 where CustomerERP = '"&customerERP_id&"'"
	Append_TestHTML StepCounter, "Card table",query_defaultPINAdvType, "PASSED"
	Call Update_db_query(query_defaultPINAdvType, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	
	query_defaultPINAdvType = "Update Customer Set DefaultPINAdviceType = 4 where CustomerERP = '"&sub_customerERP_id&"'"
	Append_TestHTML StepCounter, "Card table",query_defaultPINAdvType, "PASSED"
	Call Update_db_query(query_defaultPINAdvType, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2

	query_CustomerPINChange = "Update CustomerPINChange Set PINChangeAllowedByCardholder = 1 where CustomerID = '"&db_CustomerID&"'"
	Append_TestHTML StepCounter, "CustomerPINChange table",query_CustomerPINChange, "PASSED"
	Call Update_db_query(query_CustomerPINChange, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	
	query_CustomerPINChange = "Update CustomerPINChange Set PINChangeAllowedFromFleetPIN = 1 where CustomerID = '"&db_CustomerID&"'"
	Append_TestHTML StepCounter, "CustomerPINChange table",query_CustomerPINChange, "PASSED"
	Call Update_db_query(query_CustomerPINChange, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	
	query_CustomerPINChange = "Update CustomerPINChange Set PINChangeAllowedByCardholder = 0 where CustomerID = '"&db_sub_CustomerID&"'"
	Append_TestHTML StepCounter, "CustomerPINChange table",query_CustomerPINChange, "PASSED"
	Call Update_db_query(query_CustomerPINChange, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2
	
	query_CustomerPINChange = "Update CustomerPINChange Set PINChangeAllowedFromFleetPIN = 0 where CustomerID = '"&db_sub_CustomerID&"'"
	Append_TestHTML StepCounter, "CustomerPINChange table",query_CustomerPINChange, "PASSED"
	Call Update_db_query(query_CustomerPINChange, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2	
End  function

Function preCondition_HideSSPSDetails()
On error resume next
	bFlag = True
	query_SysVarColco = "Update SysVarColco Set value = 0 where SysVarID = '140'"
	Append_TestHTML StepCounter, "Pre condition",query_SysVarColco, "PASSED"
	Call Update_db_query(query_SysVarColco, 1, "GFN_SHELL_SPRINTQA_PH_BATCH")
	wait 2	
End Function

Function validate_fields_createCard()
On error resume next
	bFlag = True
	Call pageNavigation("Card maintenance","Link_Cardmaintenance")
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").Webtable("html id:=.*grdResults")	
	If tableObj.Exist then
		Append_TestHTML StepCounter, "Card maintenance", "User is displayed with 'Card table' with list of 'Active' cards", "PASSED"	
		Call operateOnTblElement(".*grdResults","Card table","Active")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_TypeOfPin") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Card maintenance' screen", "PASSED"		
			Call fieldPresent("WebCheckbox_IsMagStripe","Is MagStripe")
			Call fieldPresent("WebCheckbox_IsEMVContact","Is EMVContact")
			Call fieldPresent("WebCheckbox_IsEMVContactLess","Is EMVContactLess")
			Call fieldPresent("WebCheckbox_IsRFID","Is RFID")
			Call fieldPresent("WebCheckbox_IsVirtual","Is Virtual")
			
			Call fieldPresent("WebList_TypeOfPin","Type Of Pin")
			Call fieldPresent("WebCheckbox_PinChangeSupported","Pin Change Supported")
			Call fieldPresent("WebList_PinAdviceType","Pin Advice Type")
			Call fieldPresent("WebEdit_EvRfId","Ev RfId")
			Call fieldPresent("WebEdit_EvPrintedNumber","Ev PrintedNumber")
			Call fieldPresent("WebList_PinSelectionMethod","Pin Selection Method")
		else
			Append_TestHTML StepCounter, "Card maintenance", "User is not displayed with 'Card table'", "FAILED"
		End  IF	
	else
		Append_TestHTML StepCounter, "Card maintenance", "User does  not navigates to 'New Card' screen", "FAILED"
	End  IF		
End Function

Function getTblEmement(tblID,tblName,cellData)
On error resume next
	bFlag = True
Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*"&tblID)

	If tableObj.exist Then
		Append_TestHTML StepCounter, "Select table item", "User is displayed with '"&tblName & "' table", "PASSED"				
		rcount = tableObj.Rowcount
		If rcount > 1 Then
			cardGrpName = tableObj.GetCelldata(2,1)
			ui_cellData = tableObj.GetCelldata(2,2)
			Set desc = description.Create
			desc("micclass").value = "WebElement"
			Set childObj = tableObj.ChildObjects(desc)
			chdCount =  childObj.count
			For i = 1 To chdCount Step 1
				ui_cellData =  childObj(i).GetRoproperty("innertext")
				If Ucase(Trim(cellData)) = Ucase(Trim(ui_cellData)) Then
					Append_TestHTML StepCounter, "Select table item", "Search the table data '"&cardGrpName&"' in '"&tblName&"' table", "PASSED"
					childObj(i).highlight
					 childObj(i).click
					 Append_TestHTML StepCounter, "Select table item", "Click on table data '"&cardGrpName&"' in '"&tblName&"' table", "PASSED"
					 Flag  = True
					 Exit For
				End If
			Next
			If Flag  = False Then
				Append_TestHTML StepCounter, "Select table item", "Expected table data '"&cardGrpName&"' is not found in '"&tblName&"' table", "FAILED"
			End If
		End If
	else
		Append_TestHTML StepCounter, "Select table item", "User is not displayed with '"&tblName&"' table",  "FAILED"
	end if	

End Function

Function ui_changes_navigate_searchForCards(strStatus,cardPANNum)
On error resume next
	bFlag = True
	Call PageNavigation("Search For Cards","Link_SearchForCards")	
	Call enterWebList_value("WebList_SearchCardStatus","Card Status",strStatus)
	Call enterTextbox_value("WebEdit_CardNo","Card PAN Number",cardPANNum)
	Call objectClick("WebLink_Search","Search link")
	Call operateOnTblElement(".*grdResults","Card" ,cardPANNum)		
End Function

Function customerCreditLimitMgt_fieldEnabled(enfdelLock1,enfdelLock2)
On error resume next
	bFlag = True
	Call pageNavigation("Customer Credit Limit Management","WebLink_CustCreditLimitMgt")
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_enforceDelinquencyLock", "GetROProperty", "checked")
	If ui_prop = 1 Then
		Call validate_textbox_enabled("WedEdit_LockExemptionDate","Lock Exemption Date")
		Call validate_textbox_enabled("WebEdit_DelinquentSuspensionDays","Delinquent Suspension Days")
		Call validate_textbox_enabled("WebEdit_GraceDays","Grace Days")
		Call operateOnCheckBox("WebCheckbox_enforceDelinquencyLock","Enforce Delinquency Lock",enfdelLock1)
		Call validate_textbox_disabled("WedEdit_LockExemptionDate","Lock Exemption Date")
		Call validate_textbox_disabled("WebEdit_DelinquentSuspensionDays","Delinquent Suspension Days")
		Call validate_textbox_disabled("WebEdit_GraceDays","Grace Days")
		Call operateOnCheckBox("WebCheckbox_enforceDelinquencyLock","Enforce Delinquency Lock",enfdelLock2)
		Call validate_textbox_enabled("WedEdit_LockExemptionDate","Lock Exemption Date")
		Call validate_textbox_enabled("WebEdit_DelinquentSuspensionDays","Delinquent Suspension Days")
		Call validate_textbox_enabled("WebEdit_GraceDays","Grace Days")
	else
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "EnforceDelinquencyLock check box is unchecked",  "FAILED"
	End If	
End Function

Function customerCreditLimitMgt_navigate_anotherScreen(enfdelLock1,enfdelLock2)
On error resume next
	bFlag = True
	Call pageNavigation("Customer Credit Limit Management","WebLink_CustCreditLimitMgt")
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_enforceDelinquencyLock", "GetROProperty", "checked")
	If ui_prop = 1 Then	
		Call operateOnCheckBox("WebCheckbox_enforceDelinquencyLock","Enforce Delinquency Lock",enfdelLock1)
		Call objectClick("WebLink_Save","Save link")
		Call pageNavigation("Contacts","WebLink_Contacts")
		Call pageNavigation("Customer Credit Limit Management","WebLink_CustCreditLimitMgt")
		ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_enforceDelinquencyLock", "GetROProperty", "checked")
		If ui_prop = 0 Then		
			Call operateOnCheckBox("WebCheckbox_enforceDelinquencyLock","Enforce Delinquency Lock",enfdelLock2)
			Call validate_textbox_enabled("WedEdit_LockExemptionDate","Lock Exemption Date")
			Call validate_textbox_enabled("WebEdit_DelinquentSuspensionDays","Delinquent Suspension Days")
			Call validate_textbox_enabled("WebEdit_GraceDays","Grace Days")
		else
			Append_TestHTML StepCounter, "Customer Credit Limit Management", "EnforceDelinquencyLock check box is ticked",  "FAILED"
		end  If
	else
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "EnforceDelinquencyLock check box is unticked",  "FAILED"
	End  If
End Function

Function validateFields_customerCreditLimitManagement(enfdelLock1,DelSusDays)
On error resume next
	bFlag = True
	Call pageNavigation("Customer Credit Limit Management","WebLink_CustCreditLimitMgt")
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_enforceDelinquencyLock", "GetROProperty", "checked")
	If ui_prop = 0 Then
		Call operateOnCheckBox("WebCheckbox_enforceDelinquencyLock","Enforce Delinquency Lock",enfdelLock1)
	End  If
	Call validate_checkBox_enabled("WebCheckbox_OverridesApplied","OverridesApplied")
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_OverridesApplied", "GetROProperty", "checked")
	If ui_prop = 1 Then
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "OverridesApplied check box is checked",  "PASSED"
	End  If
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_DelinquencyOverridesApplied", "GetROProperty", "checked")
	If ui_prop = 1 Then
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "DelinquencyOverridesApplied check box is checked",  "PASSED"
	End  If
	Call enterTextbox_value("WebEdit_DelinquentSuspensionDays","DelinquentSuspensionDays",DelSusDays)
	Call objectClick("WebLink_Save","Save link")	
	Call pageNavigation("Contacts","WebLink_Contacts")
	Call pageNavigation("Customer Credit Limit Management","WebLink_CustCreditLimitMgt")
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_DelinquencyOverridesApplied", "GetROProperty", "checked")
	If ui_prop = 1 Then
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "DelinquencyOverridesApplied check box is checked",  "PASSED"
	else
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "DelinquencyOverridesApplied check box is checked",  "FAILED"
	End  If
	ui_prop = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_OverridesApplied", "GetROProperty", "checked")
	If ui_prop = 0 Then
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "OverridesApplied check box is unchecked",  "PASSED"
	else
		Append_TestHTML StepCounter, "Customer Credit Limit Management", "OverridesApplied check box is checked",  "FAILED"
	End  If
	
End Function


Public Function CreateTempTopLevelCustomer(cust_data)

	On error resume next
	
	bFlag = True
	
	filePath = sCurrentDirectory & "Test Data\customer_"& countryCode & ".txt"
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	Set fileRead = fileSysObj.OpenTextFile(filePath,1)
	content = fileRead.ReadAll
	fileRead.Close
	Set DXwrite = fileSysObj.OpenTextFile(filePath,2)
	Newcontent = content + 1
	registrationnumber = Newcontent
	customername = cust_data("fullName") & registrationnumber
	DXwrite.Write Newcontent
	DXwrite.Close
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Create Top Level Customer"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CreateTopLevelCust", "Click", ""
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Start", "Mover", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Customers", "Mover", ""
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CreateTopLevelCustomer", "Click", ""
	Call  navigateStartMenu("Link_Customers","Link_CreateTopLevelCustomer","WebLIST_LOB")
	wait 1
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_LOB", "RadioSelect", cust_data("lob")

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_FullName", "Set", customername

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ShortName", "Set", cust_data("shortName")

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_TradingName", "Set", cust_data("tradingName")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Language", "RadioSelect", cust_data("langauge")

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_LegalEntity", "RadioSelect", cust_data("legalEnity")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RegNum", "Set", registrationnumber
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_VAT", "Set", cust_data("vat")

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Band", "RadioSelect", cust_data("band")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_MarketSeg", "RadioSelect", cust_data("marketingSeg")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_PriceProfile", "RadioSelect", cust_data("priceProfile")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_FeeGrpName", "RadioSelect", cust_data("feeGroupName")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_AcqChannel", "RadioSelect", cust_data("acqChannel")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_CustClass", "RadioSelect", cust_data("custClassification")
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CreditLimit", "Click", ""
	wait  2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CreditLimit", "Set", cust_data("creditLimit")
	
	If Browser("name:="&browserProp).Page("title:="&pageProp).WebList("html id:=ctl00_CPH_mvMain_ddlBillingLanguage").exist Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_BilingLanguage", "RadioSelect", cust_data("billingLang")
	End  If
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	
	wait 3  ' waiting for the DB update 
	
	cust_erp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CustERP", "GetROProperty", "value")
	
	If cust_erp = empty  Then
		Append_TestHTML StepCounter,"Customer ERP ID not found in UI","Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & " " ,"FAILED"
		bFlag = False
	else
		Append_TestHTML StepCounter,"Customer ERP ID found in UI","Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & cust_erp ,"PASSED"
	End If
	
	query = "select * from Customer where CustomerERP = '" & cust_erp & "';"
	Append_TestHTML StepCounter,"Get Customer Details",query ,"PASSED"
	
'	set dictDbResultSet = execute_db_query(query)
	set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_cust_id = dictDbResultSet("CustomerERP")
	Set dictDbResultSet = Nothing
	
	myonecustomerERP = db_cust_id
	If cust_erp = db_cust_id  Then
		Append_TestHTML StepCounter,"Customer ERP ID match in database","Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & db_cust_id ,"PASSED"
	else
		Append_TestHTML StepCounter,"Customer ERP ID mismatch in database","Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & db_cust_id ,"FAILED"
		bFlag = False
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for creating top level customer","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for creating top level customer","FAILED"
		bRunFlag = False
	End If

End Function



Public Function validate_phoneNo(invalidPhoneNo1,invalidPhoneNo2,status,useCardDelAddress)
	On error resume next	
	bFlag = True	
'	wait 1
'	Call pageNavigation("Card maintenance","Link_Cardmaintenance")
'	wait 4
'	Call enterWebList_value("WebList_SearchCardStatus","Status",status)
'	wait 2
'	Call objectClick("WebLink_Search","Search link")
'	wait 2
'	Call operateOnTblElement(".*grdResults","Card Table",status)
'	wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_OverrideAddress") Then
		Append_TestHTML StepCounter, "Card maintenance", "User is displayed with 'Card details'", "PASSED"
		Call objectClick("WebElement_OverrideAddress","Override Address tab")
		wait 2
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CardDeliveryPhoneno") Then
			Append_TestHTML StepCounter, "Card maintenance", "User navigates to 'Override Address' screen", "PASSED"
			Call enterTextbox_value("WebEdit_CardDeliveryPhoneno","CardDelivery Phone number",invalidPhoneNo1)
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhoneDelivery")
			
			Call enterTextbox_value("WebEdit_CardDeliveryPhoneno","CardDelivery Phone number",invalidPhoneNo2)
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhoneDelivery")
			
			Call operateOnCheckBox("WebCheckbox_UseCardDeliveryAddress","UseCardDeliveryAddress",useCardDelAddress)
			
			Call enterTextbox_value("WebEdit_PinDeliveryPhoneno","PinDelivery Phone number",invalidPhoneNo1)	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhonePinDelivery")	
			
			Call enterTextbox_value("WebEdit_PinDeliveryPhoneno","PinDelivery Phone number",invalidPhoneNo2)	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			Call validate_errorMsg()	
			Call verify_fieldMandatory("txtMobilePhonePinDelivery")
		else
			Append_TestHTML StepCounter, "Card maintenance", "User does not navigates to 'Override Address' screen", "FAILED"
		end if
	else
		Append_TestHTML StepCounter, "Card maintenance", "User is not displayed with 'Card details'", "FAILED"
	end if		
End Function
