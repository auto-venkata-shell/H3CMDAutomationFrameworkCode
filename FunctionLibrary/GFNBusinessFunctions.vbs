Public customerERP_id, cardPAN_no, JobRundate, cardExpiry_date, dx026FinalFName, Trans_BatchID, settllReportDate, dx350_Summary_Docno, dx350_Balance, dx350_Paymentduedate,registrationnumber,customername,custBankNamedesc

Public db_211_job_id , db_264_job_id, 	db_305_job_id, db_bill_doc_no, db_bill_sum_doc_no,db_289_job_id,db_380_job_id,db_settlement_doc_no,db_settlement_sum_doc_no,DX26SessionBillDate,DX26SessionSettleDate

Public srecordid,sorecordid,sERP,newSERP,newSiteOwnerName,newSiteName,CO_Owner_SiteName,CO_SiteName,DO_Owner_SiteName,DO_SiteName,siteeffDate,siteTerminateDate,old_rec_AddressLine,old_rec_City,old_rec_Telphone

Public DX301_FName,DX301_FileLoc,db_bill_sessionid,customer_sys_id,Trans_statement_acc_id,Bill_doc_Number1,Bill_doc_Number2,Bill_doc_Number3

Public DX602_FilePrefixNameval,DX350_FilePrefixNameval,DX451_FilePrefixNameval,site451Name,siteunqNo,siteFName,siteSName

Public customerERP_sub_id, myonecustomerERP

'
'DX301_FileLoc = "\\AEWNW00235iis1.europe.shell.com\outbound\SFN\85\DX301\DX301_GFN_INV_085_001036_20220530_143942.DAT"'III
'DX301_FName = "DX301_GFN_INV_085_001036_20220530_143942.DAT"'III
''
'db_bill_doc_no="A000008475"
'db_bill_sessionid ="402"
'CO_Owner_SiteName = "Autotest2 10000024"
'CO_SiteName = "SHELL Update $95487"
'DO_Owner_SiteName = "Autotest2 10000025"CreateSubCust
'DO_SiteName = "SHELL10000025"
'dx350_Balance= "791"
' dx350_Summary_Docno="35/ID0000004660029/2021"
'dx350_Paymentduedate = "2022-01-16"
''Public TestdataEnv1,TestdataEnv2,Testdatastatus1,Testdatastatus2
''
customerERP_id = "PH50001974"
''customerERP_id = "NL00000001"
''customerERP_sub_id = "PH50001740"
''XML_Cadrid = 151494
''XML_Cardpanid=16928710
'
''customerERP_id = "BE00000347"	'6iii  'PH50001822   'PH50001618
''customerERP_id = "BE00000342"
'customerERP = customerERP_id'IIII
'''''''billReportDate="2021-12-30"
'cardPANNum = "7077861008414000010"
'cardPAN_no = cardPANNum
'cardExpiry_date = "2033-05-31"
'DX26SessionBillDate = "2022-11-14"
'31/01/2026
'JobRundate=date()
''JobRundate="12/29/2021"
'Trans_BatchID =  "240"



'******************************* HEADER ******************************************
' Description : The function to create top level customer in GFN application
' Creator :  Pradeep Kumar
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function CreateTopLevelCustomer(cust_data)

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
	
	wait 5  ' waiting for the DB update 
	
	cust_erp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CustERP", "GetROProperty", "value")
'	msgbox cust_erp
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
	customerERP_id = db_cust_id
	customerERP = db_cust_id
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

'******************************* HEADER ******************************************
' Description : The function to add address for existing customer in GFN application
' Creator :  Pradeep Kumar
' Date : 30th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function AddCustomerAddress(cust_data)

	On error resume next
'	msgbox "Address Funtion"
	bFlag = True
	wait 2
'	Call menuSelection("Customer Details","Addresses")
	If appEnvName = "RD" Then
		Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("Customer Details").Click
		wait 2
		Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("Addresses").Click
		wait 2
	Else
		Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Customer Details").Click
		wait 2
		Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Addresses").Click
		wait 2
	End If
'	msgbox "Save Button Check"
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "New_WebELement_No", "Click", ""
	End If


'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	
'		
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'	
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_data("custERP")
'		
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
''	If customerSearch() Then
'		wait 1
''		Call navigateCustomerSummaryMenuoption("Link_Addresses","Link_NewAddress")
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CustomerDetails", "Mover", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CustomerDetails", "Click", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "","Link_Addresses", "Mover", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Addresses", "Click", ""
		wait 2
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", customerERP_id
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Addresses"
		
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Addresses", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_NewAddress", "Click", ""
	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_AddressBox", "Set", cust_data("address")
	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CityBox", "Set", cust_data("city")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_AddRegion", "RadioSelect", cust_data("region")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ZipCodeBox", "Set", cust_data("zipCode")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_TelephoneBox", "Set", cust_data("telephone")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_Main", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_Registered", "Click", ""
		
		
		Append_TestHTML StepCounter,"Customer Address details ","Maintained all input values" ,"PASSED"
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		
		wait 3   ' waiting for the DB update
		
'	End If 
	' Check if customer address is successfully created
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_AddressList")  Then
		'Call getoutputtabledata(2,cust_data("address"))
		addr_line = cust_data("address") & cust_data("zipCode") & cust_data("city")
		
		web_cust_addr = VerifyObjectExistInWebTbl("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame","WebTable_AddressList", "WebElement", "innertext", addr_line)
		web_cust_addr = getoutputtabledata(1,cust_data("address"))
		If  web_cust_addr = True Then
			Append_TestHTML StepCounter,"Customer Address added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Address addition Failed","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA","FAILED"
			bFlag = False
		End If
	End If
  	web_cust_addr = empty

	query = "select CustomerID from Customer where CustomerERP = '" & cust_data("custERP") & "';"
		Append_TestHTML StepCounter,"Execute Query",query ,"PASSED"

	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_cid = dbRecordSet("CustomerID")

	set dbRecordSet = Nothing
	query1 = "select * from CustomerAddress where CustomerId = '" & db_cid & "';"
'	set dbRecordSet = execute_db_query(query)
		Append_TestHTML StepCounter,"Execute Query",query1 ,"PASSED"

	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_address = dbRecordSet("AddressLines")
	db_zipCode = dbRecordSet("Zipcode")
	db_city = dbRecordSet("City")
	db_telephone = dbRecordSet("Telephone")
	set dbRecordSet = Nothing
	If cust_data("address") = db_address  Then
		Append_TestHTML StepCounter,"Address Line match in database","Expected Value: " & cust_data("address") & VBCRLF & "Actual Value: " & db_address ,"PASSED"
	else
		Append_TestHTML StepCounter,"Address Line mismatch in database","Expected Value: " & cust_data("address") & VBCRLF & "Actual Value: " & db_address ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("zipCode") = db_zipCode  Then
		Append_TestHTML StepCounter,"Zip Code match in database","Expected Value: " & cust_data("zipCode") & VBCRLF & "Actual Value: " & db_zipCode ,"PASSED"
	else
		Append_TestHTML StepCounter,"Zip Code mismatch in database","Expected Value: " & cust_data("zipCode") & VBCRLF & "Actual Value: " & db_zipCode ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("city") = db_city  Then
		Append_TestHTML StepCounter,"City match in database","Expected Value: " & cust_data("city") & VBCRLF & "Actual Value: " & db_city ,"PASSED"
	else
		Append_TestHTML StepCounter,"City Line mismatch in database","Expected Value: " & cust_data("city") & VBCRLF & "Actual Value: " & db_city ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("telephone") = db_telephone  Then
		Append_TestHTML StepCounter,"Telephone match in database","Expected Value: " & cust_data("telephone") & VBCRLF & "Actual Value: " & db_telephone ,"PASSED"
	else
		Append_TestHTML StepCounter,"Telephone mismatch in database","Expected Value: " & cust_data("telephone") & VBCRLF & "Actual Value: " & db_telephone ,"FAILED"
		bFlag = False
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for adding customer address","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for adding customer address","FAILED"
		bRunFlag = False
	End If

End Function

'******************************* HEADER ******************************************
' Description : The function to add contact details for existing customer in GFN application
' Creator :  Pradeep Kumar
' Date : 30th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function AddCustomerContact(cust_data)

	On error resume next
	
	bFlag = True
	Call menuSelection("Addresses","Contacts")
	
'		Call customerSearch() 
'		wait 1
'		Call navigateCustomerSummaryMenuoption("Link_Contacts","Link_Newcontact")
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'''''	
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'''''	wait 1
'''''	
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'''''
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_data("custERP")
'''''	
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'''''	
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Contacts"
'''''	
'''''	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Contacts", "Click", ""
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CustomerDetails", "Mover", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CustomerDetails", "Click", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "","Link_Contacts", "Mover", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Contacts", "Click", ""
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewContact", "Click", ""

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ForeName", "Set", cust_data("foreName")

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SurName", "Set", cust_data("surName")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Email", "Set", cust_data("email")
		Append_TestHTML StepCounter,"Customer Contact details ","Maintained all input values" ,"PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	
	wait 3   ' waiting for the DB update
	
	' Check if customer contact is successfully created
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_ContactList")  Then
		web_cust_contact = VerifyObjectExistInWebTbl("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame","WebTable_ContactList", "WebElement", "innertext", cust_data("email"))
		web_cust_contact = getoutputtabledata(5,cust_data("email"))
		
		If  web_cust_contact = True Then
			Append_TestHTML StepCounter,"Customer Contact added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Contact addition Failed","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA","FAILED"
			bFlag = False
		End If
	End If
  	web_cust_contact = empty
  	
  	query = "select * from CustomerContact where CustomerId = (select customerid from Customer where CustomerERP = '" & cust_data("custERP") & "');"
'	set dbRecordSet = execute_db_query(query)
	set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_foreName = dictDbResultSet("ForeName")
	db_lastName = dictDbResultSet("LastName")
	db_email = dictDbResultSet("EmailAddress")
Set dictDbResultSet=Nothing
	If cust_data("foreName") = db_foreName  Then
		Append_TestHTML StepCounter,"Customer Forename match in database","Expected Value: " & cust_data("foreName") & VBCRLF & "Actual Value: " & db_foreName ,"PASSED"
	else
		Append_TestHTML StepCounter,"Customer Forename mismatch in database","Expected Value: " & cust_data("foreName") & VBCRLF & "Actual Value: " & db_foreName ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("surName") = db_lastName  Then
		Append_TestHTML StepCounter,"Customer Surname match in database","Expected Value: " & cust_data("surName") & VBCRLF & "Actual Value: " & db_lastName ,"PASSED"
	else
		Append_TestHTML StepCounter,"Customer Surname mismatch in database","Expected Value: " & cust_data("surName") & VBCRLF & "Actual Value: " & db_lastName ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("email") = db_email  Then
		Append_TestHTML StepCounter,"Customer Email match in database","Expected Value: " & cust_data("email") & VBCRLF & "Actual Value: " & db_email ,"PASSED"
	else
		Append_TestHTML StepCounter,"Customer Email Line mismatch in database","Expected Value: " & cust_data("email") & VBCRLF & "Actual Value: " & db_email ,"FAILED"
		bFlag = False
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for adding customer contact","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for adding customer contact","FAILED"
		bRunFlag = False
	End If

End Function

'******************************* HEADER ******************************************
' Description : The function to add bank details for existing customer in GFN application
' Creator :  Pradeep Kumar
' Date : 01st December, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function AddCustomerBankDetails(cust_data)

	On error resume next
	
	bFlag = True
	
	flPath = sCurrentDirectory & "Test Data\Bank.txt"
	Set fileSysObj1 = createObject("Scripting.FileSystemObject")
	Set fileRead = fileSysObj1.OpenTextFile(flPath,1)
	content = fileRead.ReadAll
	fileRead.Close
	Set Dwrite = fileSysObj1.OpenTextFile(flPath,2)
	Newcontent = content + 1
	CBankAccnnumber = Newcontent
	Dwrite.Write Newcontent
	Dwrite.Close
	Set fileRead = Nothing
	Set fileSysObj1 = Nothing
	custBankNamedesc = CBankAccnnumber
	Call menuSelection("Contacts","Banks")
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_data("custERP")
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Banks"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Banks", "Click", ""
'	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewBankAcc", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_BankType", "RadioSelect", cust_data("bankType")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_BankName", "RadioSelect", cust_data("bankName")
	
	
'	 Browser("name:="&browserProp).Page("title:="&pageProp).WebEdit("html id:=ctl00_CPH_mvBanks_txtBankBranchName").exist 
	If Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=ctl00_CPH_mvBanks_txtBankBranchName").exist Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_BranchName", "Set", cust_data("branchName")
	End  If

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_BankAddressBox", "Set", cust_data("address")

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_City", "Set", cust_data("city")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_BankRegion", "RadioSelect", cust_data("region")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_PostCode", "Set", cust_data("postCode")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SortCode", "Set", cust_data("sortCode")
	If countryCode <> "BE" Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_AccNum", "Set", CBankAccnnumber
	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ReAccNum", "Set", CBankAccnnumber
	End If
	
	
	Append_TestHTML StepCounter,"Customer Bank Account details ","Maintained all input values" ,"PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	wait 3   ' waiting for the DB update
	
	' Check if bank details updated, if yes click on the created acc num to assign debit mandates	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_BankAccDetails")  Then
		web_acc_num = VerifyObjectExistInWebTbl("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame","WebTable_BankAccDetails", "WebElement", "innertext", cust_data("sortCode"))
		web_acc_num = getoutputtabledata(3,cust_data("sortCode"))
		
		If  web_acc_num = True Then
			WebTbl_ClickCell_Dynamic "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_BankAccDetails", "WebElement", "innertext", cust_data("sortCode")
			'''msgbox Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=ctl00_CPH_mvBanks_gvCustomerBank","html tag:=TABLE").GetCellData(2,4)
			If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("html tag:=TD","innertext:="&cust_data("sortCode")).Exist Then
				Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("html tag:=TD","innertext:="&cust_data("sortCode")).Click
			End If
			Append_TestHTML StepCounter,"Customer Bank Account details added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Bank Account details addition Failed","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA","FAILED"
			bFlag = False
		End If
	End If
	
	If  web_acc_num = True Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_DDMandate", "Click", ""
			Append_TestHTML StepCounter,"Customer direct debit mandates tab navigation ","Navigated to Direct Debit Mandates tab" & "NA" ,"PASSED"
	  	wait 4
	  	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewBankMandate", "Click", ""
			Append_TestHTML StepCounter,"Navigate to new Bank Mandate detail screen","Clicked on button and navigating to the screen" & "NA" ,"PASSED"
	  	
	  	
	  	
	  	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_DDAuthorizerName", "Set", cust_data("ddAuthName")
			Append_TestHTML StepCounter,"Set all required details","Filled with Auth name " & "NA" ,"PASSED"
	  	
	  	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter,"Customer Bank Account details added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
	  	
	  	wait 2
	  	
	  	' Check if customer DD mandate is successfully created
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_ContactList")  Then
			web_cust_dd_mandate = VerifyObjectExistInWebTbl("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame","WebTable_ContactList", "WebElement", "innertext", cust_data("custERP"))
			If  web_cust_dd_mandate = True Then
				Append_TestHTML StepCounter,"Customer DD Mandate added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
			else
				Append_TestHTML StepCounter,"Customer DD Mandate addition Failed","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA","FAILED"
				bFlag = False
			End If
		End If
	  	web_cust_dd_mandate = empty
	
	End If
  	
  	web_acc_num = empty
	wait 2   ' waiting for the DB update
	
	query1 = "select * from CustomerBankAccount where CustomerId = (select customerid from Customer where CustomerERP = '" & cust_data("custERP") & "');"
		Append_TestHTML StepCounter,"Get Bank Account ID from Cusomer Bank Account",query1 ,"PASSED"
	
'	set dbRecordSet1 = execute_db_query(query1)
	set dbRecordSet1 = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	
	wait 2
	
	query2 = "select * from BankAccount where BankAccountID = '" & dbRecordSet1("BankAccountID") & "';"
		Append_TestHTML StepCounter,"Get Bank details",query2 ,"PASSED"
	
'	set dictDbResultSet2 = execute_db_query(query2)
	set dictDbResultSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	
	wait 2
	
	db_acc_num = dictDbResultSet2("AccountNumber")
	db_address = dictDbResultSet2("AddressLine1")
	db_city = dictDbResultSet2("City")
	db_postCode = dictDbResultSet2("PostCode")
	db_sortCode = dictDbResultSet2("SortCode")
	
	If cstr(CBankAccnnumber)= cstr(db_acc_num)  Then
		Append_TestHTML StepCounter,"Account Number match in database","Expected Value: " & CBankAccnnumber & VBCRLF & "Actual Value: " & db_acc_num ,"PASSED"
	else
		Append_TestHTML StepCounter,"Account Number mismatch in database","Expected Value: " & CBankAccnnumber & VBCRLF & "Actual Value: " & db_acc_num ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("address") = db_address  Then
		Append_TestHTML StepCounter,"Address Line match in database","Expected Value: " & cust_data("address") & VBCRLF & "Actual Value: " & db_address ,"PASSED"
	else
		Append_TestHTML StepCounter,"Address Line mismatch in database","Expected Value: " & cust_data("address") & VBCRLF & "Actual Value: " & db_address ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("city") = db_city  Then
		Append_TestHTML StepCounter,"City match in database","Expected Value: " & cust_data("city") & VBCRLF & "Actual Value: " & db_city ,"PASSED"
	else
		Append_TestHTML StepCounter,"City Line mismatch in database","Expected Value: " & cust_data("city") & VBCRLF & "Actual Value: " & db_city ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("postCode") = db_postCode  Then
		Append_TestHTML StepCounter,"Post Code match in database","Expected Value: " & cust_data("postCode") & VBCRLF & "Actual Value: " & db_postCode ,"PASSED"
	else
		Append_TestHTML StepCounter,"Post Code mismatch in database","Expected Value: " & cust_data("postCode") & VBCRLF & "Actual Value: " & db_postCode ,"FAILED"
		bFlag = False
	End If
	
	If cust_data("sortCode") = db_sortCode  Then
		Append_TestHTML StepCounter,"Sort Code match in database","Expected Value: " & cust_data("sortCode") & VBCRLF & "Actual Value: " & db_sortCode ,"PASSED"
	else
		Append_TestHTML StepCounter,"Sort Code mismatch in database","Expected Value: " & cust_data("sortCode") & VBCRLF & "Actual Value: " & db_sortCode ,"FAILED"
		bFlag = False
	End If
	
	query3 = "select * from BankMandates where BankAccountID = '" & dbRecordSet1("BankAccountID") & "';"
'	set dictDbResultSet3 = execute_db_query(query3)
	set dictDbResultSet3 = execute_db_query(query3, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	db_ddName = dictDbResultSet3("DirectDebitName")
	
	If cust_data("ddAuthName") = db_ddName  Then
		Append_TestHTML StepCounter,"DD Authorizer Name match in database","Expected Value: " & cust_data("ddAuthName") & VBCRLF & "Actual Value: " & db_ddName ,"PASSED"
	else
		Append_TestHTML StepCounter,"DD Authorizer Name mismatch in database","Expected Value: " & cust_data("ddAuthName") & VBCRLF & "Actual Value: " & db_ddName ,"FAILED"
		bFlag = False
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for adding customer bank account and mandate details","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for adding customer bank account and mandate details","FAILED"
		bRunFlag = False
	End If
	set dictDbResultSet3 = Nothing
	Set dictDbResultSet2=Nothing
	Set dbRecordSet1= Nothing
End Function

'******************************* HEADER ******************************************
' Description : The function to add payment details for existing customer in GFN application
' Creator :  Pradeep Kumar
' Date : 14th December, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function AddCustomerPaymentDetails(cust_data)

	On error resume next
	
	bFlag = True
	Call menuSelection("Banks","Payments")
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_data("custERP")
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Payments"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Payments", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCreditStatus", "Click", ""
	wait 2
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_Paymentmethod", "RadioSelect", cust_data("paymentmethod")
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_InsuredLimit", "Set", cust_data("insuredLimit")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_RiskType", "RadioSelect", cust_data("riskType")
		Append_TestHTML StepCounter,"Customer Payment details ","Maintained all input values" ,"PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	ui_credit_limit = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CreditLimit", "GetROProperty", "value")
	
	wait 2   ' waiting for the DB update
		Append_TestHTML StepCounter,"Customer Payment details ","Details are entered" ,"PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebImage_Back", "Click", ""
	
	' Check if bank details updated, if yes click on the created acc num to assign debit mandates	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_Yes", "Click", ""
	End If
		Append_TestHTML StepCounter,"Customer Payment details ","Details are saved" ,"PASSED"
	
	' Check if bank details updated, if yes click on the created acc num to assign debit mandates	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_PaymentList")  Then
		web_risk_type = VerifyObjectExistInWebTbl("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","NoFrame", "WebTable_PaymentList", "WebElement", "innertext", cust_data("riskType"))
		
		web_risk_type = getoutputtabledata(1,cust_data("riskType"))
		
		If  web_risk_type = True Then
			Append_TestHTML StepCounter,"Customer Payment details added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Payment details addition Failed","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA","FAILED"
			bFlag = False
		End If
	End If
  	
  	web_risk_type = empty
	
	wait 2   ' waiting for the DB update
	
	query = "select * from CustomerPaymentDetail where CustomerId = (select customerid from Customer where CustomerERP = '" & cust_data("custERP") & "')"
'	set dbRecordSet = execute_db_query(query)
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	
	
	
	wait 2
	
	db_credit_limit = dbRecordSet("CreditLimit")
	ui_credit_limit = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CreditLimitbtn", "GetROProperty", "value")
	set dbRecordSet = Nothing
	If cdbl(ui_credit_limit) = cdbl(db_credit_limit)  Then
		Append_TestHTML StepCounter,"Credit Limit match in database","Expected Value: " & ui_credit_limit & VBCRLF & "Actual Value: " & db_credit_limit ,"PASSED"
	else
		Append_TestHTML StepCounter,"Credit mismatch in database","Expected Value: " & ui_credit_limit & VBCRLF & "Actual Value: " & db_credit_limit ,"FAILED"
		bFlag = False
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for adding customer payment details","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for adding customer payment details","FAILED"
		bRunFlag = False
	End If

End Function


'******************************* HEADER ******************************************
' Description : The function to Verify Manual Fee Creation for existing customer in GFN application
' Creator :  Pradeep Kumar
' Date : 14th December, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

	
	Public Function VerifyManualFeeCreation(cust_data)

	On error resume next
	
	bFlag = True
'					Call navigateCustomerSummaryMenuoption("Search for Fee Rules","WebTable_FeeRulePreview")
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
	wait 1
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_data("custERP")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustSummary", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_FeeRule", "Click", ""
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Fee Rules"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules", "Click", ""

'	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
'		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
'		wait 2
'	End If
	
'	Append_TestHTML StepCounter,"Fee Rule Preview" ,"Verify Fee Rule screen validation" ,"PASSED"
'			
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewFeeRule", "Click", ""
'			Append_TestHTML StepCounter,"Click on New Fee Rulefor Manual Fee" ,"Navigation to New Fee Rule screen" ,"PASSED"
'			
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_FeeSearchNewRule", "Click", ""
'			Append_TestHTML StepCounter,"Search for fee rule ","Clicked on search button" ,"PASSED"
'			
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_FeeRuleDesc", "Set", "Late Payment"
'			Append_TestHTML StepCounter,"Customer Manual New Fee Rule Description ","Maintain Fee Rule Description" ,"PASSED"
'			
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'			Append_TestHTML StepCounter,"Click on Search link for fee rule list ","Navigated to Cust Fee Rule list" ,"PASSED"
'			
'			WebTbl_ClickCell_Dynamic "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_CustFeeRuleList", "WebElement", "innertext", "Late Payment Fee"
'						Append_TestHTML StepCounter,"Search and Fee Rule ","Rule added" ,"PASSED"
'					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'			wait 2
'		

	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If
	If bFlag = True Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Manual Fees"
	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_ManualFees", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewManualFee", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_FeeType", "RadioSelect", cust_data("feeType")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_FeeRule", "RadioSelect", cust_data("feeRule")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Quantity", "Set", cust_data("quantity")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_UnitPrice", "Set", cust_data("unitPrice")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ManualFeeText", "Set", cust_data("manualFeeText")
		Append_TestHTML StepCounter,"Create new manual Fee details","All details are maintained" ,"PASSED"
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		
		wait 4   ' waiting for the UI update
		Append_TestHTML StepCounter,"Manual Fee details ","Details are saved" ,"PASSED"
		wait 15
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebImage_Back", "Click", ""
'	
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_Yes", "Click", ""
'	End If
'		Append_TestHTML StepCounter,"Manual Fee details ","Details are saved" ,"PASSED"
		wait 4
		' This steps are needed as the update is not visible immediate. so navigating to some random page and back manual fees page
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewFeeRule", "Click", ""
'			wait 3
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchNewFeeRule", "Click", ""
'		wait 4
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Manual Fees"
'	wait 2
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_ManualFees", "Click", ""
'		wait 4
		' Now the changes will vbe visible on UI
		' Check if bank details updated, if yes click on the created acc num to assign debit mandates	
'		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_ManualFeesList")  Then
'			web_manual_fee = VerifyObjectExistInWebTbl("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","NoFrame", "WebTable_ManualFeesList", "WebElement", "innertext", cust_data("feeRule"))
'			'web_manual_fee = True
'			If  web_manual_fee = True Then
'				Append_TestHTML StepCounter,"Customer Manual Fee Rule added Successfully","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA" ,"PASSED"
'			else
'				Append_TestHTML StepCounter,"Customer Manual Fee Rule addition Failed","Expected Value: " & "NA" & VBCRLF & "Actual Value: " & "NA","FAILED"
'				bFlag = False
'			End If
'		End If
'	  	web_manual_fee = empty	
	End If
	
	wait 12   ' waiting for the UI update
	If customerERP_id <> "" Then
		query = "Select * from FeeRule  where FeeRuleDescription like '" &  cust_data("feeRule") & "';"
	'	set dbRecordSet_Fee_Item = execute_db_query(query_fee_item)
		Append_TestHTML StepCounter,"Get Fee Rule ID ",query,"PASSED"
	
		set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		
		db_Fee_RID = dbRecordSet("FeeRuleID")
		set dbRecordSet = Nothing
	
		query_fee_item = "select * from FeeItemUnbilled where PaymentCustomerID = (select customerid from Customer where CustomerERP = '" & cust_data("custERP") & "') and FeeRuleID = '" & db_Fee_RID &  "';"
	'	set dbRecordSet_Fee_Item = execute_db_query(query_fee_item)
		Append_TestHTML StepCounter,"FeeItemUnbilled table details ",query_fee_item,"PASSED"
	
		set dbRecordSet_Fee_Item = execute_db_query(query_fee_item, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		
		db_Fee_Type = dbRecordSet_Fee_Item("FeeTypeID")
		db_quantity = dbRecordSet_Fee_Item("Quantity")
		db_unit_price = dbRecordSet_Fee_Item("OriginalFeeUnitPrice")
		db_manual_fee_text = dbRecordSet_Fee_Item("FeeTextLine")
		set dbRecordSet_Fee_Item = Nothing
		
		If  db_Fee_Type <> ""  Then  ' Fee type 38 = Manual Fee
			Append_TestHTML StepCounter,"Fee Type is Manual (38) in database","Expected Value: " & "38" & VBCRLF & "Actual Value: " & db_Fee_Type ,"PASSED"
		else
			Append_TestHTML StepCounter,"Fee Type is NOT Manual (38) in database","Expected Value: " & "38" & VBCRLF & "Actual Value: " & db_Fee_Type ,"FAILED"
			bFlag = False
		End If
		
		If  db_quantity <> ""  Then
			Append_TestHTML StepCounter,"Quantity match in database","Expected Value: " & cust_data("quantity") & VBCRLF & "Actual Value: " & db_quantity ,"PASSED"
		else
			Append_TestHTML StepCounter,"Quantity mismatch in database","Expected Value: " & cust_data("quantity") & VBCRLF & "Actual Value: " & db_quantity ,"FAILED"
			bFlag = False
		End If
		
		If db_unit_price <> ""  Then
			Append_TestHTML StepCounter,"Unit Price match in database","Expected Value: " & cust_data("unitPrice") & VBCRLF & "Actual Value: " & db_unit_price ,"PASSED"
		else
			Append_TestHTML StepCounter,"Unit Price mismatch in database","Expected Value: " & cust_data("unitPrice") & VBCRLF & "Actual Value: " & db_unit_price ,"FAILED"
			bFlag = False
		End If
		
		If db_manual_fee_text  <> "" Then
			Append_TestHTML StepCounter,"Manual Fee text match in database","Expected Value: " & cust_data("manualFeeText") & VBCRLF & "Actual Value: " & db_manual_fee_text ,"PASSED"
		else
			Append_TestHTML StepCounter,"Manual Fee text mismatch in database","Expected Value: " & cust_data("manualFeeText") & VBCRLF & "Actual Value: " & db_manual_fee_text ,"FAILED"
			bFlag = False
		End If
	
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for verifying Manual Fee Creation details","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for for verifying Manual Fee Creation details details","FAILED"
		bRunFlag = False
	End If

End Function
	
	
'******************************* HEADER ******************************************
' Description : The function to create customer card types in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function CreateCustCardTypes(cust_erp)

	On error resume next
	
	bFlag = True
'	Call menuSelection("Info Subscriptions","Card Types")
	wait 1
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
'		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
'		bFlag = False
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Types"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardTypes", "Click", ""
	
	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCardType") Then
		Append_TestHTML StepCounter, "Card Types", "New Card Type screen navigatiion", "PASSED"
	else
		Append_TestHTML StepCounter, "Card Types", "Navigation to card types page failed", "FAILED"
		bFlag = False
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCardType", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebCheckBox_CardReissue") Then
		Append_TestHTML StepCounter, "Create New Card", "Navigated to New card screen", "PASSED"
	else
		Append_TestHTML StepCounter, "Create New Card", "Navigation to create card page failed", "FAILED"
		bFlag = False
	End If
	
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cards_WebLIST_CardType", "RadioSelect", "#6"
	
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_CardReissue", "Click", ""

	minCardLife = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_MinCardLife","GetROProperty","value")
	
	minReissuePeriod = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_MinReissuePeriod","GetROProperty","value")
	
	expiryPeriod = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_ExpiryPeriod","GetROProperty","value")
	
	embossingName = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossingName","GetROProperty","value")
	
	isActive = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_CTActive","GetROProperty","checked")
	
	If isActive = 1 or isActive = True Then
		isActive = True
	else
		isActive = False
	End If
	'msgbox "check cardtype selection"
	
	isReissue = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_Reissue","GetROProperty","checked")
	If isReissue = 1 or isReissue = True Then
		isReissue = True
	else
		isReissue = False
	End If
		Append_TestHTML StepCounter,"Customer card type details ","Maintained all input values" ,"PASSED"
	
	wait 5
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	'msgbox "check success"
	wait 3   ' waiting for the DB update
	
	query = "select * from CustomerCardType where CustomerID = (select customerid from Customer where CustomerERP = '" & cust_erp & "');"
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
' Description : The function to open GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function OpenApplication(url)
bSingleSignOnFlag = True
bSignOnFlag = False
	On error resume next
	If bSingleSignOnFlag = True and bSignOnFlag = False Then		
		systemutil.CloseProcessByName "msedge.exe"	
		wait 2		
		systemutil.Run "msedge.exe", url
		wait 4
		If Browser("name:=ntp\.msn\.com").Exist or (Window("micClass:=Browser","name:=Home Page").Exist = False) Then
			
			Browser("creationTime:=0").Highlight
			wait 2
			Browser("creationTime:=0").Maximize
			wait 2
			Browser("creationTime:=0").WinObject("abs_x:=4","abs_y:=104").Click
			wait 2
			Set WshShell = CreateObject("WScript.Shell")
			'Browser("name:=ntp\.msn\.com").HandleDialog micCancel
			'  		wshshell.SendKeys strData
	  		wait 1
	  		wshshell.SendKeys ("{TAB}")
	  		wait 2
	  		wshshell.SendKeys ("{TAB}")
	  		wait 2
	  		wshshell.SendKeys ("{TAB}")
	  		wait 2
	  		wshshell.SendKeys ("{ENTER}")
	  		wait 2
	  		Set WshShell = Nothing
	  	Else
'	  	msgbox "outside"
		End If
		wait 5
		Window("micClass:=Browser","name:=Home Page").Maximize
		wait 2
		flag = False
		rcount = Browser("name:=Home Page").Page("title:=Home Page").WebTable("html id:=MainContent_grdServers").Rowcount
		For i = 1 To rcount Step 1
			countryTab = Browser("name:=Home Page").Page("title:=Home Page").WebTable("html id:=MainContent_grdServers").GetCellData(i,1)
			If Instr(countryTab, countryName) > 0 Then
				Browser("name:=Home Page").Page("title:=Home Page").WebTable("html id:=MainContent_grdServers").childItem(i,1,"WebElement",0).click
				wait 4
				browserProp = Browser("creationTime:=1").GetROProperty("name")
				dynamicBrowser = "name:"&browserProp
				pageProp =  Browser("creationTime:=1").Page("creationTime:=1").GetROProperty("title")
				dynamicPage = "title:"& pageProp
				'Window("micClass:=Browser","name:="&browserProp).Maximize
				Browser("creationTime:=1").Maximize
				flag = True
				Exit For
			End If
		Next
		If flag = False Then
			Append_TestHTML StepCounter,"Country Name","Expected country name is not found in the application","FAILED"				
		End If
'		If objORDict.Exists("WebBrowser_PTShell_INDO") Then
'		  objORDict.item("WebBrowser_PTShell_INDO") = dynamicBrowser
'		End If
'		If objORDict.Exists("WebPage_PTShell_INDO") Then
'		  objORDict.item("WebPage_PTShell_INDO") = dynamicPage 
'		End If

		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","WebEdit_Search") Then		
			Append_TestHTML StepCounter,"Application launch","Application is launched successfully","PASSED"
			bSignOnFlag = True
		else
			Append_TestHTML StepCounter,"Application launch failed","Application launch failed","FAILED"
			bRunFlag = False
		End If 		
	End if 	
	
End Function


'******************************* HEADER ******************************************
' Description : The function to create customer card parameters in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function CreateCustCardParams(cust_erp)

	On error resume next
	
	bFlag = True
'	Call menuSelection("Card Types","Card Parameters")
	wait 1
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
'		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
'		bFlag = False
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	
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
		Append_TestHTML StepCounter,"Customer Card Parameter details ","Maintained all input values" ,"PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	wait 2
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","WebElement_CustCardParams") Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for create customer card parameters","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for create customer card parameters","FAILED"
		bRunFlag = False
	End If
	
	'DB Validation not required
	
End Function

'******************************* HEADER ******************************************
' Description : The function to create customer price rule in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function CreateCustPriceRule(cust_erp)

	On error resume next
	
	bFlag = True
'	Call menuSelection("Card Parameters","Payments")
'	wait 1
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
'		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
'	WebLink_Search
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Parameters"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardParams", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardParams1", "Click", ""
'	


	Call menuSelection("Card Parameters","Maintain Customer Price Rules")
	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ctl00_ctl11","html tag:=SPAN","innertext:=No","index:=1").Exist Then
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ctl00_ctl11","html tag:=SPAN","innertext:=No","index:=1").Click
		wait 2
	End If
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_PriceRule") Then
		Append_TestHTML StepCounter, "Create price rule", "Opened price rule page", "PASSED"
'		''msgbox "Price Rule Page"
		'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_PriceRule", "Click", ""
'			wait 2
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewPriceRule", "Click", ""
'			wait 2
'				Append_TestHTML StepCounter, "Create new price rule", "Navigating to new price rule page", "PASSED"
'			
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchPriceRule", "Click", ""
'			wait 2
'				Append_TestHTML StepCounter, "Search for existing price rule", "Clicked on search button", "PASSED"
'			
'		'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_PriceRuleDisc", "Click", ""
'		'
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_PriceRuleCat", "RadioSelect", "#1"
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'			wait 2
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_RetailDefault", "Click", ""
'		
'			wait 2
'				Append_TestHTML StepCounter, "Click on Price Rule Description", "Rule selected", "PASSED"
'			
'			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
'				Append_TestHTML StepCounter, "Create price rule", "Adding price rule", "PASSED"
'			else
'				Append_TestHTML StepCounter, "Create price rule", "Adding price rule failed", "FAILED"
'			End If
'			
'		'	refPrice = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RefPrice","GetROProperty","default value")
'			
'			refPrice = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_PriceRuleread","GetROProperty","default value")
'			
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'			
'			wait 3   ' waiting for the DB update
'			
	else
		Append_TestHTML StepCounter, "Create price rule", "Navigation to price rule page failed", "FAILED"
	End If

	query = "select * from CustomerPriceRule where CustomerID = (select customerid from Customer where CustomerERP = '" & cust_erp & "');"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	priceRuleID = dbRecordSet("PriceRuleID")
	
	query2 = "select * from PriceRule where PriceRuleID = '" & priceRuleID & "';"
	set dbRecordSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	priceRuleDesc = dbRecordSet2("PriceRuleDescription")
	refPrice = "Retail Default"
	If instr(priceRuleDesc, refPrice)<>0 Then
		Append_TestHTML StepCounter,"Price Rule Validation","Expected Value: " & priceRuleDesc & VBCRLF & "Actual Value: " & refPrice ,"PASSED"
	else
		Append_TestHTML StepCounter,"Price Rule Validation","Expected Value: " & priceRuleDesc & VBCRLF & "Actual Value: " & refPrice ,"FAILED"
		bRunFlag = False
	End If
	
End Function

Function tblcellclick(cvalue)
On error resume next
	
	bFlag = True
Set fetblobj =Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_ucCustomerFeeRuleMaintain_mvFeeRule_ucFeeRuleSearchControl_grdResults")
Set myObject = description.Create
myObject("micclass").value = "WebElement"
Set ccobj = fetblobj.ChildObjects(myObject)
''''msgbox ccobj.count
For i = 0 To  ccobj.count-1 Step 1
	tval =  ccobj(i).GetROProperty("innertext")
	If instr(tval,cvalue) > 0 Then
		ccobj(i).click
	End If
Next
End Function

'Function addCustFeeRule(fee_data)
'	
'	On error resume next
'	
'	bFlag = True
'	Call menuSelection("Maintain Customer Price Rules","Search for Fee Rules")
'	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
'		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
'		wait 2
'	End If
'
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_FeeRule") Then
'		Append_TestHTML StepCounter, "Create fee rule", "Opened fee rule page", "PASSED"
'	else
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Fee Rules"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules1", "Click", ""
'			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_FeeRule") Then
'				Append_TestHTML StepCounter, "Create fee rule", "Opened fee rule page", "PASSED"
'			End  IF
'	End If
'	wait 3
''	'''msgbox fee_data("items")
'	icount = fee_data("items")
'For itr = 1 To icount Step 1
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_FeeRule", "Click", ""
'		Append_TestHTML StepCounter, "Navigate Fee Rule", "Appeared fee rule screen", "PASSED"
'	wait 3
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewFeeRule", "Click", ""
'		Append_TestHTML StepCounter, "Create new Fee Rule", "new fee rule creation", "PASSED"
'	wait 3
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRule", "Click", ""
'		Append_TestHTML StepCounter, "Search for Fee Rule", "Clicked on search fee rule", "PASSED"
'	wait 3
''	'''msgbox fee_data("item"&itr)
'	If itr = 1  and appEnvName = "SPRINTQA" Then
'		itval = "AllFuel2%Rule"
'	Else
'		itval =  fee_data("item"&itr)
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebEdit_FeeRuleDesc", "Set", itval
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'		Append_TestHTML StepCounter, "Click on Search link", "To add fee rules", "PASSED"
'	wait 3
'	If itr = 1  and appEnvName = "SPRINTQA" Then
'		'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CardFeeDiesel", "Click", ""		'
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_AllFeerule", "Click", ""	
'		Else
''			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CardTransFeerule", "Click", ""	
''		webtable_clickcell "WebBrowser_PTShell_INDO","No Window", "WebPage_PTShell_INDO", "No Frame", "WebTable_CustFeeRuleList", "WebElement", 2,3
'			Call tblcellclick(itval)
'	End If
'		Append_TestHTML StepCounter, "Verify fee rule and save", "fee rule verified", "PASSED"
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
'		Append_TestHTML StepCounter, "Create fee rule", "Adding fee rule", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Create fee rule", "Adding price fee failed", "FAILED"
'	End If
'	
'	feeRuleDescGFN = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_FeeRule","GetROProperty","default value")
'	'''''msgbox feeRuleDescGFN
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'	wait 4
'	cust_erp = fee_data("custERP")
'	If cust_erp <> "" and feeRuleDescGFN <> "" Then
'	
'		query = "select * from CustomerFeeRule where CustomerID = (select customerid from Customer where CustomerERP = '" & cust_erp & "') order by 1 desc;"
'		set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		wait 2
'		feeRuleID = dbRecordSet("FeeRuleID")
'		custIDforRule = dbRecordSet("CustomerFeeRuleID")
'			
'		query1 = "update CustomerFeeRule set DateEffective='2021-01-10' where CustomerFeeRuleID = '" & custIDforRule & "';"
'		Append_TestHTML StepCounter,"Back date customer Fee Rule",query1,"PASSED"
'		set dbRecordSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		wait 2
'		set dbRecordSet2 = Nothing
'		
'		query2 = "select * from FeeRule where FeeRuleID = '" & feeRuleID & "';"
'		Append_TestHTML StepCounter,"Get Fee Rule Description",query2,"PASSED"
'		
'		set dbRecordSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
'		wait 2
'		feeRuleDesc = dbRecordSet2("FeeRuleDescription")
'		
'		If instr(trim(feeRuleDesc),trim( feeRuleDescGFN))<>0 Then
'			Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"PASSED"
'		else
'			Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"FAILED"
'			bFlag = False
'		End If
'	Else
'			Append_TestHTML StepCounter,"Fee RuleSelection and  Validation","Unable to back date of custmer Fee Rule" ,"FAILED"
'		bFlag = False
'	End If	
'Next	
'End Function


Function addCustFeeRule(fee_data)
	
	On error resume next
	
	bFlag = True
'	''msgbox "Search Fee Rules page navigation"
'	Call menuSelection("Maintain Customer Price Rules","Search for Fee Rules")
'	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
'		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
'		wait 2
'	End If
'
'	
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_FeeRule") Then
'		Append_TestHTML StepCounter, "Create fee rule", "Opened fee rule page", "PASSED"
'	else
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Fee Rules"
	wait 5
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules1", "Click", ""
	wait 4
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_FeeRule") Then
				Append_TestHTML StepCounter, "Create fee rule", "Opened fee rule page", "PASSED"
			End  IF
'	End If
	wait 3
	
'	''msgbox "Search Fee Rules table"
'	'''msgbox fee_data("items")
	icount = fee_data("items")
If icount >=1  Then
		For itr = 1 To icount Step 1
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_FeeRule", "Click", ""
			Append_TestHTML StepCounter, "Navigate Fee Rule", "Appeared fee rule screen", "PASSED"
		wait 3
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewFeeRule", "Click", ""
			Append_TestHTML StepCounter, "Create new Fee Rule", "new fee rule creation", "PASSED"
		wait 3
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRule", "Click", ""
			Append_TestHTML StepCounter, "Search for Fee Rule", "Clicked on search fee rule", "PASSED"
		wait 3
		
		If itr = 1  and ( appEnvName = "SPRINTQA" or appEnvName = "RELEASEQA" ) Then
	'		itval = "AllFuel2%Rule"
	'	Else
			itval =  fee_data("item"&itr)
	'		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_FeeRuleDesc") Then
	'			If instr(itval,"Card event Fee")>0 Then
	'				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_FeeWalvePastDue", "Click", ""
	'				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckBox_FeeWalvePastCancelled", "Click", ""
	'			ElseIf instr(itval,"Card Transaction Fee")>0 Then
	'				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_Feerulebasistype", "RadioSelect", "#2"	'"Percentage of Uplift"
	'			End If
	'		End  If
		End If
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebEdit_FeeRuleDesc", "Set", itval
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
			Append_TestHTML StepCounter, "Click on Search link", "To add fee rules", "PASSED"
		wait 3
		If itr = 1  and  ( appEnvName = "SPRINTQA" or appEnvName = "RELEASEQA" ) Then
			'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CardFeeDiesel", "Click", ""		'
	'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_AllFeerule", "Click", ""	
	'		Else
	'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CardTransFeerule", "Click", ""	
	'		webtable_clickcell "WebBrowser_PTShell_INDO","No Window", "WebPage_PTShell_INDO", "No Frame", "WebTable_CustFeeRuleList", "WebElement", 2,3
	'			If instr(itval,"Card Transaction Fee")>0 Then
	'				Call tblcellclick("Percentage of Uplift")
	'			Else
					Call tblcellclick(itval)
	'			End If
				
		End If
			Append_TestHTML StepCounter, "Verify fee rule and save", "fee rule verified", "PASSED"
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
			Append_TestHTML StepCounter, "Create fee rule", "Adding fee rule", "PASSED"
		else
			Append_TestHTML StepCounter, "Create fee rule", "Adding price fee failed", "FAILED"
		End If
		
		feeRuleDescGFN = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_FeeRule","GetROProperty","default value")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		wait 4
	'	feeRuleDescGFN = "Card Transaction Fee"
		
		cust_erp = fee_data("custERP")
		If cust_erp <> "" and feeRuleDescGFN <> "" Then
		
			query = "select * from CustomerFeeRule where CustomerID = (select customerid from Customer where CustomerERP = '" & cust_erp & "') order by 1 desc;"
			set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			feeRuleID = dbRecordSet("FeeRuleID")
			custIDforRule = dbRecordSet("CustomerFeeRuleID")
			
			query1 = "update CustomerFeeRule set DateEffective='2021-01-10' where CustomerFeeRuleID = '" & custIDforRule & "';"
			Append_TestHTML StepCounter,"Back date customer Fee Rule",query1,"PASSED"
			set dbRecordSet2 = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			set dbRecordSet2 = Nothing
			
			query2 = "select * from FeeRule where FeeRuleID = '" & feeRuleID & "';"
			Append_TestHTML StepCounter,"Get Fee Rule Description",query2,"PASSED"
			
			set dbRecordSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
			feeRuleDesc = dbRecordSet2("FeeRuleDescription")
			
			If instr(trim(feeRuleDesc),trim( feeRuleDescGFN))<>0 Then
				Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"PASSED"
			else
				Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"FAILED"
				bFlag = False
			End If
			set dbRecordSet2 = Nothing
			set dbRecordSet = Nothing
		Else
				Append_TestHTML StepCounter,"Fee RuleSelection and  Validation","Unable to back date of custmer Fee Rule" ,"FAILED"
			bFlag = False
		End If	
	Next	
End If

End Function




'
'******************************* HEADER ******************************************
' Description : The function to create customer fee rule in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function CreateCustFeeRule(cust_erp)

	On error resume next
	
	bFlag = True
	
	wait 1
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
'		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Parameters"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardParams", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardParams1", "Click", ""



'	Call menuSelection("Maintain Customer Price Rules","Search for Fee Rules")
'
'	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
'		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
'		wait 2
'	End If
'
	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_FeeRule") Then
'		Append_TestHTML StepCounter, "Create fee rule", "Opened fee rule page", "PASSED"
'	else
'	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Fee Rules"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRules1", "Click", ""
	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If

			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_FeeRule") Then
				Append_TestHTML StepCounter, "Create fee rule", "Opened fee rule page", "PASSED"
			End  IF
'	End If
	wait 3
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_FeeRule", "Click", ""
		Append_TestHTML StepCounter, "Navigate Fee Rule", "Appeared fee rule screen", "PASSED"
	wait 3
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewFeeRule", "Click", ""
		Append_TestHTML StepCounter, "Create new Fee Rule", "new fee rule creation", "PASSED"
	wait 3
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchFeeRule", "Click", ""
		Append_TestHTML StepCounter, "Search for Fee Rule", "Clicked on search fee rule", "PASSED"
	wait 3
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
		Append_TestHTML StepCounter, "Click on Search link", "To add fee rules", "PASSED"
	If appEnvName = "SPRINTQA" Then
		'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CardFeeDiesel", "Click", ""		'WebEdit_FeeRuleDesc
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_AllFeerule", "Click", ""	
'		Else
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_CardTransFeerule", "Click", ""	
			Append_TestHTML StepCounter, "Verify fee rule and save", "fee rule verified", "PASSED"
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
			Append_TestHTML StepCounter, "Create fee rule", "Adding fee rule", "PASSED"
		else
			Append_TestHTML StepCounter, "Create fee rule", "Adding price fee failed", "FAILED"
		End If
		
		feeRuleDescGFN = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_FeeRule","GetROProperty","default value")
		'''''msgbox feeRuleDescGFN
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		
		wait 3   ' waiting for the DB update
	End If
	feeRuleDescGFN = "Card Transaction Fee"
If cust_erp <> "" and feeRuleDescGFN <> "" Then

	query = "select * from CustomerFeeRule where CustomerID = (select customerid from Customer where CustomerERP = '" & cust_erp & "') order by 1 desc;"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	feeRuleID = dbRecordSet("FeeRuleID")
	custIDforRule = dbRecordSet("CustomerID")
		
	query1 = "update CustomerFeeRule set DateEffective='2021-01-10' where CustomerID = '" & custIDforRule & "';"
	Append_TestHTML StepCounter,"Back date customer Fee Rule",query1,"PASSED"
	set dbRecordSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	set dbRecordSet2 = Nothing
	
	query2 = "select * from FeeRule where FeeRuleID = '" & feeRuleID & "';"
	Append_TestHTML StepCounter,"Get Fee Rule Description",query2,"PASSED"
	
	set dbRecordSet2 = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	feeRuleDesc = dbRecordSet2("FeeRuleDescription")
	
	If instr(trim(feeRuleDesc),trim( feeRuleDescGFN))<>0 Then
		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"PASSED"
	else
		Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"FAILED"
		bFlag = False
	End If
Else
		Append_TestHTML StepCounter,"Fee RuleSelection and  Validation","Unable to back date of custmer Fee Rule" ,"FAILED"
	bFlag = False
End If	
	
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
Public Function CreateCard(cust_erp, cust_vrn, card_type, emboss_type)

	On error resume next
	
	bFlag = True
	
	wait 5
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
'		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustSummary", "Click", ""
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Card Maintenance"
	wait 5
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CardMaintenance", "Click", ""
	wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCard") Then
		Append_TestHTML StepCounter, "Create card", "Navigated to create card page", "PASSED"
	Else
'		Append_TestHTML StepCounter, "Create card", "Navigation to create card page failed", "FAILED"
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCard", "Click", ""
	
'	WebList_CardType
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_CardTypeIDVal") Then
		Append_TestHTML StepCounter, "Create card", "Navigated to create new card page", "PASSED"
	else
'		Append_TestHTML StepCounter, "Create card", "Navigation to create card page failed", "FAILED"
	End If
'	''msgbox "Card Type Selection"
	wait 5
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_CardType", "RadioSelect", card_type
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_EmbossType", "RadioSelect", emboss_type
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_VRN", "Set", cust_vrn
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
		Append_TestHTML StepCounter, "Create card", "Added details on create new card page", "PASSED"
	else
		Append_TestHTML StepCounter, "Create card", "Adding details on create card page failed", "FAILED"
	End If
	
	typeOfPinVal = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CardstypeofPIN_WebList","GetROProperty","selection")
	typeOfPinVisibility = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CardstypeofPIN_WebList","GetROProperty","disabled")
				
	If ucase(typeOfPinVal) = ucase("Card PIN") and typeOfPinVisibility = "1" Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebList_PinadvType", "RadioSelect", "Paper"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "CardstypeofPINSelection_WebList", "RadioSelect", "System generated"
	End If			
	pinAdviceTypeVal = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebList_PinadvType","GetROProperty","selection")
	If ucase(pinAdviceTypeVal) <> ucase("None") Then
		Call addAddressEmailofPinadvanceType()
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		wait 5
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Cardsconf_WebElement_Save", "Click", ""	
	wait 5
	cardNum = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Card","GetROProperty","default value")
	
	expDate_actual = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Expiry","GetROProperty","default value")
	expDate_arr = split(expDate_actual,"/")
	expDate = expDate_arr(2) & "-" & expDate_arr(1) & "-" & expDate_arr(0)
	
	embossText = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EmbossText","GetROProperty","default value")	
	
	wait 3   ' waiting for the DB update
	
		query = "Select * from CardPAN where PAN="& cardNum &" ;"
'		Append_TestHTML StepCounter, "Execute Query", query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		
		db_CardPANIDi=  dbRecordSet("CardPANID")
		XML_Cardpanid = db_CardPANIDi				
		Set query = Nothing
		Set dbRecordSet = Nothing
		
	query = "select * from Card where PAN = '" & cardNum & "';"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	cardPAN = dbRecordSet("PAN")
	cardCardID= dbRecordSet("CardID")
	cardPANNum = cardPAN
	dbExpDate =  dbRecordSet("ExpiryDate")
	dbEmbossNum =  dbRecordSet("EmbossRegNumber")
	dbEmbossText =  dbRecordSet("EmbossText")
	cardPAN_no = cardPAN
	cardPANNum = cardPAN
	cardExpiry_date = dbExpDate
	XML_Cadrid = cardCardID
	If instr(cardPAN, cardNum)<>0 Then
		Append_TestHTML StepCounter,"Card PAN number Validation","Expected Value: " & cardPAN & VBCRLF & "Actual Value: " & cardNum ,"PASSED"
	else
		Append_TestHTML StepCounter,"Card PAN number Validation","Expected Value: " & cardPAN & VBCRLF & "Actual Value: " & cardNum ,"FAILED"
		bRunFlag = False
	End If
	
	If instr(expDate, dbExpDate)<>0 Then
		Append_TestHTML StepCounter,"ExpiryDate Validation","Expected Value: " & expDate & VBCRLF & "Actual Value: " & dbExpDate ,"PASSED"
	else
		Append_TestHTML StepCounter,"ExpiryDate Validation","Expected Value: " & expDate & VBCRLF & "Actual Value: " & dbExpDate ,"FAILED"
		bRunFlag = False
	End If
	
	If instr(cust_vrn, dbEmbossNum)<>0 Then
		Append_TestHTML StepCounter,"EmbossRegNumber Validation","Expected Value: " & cust_vrn & VBCRLF & "Actual Value: " & dbEmbossNum ,"PASSED"
	else
		Append_TestHTML StepCounter,"EmbossRegNumber Validation","Expected Value: " & cust_vrn & VBCRLF & "Actual Value: " & dbEmbossNum ,"FAILED"
		bRunFlag = False
	End If
	
	If instr(embossText, dbEmbossText)<>0 Then
		Append_TestHTML StepCounter,"EmbossText Validation","Expected Value: " & embossText & VBCRLF & "Actual Value: " & dbEmbossText ,"PASSED"
	else
		Append_TestHTML StepCounter,"EmbossText Validation","Expected Value: " & embossText & VBCRLF & "Actual Value: " & dbEmbossText ,"FAILED"
		bRunFlag = False
	End If
	
End Function


'******************************* HEADER ******************************************
' Description : The function to create sub customer in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function CreateSubCust(cust_erp, cust_name)

	On error resume next
	
	bFlag = True
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
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
	wait 1
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
	else
		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustSummary", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustDetails", "Click", ""
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewSubAccnt", "Click", ""
	If appEnvName = "RD" Then
		Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("New Sub Account").Click
	Else
		Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("New Sub Account").Click
	End If
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_No", "Click", ""
	End If

	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_SubFullName") Then
		Append_TestHTML StepCounter, "Create sub Customer", "Opened create sub customer page", "PASSED"
	else
		Append_TestHTML StepCounter, "Create sub Customer", "Navigation to create sub customer page failed", "FAILED"
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SubFullName", "Set", cust_name & registrationnumber1
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SubShortName", "Set", cust_name & registrationnumber1
		
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SubTradeName", "Set", cust_name & registrationnumber1
	
	cust_reg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RegNum", "GetROProperty", "value")
	
	cust_reg2 = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Reg2Num", "GetROProperty", "value")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RegNum", "Set", registrationnumber1 'cust_reg+"0"
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Reg2Num", "Set",  registrationnumber 'cust_reg2+"0"
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
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

	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_No", "Click", ""
	End If
	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") Then
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
'	End If
	
	cust_erp_sub = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_CustERP", "GetROProperty", "value")
	'cust_erp_sub = "ID00000109"
	customerERP_sub_id = cust_erp_sub
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CustERP") Then
		Append_TestHTML StepCounter, "Save sub Customer details", "Saved sub customer details and ERP generated: "&cust_erp_sub, "PASSED"
	else
		Append_TestHTML StepCounter, "Save sub Customer details", "Saving sub customer details failed", "FAILED"
	End If
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
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Customer Structure"
'	wait 2
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustomerStructure", "Click", ""
	wait 2
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_CustMain") Then
		Append_TestHTML StepCounter, "Open customer hierarchy", "Opened customer hierarchy page", "PASSED"
	else
'	Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Customer Structure").Click
		If appEnvName = "RD"  Then
			Browser("Pilipinas Shell Petroleum").Page("S.A. Belgian Shell N.V.").Link("Customer Structure").Click
		Else
			Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Customer Structure").Click
		End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebELement_ConfirmDialog")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_No", "Click", ""
	End If
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Customer Structure"
'		
'		wait 2
'		Browser("Pilipinas Shell Petroleum").Page("Pilipinas Shell Petroleum").Link("Customer Structure_2").Click
'		wait 5
'		Append_TestHTML StepCounter, "Open customer hierarchy", "Navigation to customer hierarchy page failed", "FAILED"
	End If
	
	cust_main = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustMain", "GetROProperty", "innertext")
	
	cust_sub = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustSub", "GetROProperty", "innertext")
	
	If instr(cust_main, cust_erp) <> 0 Then
		Append_TestHTML StepCounter, "Validate main customer", "Validated main customer successfully. Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & cust_main, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate main custome", "Validation of main customer failed. Expected Value: " & cust_erp & VBCRLF & "Actual Value: " & cust_main, "FAILED"
	End If
	
	If instr(cust_sub, cust_erp_sub) <> 0 Then
		Append_TestHTML StepCounter, "Validate sub customer", "Validated sub customer successfully. Expected Value: " & cust_erp_sub & VBCRLF & "Actual Value: " & cust_sub, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate sub custome", "Validation of sub customer failed. Expected Value: " & cust_erp_sub & VBCRLF & "Actual Value: " & cust_sub, "FAILED"
		bRunFlag = False
	End If
	
	'No DB validation mentioned
	
	'query = "select * from CustomerAddress where CustomerId = (select customerid from Customer where CustomerERP = '" & cust_erp & "');"
	
	'set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	'wait 2
	'db_cust_id = dbRecordSet("CustomerERP")
	
	
End Function


'******************************* HEADER ******************************************
' Description : The function to create customer info subscription in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function CreateCustInfoSub(cust_erp, op_type, contact, altDistMethod)

	On error resume next
	
	bFlag = True
	Call menuSelection("Payments","Info Subscriptions")
	wait 1
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
'	wait 1
'	
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLIST_Role") Then
'		Append_TestHTML StepCounter, "Search Customer", "Opened search customer page", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Search Customer", "Navigation to search customer page failed", "FAILED"
'	End If
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
'
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_erp
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustSummary", "Click", ""
'	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustInfoSub", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_InfoSub") Then
		Append_TestHTML StepCounter, "Customer Info Subscription", "Navigated to customer info subscription page", "PASSED"
	else
		Append_TestHTML StepCounter, "Customer Info Subscription", "Navigation to customer info subscription page failed", "FAILED"
	End If
	
		
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_InfoSub", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewInfoSub", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Actions", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewSubDetail", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebList_OutputType") Then
		Append_TestHTML StepCounter, "Customer Info Subscription", "Navigated to create new customer info subscription page", "PASSED"
	else
		Append_TestHTML StepCounter, "Customer Info Subscription", "Navigation to create new customer info subscription page failed", "FAILED"
	End If
	
	
	dist_method = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_DistMethod", "GetROProperty", "default value")
	Print op_type
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_OutputType", "RadioSelect", op_type
	Print contact
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_Contact", "RadioSelect",  contact
	Print altDistMethod
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_AltDistMethod", "RadioSelect", altDistMethod
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebCheckbox_isPrimary", "Click", ""
		Append_TestHTML StepCounter, "Customer Info Subscription", "Maintainatined all input values", "PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
		Append_TestHTML StepCounter, "Customer Info Subscription", "Created new customer info subscription", "PASSED"
	else
		Append_TestHTML StepCounter, "Customer Info Subscription", "Creating new customer info subscription failed", "FAILED"
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	
	query = "select * from CustomerInfoSubscription where CustomerId = (select customerid from Customer where CustomerERP = '" & cust_erp & "');"
		Append_TestHTML StepCounter, "CustomerInfoSubscription", query, "PASSED"
	
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_infosub_id = dbRecordSet("InfoSubscriptionID")
	
	query1 = "Select * from InfoSubscription Where InfoSubscriptionID = '" & db_infosub_id & "'"
		Append_TestHTML StepCounter, "InfoSubscription", query1, "PASSED"
	
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_infoprovider_id = dbRecordSet("InfoProviderID")
	
	
	query2 = "Select * from InfoSubscriptionDetail Where InfoSubscriptionID = '" & db_infosub_id & "'"
		Append_TestHTML StepCounter, "InfoSubscriptionDetail", query2, "PASSED"
	
	set dbRecordSet = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_distMethod_id = dbRecordSet("DistributionMethodID")	'1
	db_altDistMethod_id = dbRecordSet("AlternativeDistributionMethodID")	'5
	db_outputType_id = dbRecordSet("OutputTypeID")
	
	query3 = "Select * from LocalisedDescriptions where Culture = 'en-GB' and ColumnName like 'DistributionMethodID' and Value='" & db_distMethod_id & "'"
		Append_TestHTML StepCounter, "LocalisedDescriptions", query3, "PASSED"
	
	set dbRecordSet = execute_db_query(query3, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_distMethod = dbRecordSet("Description")
	
	query4 = "Select * from LocalisedDescriptions where Culture = 'en-GB' and ColumnName like 'DistributionMethodID' and Value='" & db_altDistMethod_id & "'"
		Append_TestHTML StepCounter, "LocalisedDescriptions", query4, "PASSED"
	
	set dbRecordSet = execute_db_query(query4, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_altDistMethod = dbRecordSet("Description")
	
	query5 = "Select * from LocalisedDescriptions where Culture = 'en-GB' and ColumnName like 'OutputTypeID' and Value='" & db_outputType_id & "'"
	set dbRecordSet = execute_db_query(query5, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_outputType = dbRecordSet("Description")
	
	If instr(dist_method, db_distMethod) <> 0 Then
		Append_TestHTML StepCounter, "Validate sub customer", "Validated sub customer successfully. Expected Value: " & dist_method & VBCRLF & "Actual Value: " & db_distMethod, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate sub custome", "Validation of sub customer failed. Expected Value: " & dist_method & VBCRLF & "Actual Value: " & db_distMethod, "FAILED"
		bRunFlag = False
	End If
	
	If instr(altDistMethod, db_altDistMethod) <> 0 Then
		Append_TestHTML StepCounter, "Validate sub customer", "Validated sub customer successfully. Expected Value: " & altDistMethod & VBCRLF & "Actual Value: " & db_altDistMethod, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate sub custome", "Validation of sub customer failed. Expected Value: " & altDistMethod & VBCRLF & "Actual Value: " & db_altDistMethod, "FAILED"
		bRunFlag = False
	End If
	
	If instr(op_type, db_outputType) <> 0 Then
		Append_TestHTML StepCounter, "Validate sub customer", "Validated sub customer successfully. Expected Value: " & op_type & VBCRLF & "Actual Value: " & db_outputType, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate sub custome", "Validation of sub customer failed. Expected Value: " & op_type & VBCRLF & "Actual Value: " & db_outputType, "FAILED"
		bRunFlag = False
	End If
	
End Function


'******************************* HEADER ******************************************
' Description : The function to back date newly created card in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function BackDateCardDetails(cust_erp, PAN_num)
'''''msgbox cust_erp
	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
	End If
	
	'Get customer ID
	Set dbRecordSet = execute_db_query("Select * from Customer where CustomerERP = '" & cust_erp & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Customer ID", "Select * from Customer where CustomerERP = '" & cust_erp & "'", "PASSED"
	wait 2
	db_cust_id = dbRecordSet("CustomerID")
	Set dbRecordSet = Nothing
	If db_cust_id <> "" Then
		Append_TestHTML StepCounter, "Customer ID", "Customer ID: " & db_cust_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Customer ID", "DB values not found. Customer ID: " & db_cust_id, "FAILED"
	End If
	
	'Get PAN ID
	Set dbRecordSet = execute_db_query("Select * from Card where PAN = '" & PAN_num & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Card ID", "Select * from Card where PAN = '" & PAN_num & "'", "PASSED"
	wait 2
	db_card_id = dbRecordSet("CardID")
	Set dbRecordSet = Nothing
	If db_card_id <> "" Then
		Append_TestHTML StepCounter, "Card ID", "Card ID: " & db_card_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Card ID", "DB values not found. Card ID: " & db_card_id, "FAILED"
	End If
	
	'Get actual date established and start date
	query = "Select * from Customer where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Get Dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_established_date = dbRecordSet("DateEstablished")
	db_start_date = dbReco
	rdSet("StartDate")
	
	est_date_arr = split(db_established_date," ")	
	start_date_arr = split(db_start_date," ")
	
	new_est_date = "2021-01-10 " & est_date_arr(1) 
	new_start_date = "2021-01-10 " & start_date_arr(1)
	Set dbRecordSet = Nothing
	If db_established_date <> "" and db_start_date <> "" Then
		Append_TestHTML StepCounter, "Get Dates", "DateEstablished: " & db_established_date & VBCRLF & "StartDate: " & db_start_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Get Dates", "DB values not found. DateEstablished: " & db_established_date & VBCRLF & "StartDate: " & db_start_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update Customer Set DateEstablished = '" & new_est_date & "', StartDate = '" & new_start_date & "' where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Update dates", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get updated date established and start date
	query2 = "Select * from Customer where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Get Dates", query2, "PASSED"
	set dbRecordSet = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_updated_established_date = dbRecordSet("DateEstablished")
	db_updated_start_date = dbRecordSet("StartDate")
	Set dbRecordSet = Nothing
	If db_updated_established_date <> "" and db_updated_start_date <> "" Then
		Append_TestHTML StepCounter, "Get Dates", "Updated DateEstablished: " & db_updated_established_date & VBCRLF & "Updated StartDate: " & db_updated_start_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Get Dates", "DB values not found. Updated DateEstablished: " & db_updated_established_date & VBCRLF & "Updated StartDate: " & db_updated_start_date, "FAILED"
	End If
	
	timestamp_arr = split(db_updated_established_date," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	db_updated_established_date = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If db_updated_established_date =new_est_date  Then
		Append_TestHTML StepCounter, "Established date", "Updated established date successfully. Expected Value: " & new_est_date & VBCRLF & "Actual Value: " & db_updated_established_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Established date", "Update of established date failed. Expected Value: " & new_est_date & VBCRLF & "Actual Value: " & db_updated_established_date, "FAILED"
	End If
	
	timestamp_arr = split(db_updated_start_date," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	db_updated_start_date = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If db_updated_start_date =new_start_date  Then
		Append_TestHTML StepCounter, "Start date", "Updated start date successfully. Expected Value: " & new_start_date & VBCRLF & "Actual Value: " & db_updated_start_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Start date", "Update of start date failed. Expected Value: " & new_start_date & VBCRLF & "Actual Value: " & db_updated_start_date, "FAILED"
		bRunFlag = False
	End If
	
	'---------------------------------------------------------------------------------------
	'Get Info Subscription ID
	Set dbRecordSet = execute_db_query("Select * from CustomerInfoSubscription where CustomerID = '" & db_cust_id & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Info Subscription ID", "Select * from CustomerInfoSubscription where CustomerID = '" & db_cust_id & "'", "PASSED"
	wait 2
	db_infosub_id = dbRecordSet("InfoSubscriptionID")
	Set dbRecordSet = Nothing
	If db_infosub_id <> "" Then
		Append_TestHTML StepCounter, "Info Subscription ID", "Info Subscription ID: " & db_infosub_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Info Subscription ID", "DB values not found. Info Subscription ID: " & db_infosub_id, "FAILED"
	End If
	
	'Get actual date established and start date
	query = "Select * from InfoSubscription where InfoSubscriptionID = '" & db_infosub_id & "'"
	Append_TestHTML StepCounter, "Info Subscription", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update InfoSubscription Set DateEffective = '" & new_effective_date & "' Where InfoSubscriptionID = '" & db_infosub_id & "'"
	Append_TestHTML StepCounter, "Update DateEffective", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from InfoSubscription where InfoSubscriptionID = '" & db_infosub_id & "'"
	Append_TestHTML StepCounter, "Verify updated dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If

	'-----------------------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from CustomerPriceRule where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Get Customer PriceRule", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CustomerPriceRule Set DateEffective = '" & new_effective_date & "' Where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Update DateEffective", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from CustomerPriceRule where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Verify updated dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If
'------------------------------------------------------------------------------------------------------------------------------


	'Get actual date established and start date
	query = "Select * from CustomerCard where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Customer Card", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CustomerCard Set DateEffective = '" & new_effective_date & "' Where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Update Customer Card", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from CustomerCard where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Validate Customer Card", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If
	'-----------------------------------------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from Card where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Get Card table", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_date = dbRecordSet("InitialCreationDateTime")
	date_arr = split(db_date," ")	
	new_date = "2021-01-10 " & date_arr(1) 
	Set dbRecordSet = Nothing
	If db_date <> "" Then
		Append_TestHTML StepCounter, "InitialCreationDateTime", "InitialCreationDateTime: " & db_date, "PASSED"
	else
		Append_TestHTML StepCounter, "InitialCreationDateTime", "DB values not found. InitialCreationDateTime: " & db_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update Card Set InitialCreationDateTime = '" & new_date & "' Where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Update Card table", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from Card where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Validate card table", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_updated_date = dbRecordSet("InitialCreationDateTime")
	Set dbRecordSet = Nothing
	If db_updated_date <> "" Then
		Append_TestHTML StepCounter, "Validate updated date", "InitialCreationDateTime: " & db_updated_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "DB values not found. InitialCreationDateTime: " & db_updated_date, "FAILED"
	End If
	
	timestamp_arr = split(db_updated_date," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	db_updated_date = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If db_updated_date =new_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated date successfully. Expected Value: " & new_date & VBCRLF & "Actual Value: " & db_updated_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating date failed. Expected Value: " & new_date & VBCRLF & "Actual Value: " & db_updated_date, "FAILED"
		bRunFlag = False
	End If
	'-----------------------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from CustomerFeeRule where CustomerID = '" & db_cust_id & "';"		' and FeeRuleID = 21"
	Append_TestHTML StepCounter, "Get CustomerFeeRule", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	db_customer_feerule = dbRecordSet("CustomerFeeRuleID")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	'query1 = "Update CustomerFeeRule Set DateEffective = '" & new_effective_date & "' Where CustomerFeeRuleID = '" & db_customer_feerule & "'"
	query1 = "Update CustomerFeeRule Set DateEffective = '" & new_effective_date & "' Where CustomerID = '" & db_cust_id & "';"
	Append_TestHTML StepCounter, "Update CustomerFeeRule", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from CustomerFeeRule where CustomerID = '" & db_cust_id & "';"    'and FeeRuleID = 21"
	Append_TestHTML StepCounter, "Verify updated dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If
	'-------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from CardStatusHistory where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Get CardStatusHistory", query, "PASSED"
	'''''msgbox query
'	and ModifiedBy = 'EFN'"
	set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_date = dbRecordSet("DateModified")
	db_date2 = dbRecordSet("DateModifiedNextStatus")
	'''''msgbox db_date
	date_rows = split(db_date,"|")
	date_arr1 = split(date_rows(0)," ")	
	date_arr2 = split(date_rows(1)," ")
	actual_date1 = 	date_rows(0)
	actual_date2 = 	date_rows(1)
	new_date1 = "2021-01-10 " & date_arr1(1) 
	new_date2 = "2021-01-10 " & date_arr2(1) 
	
	date_rows2 = split(db_date2,"|")
	date_arr3 = split(date_rows2(0)," ")
	new_date3 = "2021-01-10 " & date_arr3(1) 
	actual_date3 = 	date_rows2(0)
	
	cardStatusID = split(dbRecordSet("CardStatusHistoryID"),"|")
	cardStatusID1 = cardStatusID(0)
	cardStatusID2 = cardStatusID(1)
	Set dbRecordSet = Nothing
	If actual_date1 <> "" and  actual_date2 <> "" and  actual_date3 <> "" Then
		Append_TestHTML StepCounter, "DateModified & DateModifiedNextStatus", "DateModified1: " & actual_date1 & VBCRLF & "DateModified2: " & actual_date2 & VBCRLF & "DateModifiedNextStatus: " & actual_date3 , "PASSED"
	else
		Append_TestHTML StepCounter, "DateModified & DateModifiedNextStatus", "DB values not found. DateModified1: " & actual_date1 & VBCRLF & "DateModified2: " & actual_date2 & VBCRLF & "DateModifiedNextStatus: " & actual_date3, "FAILED"
	End If
	If cardStatusID1 <> "" and  cardStatusID2 <> "" Then
		Append_TestHTML StepCounter, "CardStatusHistoryID", "CardStatusHistoryID1: " & cardStatusID1 & VBCRLF & "CardStatusHistoryID2: " & cardStatusID2, "PASSED"
	else
		Append_TestHTML StepCounter, "CardStatusHistoryID", "DB values not found. CardStatusHistoryID1: " & cardStatusID1 & VBCRLF & "CardStatusHistoryID2: " & cardStatusID2, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CardStatusHistory Set DateModified = '" & new_date1 & "', DateModifiedNextStatus = '" & new_date3 & "' Where CardStatusHistoryID = '" & cardStatusID1 & "'"
	Append_TestHTML StepCounter, "Update CardStatusHistory", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	query2 = "Update CardStatusHistory Set DateModified = '" & new_date2 & "' Where CardStatusHistoryID = '" & cardStatusID2 & "'"
	Append_TestHTML StepCounter, "Update CardStatusHistory", query2, "PASSED"
	set dbRecordSet = execute_db_query(query2, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	
	'Get actual date established and start date
	query = "Select * from CardStatusHistory where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Validate CardStatusHistory", query, "PASSED"
'	and ModifiedBy = 'EFN'"
	set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_date = dbRecordSet("DateModified")
	db_date2 = dbRecordSet("DateModifiedNextStatus")
	date_rows = split(db_date,"|")
	
	date_rows2 = split(db_date2,"|")
	updated_date1 = date_rows(0)
	updated_date2 = date_rows(1)
	updated_date3 = date_rows2(0)
	Set dbRecordSet = Nothing
	If updated_date1 <> "" and  updated_date2 <> "" and  updated_date3 <> "" Then
		Append_TestHTML StepCounter, "Updated DateModified & DateModifiedNextStatus", "DateModified1: " & updated_date1 & VBCRLF & "DateModified2: " & updated_date2 & VBCRLF & "DateModifiedNextStatus: " & updated_date3 , "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateModified & DateModifiedNextStatus", "DB values not found. DateModified1: " & updated_date1 & VBCRLF & "DateModified2: " & updated_date2 & VBCRLF & "DateModifiedNextStatus: " & updated_date3, "FAILED"
	End If
	
	timestamp_arr = split(updated_date1," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	updated_date1 = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If updated_date1 =new_date1  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated DateModified1 successfully. Expected Value: " & new_date1 & VBCRLF & "Actual Value: " & updated_date1, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating DateModified1 failed. Expected Value: " & new_date1 & VBCRLF & "Actual Value: " & updated_date1, "FAILED"
		bRunFlag = False
	End If
	
	timestamp_arr = split(updated_date2," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	updated_date2 = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If updated_date2 =new_date2  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated DateModified2 successfully. Expected Value: " & new_date2 & VBCRLF & "Actual Value: " & updated_date2, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating DateModified2 failed. Expected Value: " & new_date2 & VBCRLF & "Actual Value: " & updated_date2, "FAILED"
		bRunFlag = False
	End If
	
	timestamp_arr = split(updated_date3," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	updated_date3 = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If updated_date3 =new_date3  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated DateModifiedNextStatus successfully. Expected Value: " & new_date3 & VBCRLF & "Actual Value: " & updated_date3, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating DateModifiedNextStatus failed. Expected Value: " & new_date3 & VBCRLF & "Actual Value: " & updated_date3, "FAILED"
		bRunFlag = False
	End If
	
''''msgbox "Back dated both card & Customer"	
End Function



'******************************* HEADER ******************************************
' Description : The function to verify file watcher job in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function VerifyFileWatcherJob(cust_erp, PAN_num, file_name)

	On error resume next
	
	bFlag = True
	
	If file_name = "" or file_name = empty Then
		file_name = dx026FinalFName
		fileName = dx026FinalFName 'Added fileName variable to global params, should assign file name of DX026 file to this variable in that function 
	End If
	wait 10
If PAN_num <> "" and cust_erp <> ""  Then
	
	Set dbRecordSet = execute_db_query("Select * from Job where JobTypeID = 191 Order by 1 desc;", 1, "SFN_SHELL_SPRINTQA_IMPORT")
	Append_TestHTML StepCounter, "Run Job 191", "Select * from Job were JobTypeID = 191 Order by 1 desc", "PASSED"
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_status_id <> "" and db_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 191", "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 191", "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "FAILED"
		bRunFlag = False
	End If
	
	
	Set dbRecordSet = execute_db_query("Select * from JobLog where JobID = '" & db_id & "' Order by 1 desc;", 2, "SFN_SHELL_SPRINTQA_IMPORT")
	Append_TestHTML StepCounter, "Check status of Job 191", "Select * from JobLog where JobID = '" & db_id & "' Order by 1 desc", "PASSED"
	wait 2
	db_msg = dbRecordSet("Message")
	msg_arr = split(db_msg,"|")
	process_msg = msg_arr(0)
	receive_msg = msg_arr(1)
	Set dbRecordSet = Nothing
	If process_msg <> "" and receive_msg <> "" Then
		Append_TestHTML StepCounter, "Get status from Job 191", "Status Message1: " & process_msg & VBCRLF & "Status Message2: " & receive_msg, "PASSED"
	else
		Append_TestHTML StepCounter, "Get status from Job 191", "DB values not found. Status Message1: " & process_msg & VBCRLF & "Status Message2: " & receive_msg, "FAILED"
		bRunFlag = False
	End If
	
	
	Set dbRecordSet = execute_db_query("Select * from Job where JobTypeID = 257 Order by 1 desc;", 1, "SFN_SHELL_SPRINTQA_IMPORT")
	Append_TestHTML StepCounter, "Run Job 257", "Select * from Job where JobTypeID = 257 Order by 1 desc", "PASSED"
	wait 2
	db_id2 = dbRecordSet("ID")
	db_status_id2 = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_id2 <> "" and db_status_id2 <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 257", "ID: " & db_id2 & VBCRLF & "Status ID: " & db_status_id2, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 257", "DB values not found. ID: " & db_id2 & VBCRLF & "Status ID: " & db_status_id2, "FAILED"
		bRunFlag = False
	End If
	
	
	Set dbRecordSet = execute_db_query("Select * from JobLog where JobID = '" & db_id2 & "' Order by 1 desc;", 4, "SFN_SHELL_SPRINTQA_IMPORT")
	Append_TestHTML StepCounter, "Check status of Job 257", "Select * from JobLog where JobID = '" & db_id2 & "' Order by 1 desc", "PASSED"
	wait 2
	db_msg2 = dbRecordSet("Message")
	msg_arr2 = split(db_msg2,"|")
	msg1 = msg_arr2(0)
	msg2 = msg_arr2(1)
	msg3 = msg_arr2(2)
	msg4 = msg_arr2(3)
	Set dbRecordSet = Nothing
	If msg1 <> "" and msg2 <> "" and  msg3 <> "" and  msg4 <> "" Then
		Append_TestHTML StepCounter, "Get status from Job 257", "Status Message1: " & msg1 & VBCRLF & "Status Message2: " & msg2 & VBCRLF & "Status Message3: " & msg3 & VBCRLF & "Status Message4: " & msg4, "PASSED"
	else
		Append_TestHTML StepCounter, "Get status from Job 257", "DB values not found. Status Message1: " & msg1 & VBCRLF & "Status Message2: " & msg2 & VBCRLF & "Status Message3: " & msg3 & VBCRLF & "Status Message4: " & msg4, "FAILED"
		bRunFlag = False
	End If
	wait 10
	
	Set dbRecordSet = execute_db_query("Select * from TranBatchFile Order by 1 desc;", 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 15
	Append_TestHTML StepCounter, "Run TranBatchFile Query", "Select * from TranBatchFile Order by 1 desc;", "PASSED"
	wait 2
	db_batchID =  dbRecordSet("BatchID")
	db_FileName = dbRecordSet("Filename")
	Set dbRecordSet = Nothing
	If db_batchID <> "" and db_FileName <>"" Then
		Append_TestHTML StepCounter, "Get values from TranBatchFile", "Batch ID: " & db_batchID & VBCRLF & "File Name: " & db_FileName, "PASSED"
	else
		Append_TestHTML StepCounter, "Get values from TranBatchFile", "DB values not found. Batch ID: " & db_batchID & VBCRLF & "File Name: " & db_FileName, "FAILED"
	End If	
	If instr(file_name, db_FileName) <> 0 Then
	Trans_BatchID = db_batchID
		Append_TestHTML StepCounter, "Validate DX026 file name", "File name matches. Expected Value: " & file_name & VBCRLF & "Actual Value: " & db_FileName, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate DX026 file name", "File name validation failed. Expected Value: " & file_name & VBCRLF & "Actual Value: " & db_FileName, "FAILED"
		bRunFlag = False
	End If
	
'	wait 20
	
'	Set dbRecordSet = execute_db_query("Select * from SalesItemNational where I_BatchID =  '" & db_batchID & "';", 2, "SFN_SHELL_SPRINTQA_ID_BATCH")
'	Append_TestHTML StepCounter, "Run SalesItemNational Query", "Select * from SalesItemNational where I_BatchID =  '" & db_batchID & "'", "PASSED"
'	wait 2
'	i = 1
'	Do While i < 10
'		i = i + 1
'		If isNull(dbRecordSet("SalesItemID")) Then
'			Set dbRecordSet = execute_db_query("Select * from SalesItemNational where I_BatchID =  '" & db_batchID & "';", 2, "SFN_SHELL_SPRINTQA_ID_BATCH")
'			wait 10
'		Else
'			Exit Do	
'		End If
'	Loop
'	db_salesItemID =  dbRecordSet("SalesItemID")
'	salesIDArr = split(db_salesItemID,"|")
'	db_sales_ID1 = salesIDArr(0)
'	db_sales_ID2 = salesIDArr(1)
'	Set dbRecordSet = Nothing
'	If db_sales_ID1 <> "" and db_sales_ID2 <> "" Then
'		Append_TestHTML StepCounter, "Get sales ID values", "Sales ID1: " & db_sales_ID1 & VBCRLF & "Sales ID2: " & db_sales_ID2, "PASSED"
'	else
'		Append_TestHTML StepCounter, "Get sales ID values", "DB values not found. Sales ID1: " & db_sales_ID1 & VBCRLF & "Sales ID2: " & db_sales_ID2, "FAILED"
'		bRunFlag = False
'	End If

Else

		Append_TestHTML StepCounter, "Validate Sales details in SalesItemNational table","Unable to verify becaus no card details or customer erp numbers", "FAILED"
	bRunFlag = False
	bFlag=False
End If	
End  Function




'******************************* HEADER ******************************************
' Description : The function to verify transaction billing in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function VerifyTxnBilling(bill_report_date, PAN_num)

	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
		
		
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Start", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Transactions", "Click", ""
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_BillAcceptance", "Click", ""
	
	wait 10
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Preview") Then
		Append_TestHTML StepCounter, "Billing Acceptance", "Navigated to billing acceptance page", "PASSED"
		wait 5
	else
		Append_TestHTML StepCounter, "Billing Acceptance", "Navigation to billing acceptance page failed", "FAILED"
		bFlag = False
		
	End If
	''''msgbox "Next will Billing Preview"
	query = "Select * from BillingAcceptance order by 1 desc;"
	Append_TestHTML StepCounter, "Get last billing report date from BillingAcceptance Table",query, "PASSED"
	set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	currentBillinRepDate= dictDbResultSet("BillingReportDate")
	''''msgbox currentBillinRepDate
	set dictDbResultSet = Nothing
	Append_TestHTML StepCounter, "Current Billing Report Date","Billing Report Date is:"& currentBillinRepDate  , "PASSED"
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Preview", "Click", ""
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_BillPreview", "Click", ""
	wait 10
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
		Append_TestHTML StepCounter, "Bill Preview", "Clicked on bill preview", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Bill Preview", "Clicking on bill preview button failed", "FAILED"
	End If
	
	Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_Refresh") = False
		wait 25
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Refresh", "Click", ""
		wait 15
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
			Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
		else
'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
				preview_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Preview", "GetROProperty", "innertext")
				breqreportdate = reqdateFormt()
				If instr(preview_msg, "Billing previewed") <> 0 or instr(preview_msg, breqreportdate) <> 0 Then
					Append_TestHTML StepCounter, "Validate bill preview", "Billing previewed successfully. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "PASSED"
'				else
'					Append_TestHTML StepCounter, "Validate bill preview", "Billing preview failed. Expected Value: " & bill_report_date_req  & VBCRLF & breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "FAILED"
'					bRunFlag = False
				End If
		End If	
	Loop
	
	query = "Select * from BillingAcceptance order by 1 desc;"
	Append_TestHTML StepCounter, "Get last billing report date from BillingAcceptance Table",query, "PASSED"
	set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	currentBillinRepDate= dictDbResultSet("BillingReportDate")
	''''msgbox currentBillinRepDate
	set dictDbResultSet = Nothing
	Append_TestHTML StepCounter, "Current Billing Report Date","Billing Report Date is:"& currentBillinRepDate  , "PASSED"
	
	
	bill_report_date_req = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_BillReportDate", "GetROProperty", "default value")
	bill_report_date_arr = split(bill_report_date_req,"/")
	bill_report_date = bill_report_date_arr(2) & "-" & bill_report_date_arr(1) & "-" & bill_report_date_arr(0)
	billReportDate = bill_report_date
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_BillReportDate") Then
		Append_TestHTML StepCounter, "Bill report date", "Captured billing report date displayed: "& billReportDate, "PASSED"
	else
		Append_TestHTML StepCounter, "Bill report date", "Billing report date not found", "FAILED"
		bRunFlag = False
		bFlag = False
	End If
	''''msgbox bill_report_date
	'''''msgbox Trans_BatchID
	wait 25
	'Perform DB operations
	If bill_report_date <> "" and Trans_BatchID <> "" Then
		
		
		query = "Select * from SalesItemUnbilled where batchid = '" & Trans_BatchID & "'"
		Append_TestHTML StepCounter,"Check billing Report date",query, "PASSED"
		Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		db_actual_report_date = dbRecordSet("BillingReportDate")
		date_arr = split(db_actual_report_date,"|")
		db_actual_report_date1 = date_arr(0)
		db_actual_report_date2 = date_arr(1)
		Set dbRecordSet = Nothing
		If db_actual_report_date1 <> "" and db_actual_report_date2 <> "" Then
			Append_TestHTML StepCounter, "Actual billing Report date", "Actual billing report date1: " & db_actual_report_date1 & VBCRLF & "Actual billing report date2: " & db_actual_report_date2, "PASSED"
		else
			Append_TestHTML StepCounter, "Actual billing Report date", "DB values not found. Actual billing report date1: " & db_actual_report_date1 & VBCRLF & "Actual billing report date2: " & db_actual_report_date2, "FAILED"
			bFlag = False
			
		End If
		
		query1 = "Select * from Job where JobTypeId = 19 order by 1 desc;"
		Append_TestHTML StepCounter, "Run Job 19", "Select * from Job where JobTypeId = 19 order by 1 desc;", "PASSED"
		Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		db_id_19 = dbRecordSet("ID")
		db_status_id_19 = dbRecordSet("StatusID")
		Set dbRecordSet = Nothing
		If db_id_19 <> "" and db_status_id_19 <> "" Then
			Append_TestHTML StepCounter, "Check Job 19 ID", "ID: " & db_id_19 & VBCRLF & "Status: " & db_status_id_19, "PASSED"
		else
			Append_TestHTML StepCounter, "Check Job 19 ID", "DB values not found. ID: " & db_id_19 & VBCRLF & "Status: " & db_status_id_19, "FAILED"
		bFlag = False
			
		End If
		
		query2 = "Select * from InfoProviderNextRunDate where InfoProviderID = 4"
		Append_TestHTML StepCounter, "Check InfoProviderNextRunDate", "Select * from InfoProviderNextRunDate where InfoProviderID = 4", "PASSED"
		Set dbRecordSet = execute_db_query(query2, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		db_next_report_date = dbRecordSet("NextReportDate")
		report_date_arr = split(db_next_report_date,"|")
		db_next_report_date1 = report_date_arr(0)
		db_next_report_date2 = report_date_arr(1)
		db_next_run_date = dbRecordSet("NextRunDate")
		run_date_arr = split(db_next_run_date,"|")
		db_next_run_date1 = run_date_arr(0)
		db_next_run_date2 = run_date_arr(1)
		Set dbRecordSet = Nothing
		If db_next_report_date1 <> "" and db_next_report_date2 <> ""  and db_next_run_date1 <> "" and db_next_run_date2 <> "" Then
			Append_TestHTML StepCounter, "Check dates in InfoProviderNextRunDate", "Next Report Date1: " & db_next_report_date1 & VBCRLF & "Next Report Date2: " & db_next_report_date2 & VBCRLF & "Next Run Date1: " & db_next_run_date1 & VBCRLF & "Next Run Date2: " & db_next_run_date2, "PASSED"
		else
			Append_TestHTML StepCounter, "Check dates in InfoProviderNextRunDate", "DB values not found. Next Report Date1: " & db_next_report_date1 & VBCRLF & "Next Report Date2: " & db_next_report_date2 & VBCRLF & "Next Run Date1: " & db_next_run_date1 & VBCRLF & "Next Run Date2: " & db_next_run_date2, "FAILED"
		bFlag = False
		
		End If
		
		query3 = "Select * from BillingAcceptance order by 1 desc;"
		Append_TestHTML StepCounter, "Check BillingAcceptance", "Select * from BillingAcceptance order by 1 desc;", "PASSED"
		Set dbRecordSet = execute_db_query(query3, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		db_billing_id = dbRecordSet("BillingAcceptanceID")
		db_billing_report_date = dbRecordSet("BillingReportDate")
		Set dbRecordSet = Nothing
		If db_billing_id <> "" and db_billing_report_date <> "" Then
			Append_TestHTML StepCounter, "Check BillingAcceptance", "Billing ID: " & db_billing_id & VBCRLF & "Billing Report Date: " & db_billing_report_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Check BillingAcceptance", "DB values not found. Billing ID: " & db_billing_id & VBCRLF & "Billing Report Date: " & db_billing_report_date, "FAILED"
		bFlag = False
		
		End If
		If bill_report_date = db_billing_report_date  Then
			Append_TestHTML StepCounter, "Validate UI and Billing Acceptance table report data", "Billing dates are same- Report Date: " & db_billing_report_date, "PASSED"
			Else
			Append_TestHTML StepCounter, "Validate UI and Billing Acceptance table report data", "Billing dates are different- Acceptance table Report Date: " & db_billing_report_date & "UI-"&bill_report_date, "PASSED"
		bFlag = False
			
		End If
'		query4 = "Update SalesItemUnbilled Set BillingReportDate = '" & bill_report_date & "' where batchid = '" & Trans_BatchID & "'"
'		Append_TestHTML StepCounter, "Update BillingReportDate in SalesItemUnbilled", query4, "PASSED"
'		Set dbRecordSet = execute_db_query(query4 ,1,"SFN_SHELL_SPRINTQA_ID_BATCH")
'		wait 2
'		Set dbRecordSet = Nothing
		wait 15
		Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
		Append_TestHTML StepCounter, "Check updated billing date", query, "PASSED"
		wait 2
		db_updated_report_date = dbRecordSet("BillingReportDate")
		date_updated_arr = split(db_updated_report_date,"|")
		db_updated_report_date1 = date_updated_arr(0)
		db_updated_report_date2 = date_updated_arr(1)
		Set dbRecordSet = Nothing
		If db_updated_report_date1 <> "" and db_updated_report_date2 <> "" Then
			Append_TestHTML StepCounter, "Updated billing Report date", "Updated billing report date1: " & db_updated_report_date1 & VBCRLF & "Updated billing report date2: " & db_updated_report_date2, "PASSED"
		else
			Append_TestHTML StepCounter, "Updated billing Report date", "DB values not found. Updated billing report date1: " & db_updated_report_date1 & VBCRLF & "Updated billing report date2: " & db_updated_report_date2, "FAILED"
		bFlag = False
		
		End If
		
		If instr(db_updated_report_date1, bill_report_date) <> 0 Then
			Append_TestHTML StepCounter, "Validate billing Report date", "Billing report date1 is updated successfully. Expected Value: " & db_updated_report_date1 & VBCRLF & "Actual Value: " & bill_report_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate billing Report date", "Billing report date1 update failed. Expected Value: " & db_updated_report_date1 & VBCRLF & "Actual Value: " & bill_report_date, "FAILED"
		bFlag = False
		
		End If
		
		If instr(db_updated_report_date2, bill_report_date) <> 0 Then
			Append_TestHTML StepCounter, "Validate billing Report date", "Billing report date2 is updated successfully. Expected Value: " & db_updated_report_date2 & VBCRLF & "Actual Value: " & bill_report_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate billing Report date", "Billing report date2 update failed. Expected Value: " & db_updated_report_date2 & VBCRLF & "Actual Value: " & bill_report_date, "FAILED"
			bRunFlag = False
		bFlag = False
		
		End If
	'----------------Updating InfoProviderNextrundate table data	
		
		bill_next_report_date = nextdateDBFormt(bill_report_date)
		
'		query_Info_next_run = "update InfoProviderNextRunDate set NextReportDate='"& bill_report_date &"',NextRunDate='"& bill_next_report_date &"' where infoproviderid = 4;"
'		Append_TestHTML StepCounter, "Update NextReportdate in InfoProviderNextRunDate", query_Info_next_run, "PASSED"
'		Set dbRecordSet = execute_db_query(query_Info_next_run ,1,"SFN_SHELL_SPRINTQA_ID_BATCH")
'		wait 2
'		Set dbRecordSet = Nothing
		query_Info_run = "Select * from InfoProviderNextRunDate where infoproviderid = 4;"
		Set dbRecordSet = execute_db_query(query_Info_run, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
		Append_TestHTML StepCounter, "Check updated reported dates", query_Info_run , "PASSED"
		wait 2
		db_Info_report_date = dbRecordSet("NextReportDate")
		Info_report_date = split(db_Info_report_date,"|")
		db_info_report_date1 = Info_report_date(0)
		db_info_report_date2 = Info_report_date(1)
		
			db_Info_nrun_date = dbRecordSet("NextRunDate")
		Info_report_nrdate = split(db_Info_nrun_date,"|")
		db_info_report_ndate1 = Info_report_nrdate(0)
		db_info_report_ndate2 = Info_report_nrdate(1)
		
		Set dbRecordSet = Nothing
		If db_info_report_date1 <> "" and db_info_report_date2 <> "" Then
			Append_TestHTML StepCounter, "Updated next Report date", "Updated next Info report date1: " & db_info_report_date1 & VBCRLF & "Updated next info report date2: " & db_info_report_date2, "PASSED"
		else
			Append_TestHTML StepCounter, "Updated next Report date", "DB values are not Updated. next Info report date1: " & db_info_report_date1 & VBCRLF & "Updated next info report date2: " & db_info_report_date2, "FAILED"
		bFlag = False
		
		End If
		
		If instr(db_info_report_date1, bill_report_date) <> 0 Then
			Append_TestHTML StepCounter, "Validate next Report date", "Next report date1 is updated successfully. Expected Value: " & db_info_report_date1 & VBCRLF & "Actual Value: " & bill_report_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate next Report date", "Next report date1 update failed. Expected Value: " & db_info_report_date1 & VBCRLF & "Actual Value: " & bill_report_date, "FAILED"
		bFlag = False
		
		End If
'		If instr(db_info_report_ndate2,bill_next_report_date) <> 0 Then
'			Append_TestHTML StepCounter, "Validate next Report run ate", "Next run report date is updated successfully. Expected Value: " & db_info_report_ndate2 & VBCRLF & "Actual Value: " & bill_next_report_date, "PASSED"
'		else
'			Append_TestHTML StepCounter, "Validate next Report run date", "Next run report date update failed. Expected Value: " & db_info_report_ndate2 & VBCRLF & "Actual Value: " & bill_next_report_date, "FAILED"
'		End If
	'-------------------------------------------	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''''''		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_BillPreview", "Click", ""
''''''		
''''''		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
''''''			Append_TestHTML StepCounter, "Bill Preview", "Clicked on bill preview", "PASSED"
''''''		'else
''''''		'	Append_TestHTML StepCounter, "Bill Preview", "Clicking on bill preview button failed", "FAILED"
''''''		End If
''''''		
''''''		Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_Refresh") = False
''''''			wait 15
''''''			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Refresh", "Click", ""
''			wait 15
''''''			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
''''''				Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
''''''			else
''''''	'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
''''''				preview_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Preview", "GetROProperty", "innertext")
''''''				breqreportdate = reqdateFormt()
''''''				If instr(preview_msg, "Billing previewed") <> 0 or instr(preview_msg, breqreportdate) <> 0 Then
''''''					Append_TestHTML StepCounter, "Validate bill preview", "Billing previewed successfully. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "PASSED"
'''''''				else
'''''''					Append_TestHTML StepCounter, "Validate bill preview", "Billing preview failed. Expected Value: " & bill_report_date_req  & VBCRLF & breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "FAILED"
'''''''					bRunFlag = False
''''''				End If
''''''			End If
''''''		Loop
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Company") Then
			Append_TestHTML StepCounter, "Validate Txns", "Transactions updated with bill report date", "PASSED"
		else
			Append_TestHTML StepCounter, "Validate Txns", "Transaction details not displayed", "FAILED"
		bFlag = False
			
		End If
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_BilledFees", "Click", ""
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Currency") Then
			Append_TestHTML StepCounter, "Validate Fee details", "Fee details updated with bill report date", "PASSED"
		else
			Append_TestHTML StepCounter, "Validate Fee details", "Fee details not displayed", "FAILED"
			bFlag = False
			
		End If
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_BilledTxns", "Click", ""
		wait 2
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Cutoff", "Click", ""
		wait 4
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
			Append_TestHTML StepCounter, "Cutoff", "Clicked on cutoff", "PASSED"
		'else
			'Append_TestHTML StepCounter, "Cutoff", "Clicking on cutoff failed", "FAILED"
		End If
		
		'add refresh in between and validate side text
		Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_Refresh") = False
			wait 25
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Refresh", "Click", ""
			wait 15
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
				Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
			else
	'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
				cutoff_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Cutoff", "GetROProperty", "innertext")
				breqreportdate = reqdateFormt()
				If instr(cutoff_msg, "Billing cut off") <> 0 or instr(cutoff_msg, breqreportdate) <> 0 Then
					Append_TestHTML StepCounter, "Validate bill cutoff", "Billing cutoff successful. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate  & VBCRLF & "Actual Value: " & cutoff_msg, "PASSED"
				else
					Append_TestHTML StepCounter, "Validate bill cutoff", "Billing cutoff failed. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate  & VBCRLF & "Actual Value: " & cutoff_msg, "FAILED"
					bRunFlag = False
					bFlag = False
					
				End If
			End If
		Loop
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Signoff", "Click", ""
		wait 20
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_OK") Then
			Append_TestHTML StepCounter, "Signoff", "Clicked on signoff", "PASSED"
		else
			Append_TestHTML StepCounter, "Signoff", "Clicking on signoff failed", "FAILED"
			bFlag = False
			
		End If
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebElement_OK", "Click", ""
		
		wait 10
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
			Append_TestHTML StepCounter, "Signoff", "Signoff started", "PASSED"
'		else
'			Append_TestHTML StepCounter, "Signoff", "Signoff not started", "FAILED"
		End If
		
		'again refresh and validate
		Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_Refresh") = False
			wait 25
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebButton_Refresh", "Click", ""
			wait 15
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_Refresh") Then
				Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
			else
				wait 2
	'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
			End If
		Loop
		
		preview_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Preview", "GetROProperty", "innertext")
		breqreportdate = reqdateFormt()
		If instr(preview_msg, "Billing previewed") <> 0 or instr(preview_msg, breqreportdate) <> 0 Then
			Append_TestHTML StepCounter, "Validate bill preview", "Billing previewed successfully. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate bill preview", "Billing preview failed. Expected Value: " & bill_report_date_req  & VBCRLF & breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "FAILED"
			bRunFlag = False
			bFlag = False
			
		End If
		
		cutoff_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Cutoff", "GetROProperty", "innertext")
		
		If instr(cutoff_msg, "Billing cut off") <> 0 or instr(cutoff_msg, breqreportdate) <> 0 Then
			Append_TestHTML StepCounter, "Validate bill cutoff", "Billing cutoff successful. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate  & VBCRLF & "Actual Value: " & cutoff_msg, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate bill cutoff", "Billing cutoff failed. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate  & VBCRLF & "Actual Value: " & cutoff_msg, "FAILED"
			bRunFlag = False
			bFlag = False
			
		End If
		
		signoff_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Signoff", "GetROProperty", "innertext")
	
		If instr(signoff_msg, "Billing signed off") <> 0 or instr(signoff_msg, breqreportdate) <> 0 Then
			Append_TestHTML StepCounter, "Validate bill signoff", "Billing signoff successful. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate  & VBCRLF & "Actual Value: " & signoff_msg, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate bill signoff", "Billing signoff failed. Expected Value: " & bill_report_date_req & VBCRLF & breqreportdate  & VBCRLF & "Actual Value: " & signoff_msg, "FAILED"
			bRunFlag = False
			bFlag = False
			
		End If
		
		next_preview_date = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_NextPreview", "GetROProperty", "value")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_NextPreview") Then
			Append_TestHTML StepCounter, "Validate Next preview date", "Next preview date displayed: " & next_preview_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Validate Next preview date", "Next preview date not displayed: " & next_preview_date, "FAILED"
			bRunFlag = False
			bFlag = False
			
		End If
		
		'DB operations	
		Set dbRecordSet = execute_db_query(query2, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
		Append_TestHTML StepCounter, "Check updated InfoProviderNextRunDate", "Select * from InfoProviderNextRunDate where InfoProviderID = 4", "PASSED"
		wait 2
		db_next_report_date1 = dbRecordSet("NextReportDate")
		report_date_arr1 = split(db_next_report_date1,"|")
		db_next_report_date3 = report_date_arr1(0)
		db_next_report_date4 = report_date_arr1(1)
		db_next_run_date = dbRecordSet("NextRunDate")
		run_date_arr1 = split(db_next_run_date,"|")
		db_next_run_date3 = run_date_arr1(0)
		db_next_run_date4 = run_date_arr1(1)
		Set dbRecordSet = Nothing
		If db_next_report_date3 <> "" and db_next_report_date4 <> ""  and db_next_run_date3 <> "" and db_next_run_date4 <> "" Then
			Append_TestHTML StepCounter, "Check updated dates in InfoProviderNextRunDate", "Updated Next Report Date1: " & db_next_report_date3 & VBCRLF & "Updated Next Report Date2: " & db_next_report_date4 & VBCRLF & "Updated Next Run Date1: " & db_next_run_date3 & VBCRLF & "Updated Next Run Date2: " & db_next_run_date4, "PASSED"
		else
			Append_TestHTML StepCounter, "Check updated dates in InfoProviderNextRunDate", "DB values not found. Updated Next Report Date1: " & db_next_report_date3 & VBCRLF & "Updated Next Report Date2: " & db_next_report_date4 & VBCRLF & "Updated Next Run Date1: " & db_next_run_date3 & VBCRLF & "Updated Next Run Date2: " & db_next_run_date4, "FAILED"
		bRunFlag = False
			bFlag = False
		
		End If
		
		Set dbRecordSet = execute_db_query(query3, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
		Append_TestHTML StepCounter, "Check updated BillingAcceptance", "Select * from BillingAcceptance order by 1 desc;", "PASSED"
		wait 2
		db_billing_id1 = dbRecordSet("BillingAcceptanceID")
		db_billing_report_date1 = dbRecordSet("BillingReportDate")
		Set dbRecordSet = Nothing
		If db_billing_id1 <> "" and db_billing_report_date1 <> "" Then
			Append_TestHTML StepCounter, "Check updated dates in BillingAcceptance", "Updated Billing ID: " & db_billing_id & VBCRLF & "Updated Billing Report Date: " & db_billing_report_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Check updated dates in BillingAcceptance", "DB values not found. Updated Billing ID: " & db_billing_id & VBCRLF & "Updated Billing Report Date: " & db_billing_report_date, "FAILED"
			bRunFlag = False
			bFlag = False
			
		End If	
	Else
			Append_TestHTML StepCounter, "Update backdates for Billing Acceptance and Billing Preview process", "Unable to process because no BatchID of transaction or no valid billing report date", "FAILED"
			bFlag = False
	
End If	
End  Function



'******************************* HEADER ******************************************
' Description : The function to verify billing job 19 in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function VerifyBillingJob19(bill_report_date, PAN_num)

	On error resume next
	
	bFlag = True
If PAN_num <> "" Then
		
	
	query = "Select * from Job where JobTypeId = 19 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 19", "Select * from Job where JobTypeId = 19 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_id <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 19", "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 19", "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "FAILED"
			bFlag = False
		
	End If
	If db_status_id = 0 Then
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is not successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
			bFlag = False
		
	End If
	
	query1 = "Update Job Set NextRunDate = null where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Update Job 19", "Update Job Set NextRunDate = null where ID = '" & db_id & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	Set dbRecordSet = Nothing
	JobRundate = Date()
	wait 25
	query2 = "Select * from Job where JobTypeId = 19 and ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 19", "Select * from Job where JobTypeId = 19 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_status_id <> "" Then
		Append_TestHTML StepCounter, "Verify Job 19", "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Verify Job 19", "DB values not found. Status ID: " & db_status_id, "FAILED"
			bFlag = False
		
	End If
	If db_status_id = 4 Then
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID after update is successful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID after update is not successful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
		bRunFlag = False
			bFlag = False
		
	End If
Else

		Append_TestHTML StepCounter, "Trigger Job 19", "Unable to trigger job because no transaction details", "FAILED"
			bFlag = False

End If

End  Function


Public Function VerifyH3JobStatus(jobTID,bill_report_date, PAN_num)
''''msgbox "job64"
	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
	End If
	
	If bill_report_date = ""  or bill_report_date = empty Then
		bill_report_date = billReportDate
	End If
If PAN_num <> "" and bill_report_date<> "" Then
	
	query = "Select * from Job where JobTypeId = "& jobTID  &" order by 1 desc;"
	Append_TestHTML StepCounter, "Check Job Status "& jobTID, query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	db_date_created = dbRecordSet("DateCreated")
	Set dbRecordSet = Nothing
	yes_date = Date()-1
	todays_date = nextdateDBFormt(yes_date)
	
	If db_id <> "" and db_status_id <> "" and (instr(db_date_created,day(Date())) >0 and instr(db_date_created,month(Date()) )>0 ) Then
		Append_TestHTML StepCounter, "Get ID from Job " & jobTID, "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
		If db_status_id = 4 Then
			VerifyH3JobStatus = True
			Append_TestHTML StepCounter, "Validated status ID", "job Triggered successfully. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
		else
			VerifyH3JobStatus = True
			Append_TestHTML StepCounter, "Validated status ID", "job unsuccessful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
		End If
	else
		VerifyH3JobStatus = False
		Append_TestHTML StepCounter, "Get ID from Job "& jobTID , "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id & "Date Created:" & db_date_created  & "-" & todays_date , "FAILED"
			bFlag = False
		
	End If
	
	
Else
VerifyH3JobStatus = False
	Append_TestHTML StepCounter, "Validate Job" & jobTID & " status", "Unable to check the details because no billing transaction and job 19 not initiated", "FAILED"
		bRunFlag = False	
	
			bFlag = False
	
End If	
End  Function



'******************************* HEADER ******************************************
' Description : The function to verify billing job 64 in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function VerifyBillingJob64(bill_report_date, PAN_num)
''''msgbox "job64"
	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
	End If
	
	If bill_report_date = ""  or bill_report_date = empty Then
		bill_report_date = billReportDate
	End If
If PAN_num <> "" and bill_report_date<> "" Then
	
	query = "Select * from Job where JobTypeId = 64 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 64", "Select * from Job where JobTypeId = 64 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_id <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 64", "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 64", "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "FAILED"
			bFlag = False
		
	End If
'	If db_status_id = 0 Then
'		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
'	else
'		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is not successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
'	End If
	
	query1 = "Select * from SalesItem where PAN = '" & PAN_num & "'"
	Append_TestHTML StepCounter, "Verify Job 64", "Select * from SalesItem where PAN = '" & PAN_num & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_report_date = dbRecordSet("BillingReportDate")
	db_sales_item_id = dbRecordSet("SalesItemID")
	date_arr = split(db_bill_report_date,"|")
	sales_arr = split(db_sales_item_id,"|")
	db_bill_report_date1 = date_arr(0)
	db_bill_report_date2 = date_arr(1)
	db_sales_item_id1 = sales_arr(0)
	db_sales_item_id2 = sales_arr(1)
	Set dbRecordSet = Nothing
	If db_bill_report_date1 <> "" and db_bill_report_date2 <> ""  and db_sales_item_id1 <> "" and db_sales_item_id2 <> "" Then
		Append_TestHTML StepCounter, "Verify Sales Item table", "Bill Report Date1: " & db_bill_report_date1 & VBCRLF & "Bill Report Date2: " & db_bill_report_date2 & VBCRLF & "Sales Item ID1: " & db_sales_item_id1 & VBCRLF & "Sales Item ID2: " & db_sales_item_id2, "PASSED"
	else
		Append_TestHTML StepCounter, "Verify Sales Item table", "DB values not found. Bill Report Date1: " & db_bill_report_date1 & VBCRLF & "Bill Report Date2: " & db_bill_report_date2 & VBCRLF & "Sales Item ID1: " & db_sales_item_id1 & VBCRLF & "Sales Item ID2: " & db_sales_item_id2, "FAILED"
			bFlag = False
	
	End If
	If instr(db_bill_report_date1, bill_report_date) <> 0 Then
		Append_TestHTML StepCounter, "Validate Sales Item table", "Txns moved to Sales Item table with correct billing date. Expected Value: " & db_bill_report_date1 & VBCRLF & "Actual Value: " & bill_report_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Sales Item table", "Txns not found in Sales Item table. Expected Value: " & db_bill_report_date1 & VBCRLF & "Actual Value: " & bill_report_date, "FAILED"
		bRunFlag = False
			bFlag = False
	
	End If
	If instr(db_bill_report_date2, bill_report_date) <> 0 Then
		Append_TestHTML StepCounter, "Validate Sales Item table", "Txns moved to Sales Item table with correct billing date. Expected Value: " & db_bill_report_date2 & VBCRLF & "Actual Value: " & bill_report_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Sales Item table", "Txns not found in Sales Item table. Expected Value: " & db_bill_report_date2 & VBCRLF & "Actual Value: " & bill_report_date, "FAILED"
		bRunFlag = False	
			bFlag = False
	
	End If
Else
	Append_TestHTML StepCounter, "Validate Job64 status", "Unable to check the details because no billing transaction and job 19 not initiated", "FAILED"
		bRunFlag = False	
	
			bFlag = False
	
End If	
End  Function




Function setGlobaldataforFilesValidation(PAN_num)
	On error resume next
	bFlag = True
If PAN_num <> "" Then

	query1 = "Select PostingDateTime,BillingReportdate,BatchID,* from SalesItem where PAN = '" & PAN_num & "';"
	Append_TestHTML StepCounter, "Get BillingReportdate & BatchID from SalesItem", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_report_date = dbRecordSet("BillingReportDate")
	db_batch_id = dbRecordSet("BatchID")
	date_arr = split(db_bill_report_date,"|")
	batch_arr = split(db_batch_id,"|")
	db_bill_report_date1 = date_arr(0)
	db_bill_report_date2 = date_arr(1)
	db_batch_item_id1 = batch_arr(0)
	db_batch_item_id2 = batch_arr(1)
	Set dbRecordSet = Nothing
	Trans_bill_report_date = db_bill_report_date1
	Trans_sales_batchid = db_batch_item_id1
	If cdate(db_bill_report_date1) = cdate( db_bill_report_date2) Then
		
		Append_TestHTML StepCounter, "Validate Billing reportdate", "DB value found. Same Billing report date:"&Trans_bill_report_date & VBCRLF & "for transaction", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Billing reportdate", "DB value not found. different Billing report date:" & db_bill_report_date1 & VBCRLF & "Bill Report Date2: " & db_bill_report_date2 & VBCRLF , "FAILED"
	bFlag = False
	End If
	If db_batch_item_id1 =  db_batch_item_id2 Then
		
		Append_TestHTML StepCounter, "Validate Batch ID", "DB value found. Same Batch ID:"&Trans_sales_batchid & VBCRLF & "for transaction", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Batch ID", "DB value not found. different Batch ID1:" & db_batch_item_id1 & VBCRLF & "Batch ID2: " & db_batch_item_id2 & VBCRLF , "FAILED"
		bFlag = False
	End If
	
	'query2 = "Select BillingDocumentID,* from SalesItem where BillingReportDate = '" & Trans_bill_report_date & "';"
	query2 = "Select BillingDocumentID,* from SalesItem where BatchID = '" & Trans_sales_batchid & "';"
	
	Append_TestHTML StepCounter, "Get BillingDocumentID from SalesItem", query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_report_date = dbRecordSet("BillingDocumentID")
	
	docid_arr = split(db_bill_report_date,"|")
	
	db_bill_doc_id1 = docid_arr(0)
	db_bill_doc_id2= docid_arr(1)

	Set dbRecordSet = Nothing
	Trans_bill_doc_id = db_bill_doc_id1
	
	If db_bill_doc_id1 =  db_bill_doc_id2 Then
		
		Append_TestHTML StepCounter, "Validate Billing Document ID", "DB value found. Same Billing Document ID:"&Trans_bill_doc_id & VBCRLF & "for transaction", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Billing Document ID", "DB value not found. different Billing Document ID1:" & db_bill_doc_id1 & VBCRLF & "Bill Document ID2: " & db_bill_doc_id2 & VBCRLF , "FAILED"
	bFlag = False
	End If
	
	query3 = "Select * from BillingDocument where BillingDocumentID = '" & Trans_bill_doc_id & "';"
	Append_TestHTML StepCounter, "Get Sales Transaction details", query3, "PASSED"
	Set dbRecordSet = execute_db_query(query3, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_summary_docid = dbRecordSet("SummaryBillingDocumentID")
	
	db_bill_docid = dbRecordSet("BillingDocumentID")
	db_bill_sessionid = dbRecordSet("BillingSessionID")
	db_bill_invoice_cust_id = dbRecordSet("InvoiceCustomerID")
	db_bill_doc_no = dbRecordSet("BillingDocumentNumber")
	db_bill_doc_date= dbRecordSet("BillingDocumentDate")

	Set dbRecordSet = Nothing
	Trans_bill_summary_doc_id = db_bill_summary_docid
	
	If db_bill_summary_docid <> "" and db_bill_docid <> "" and db_bill_sessionid <> "" and db_bill_invoice_cust_id<> "" and  db_bill_doc_no<> ""  and db_bill_doc_date <> ""Then
		
		Append_TestHTML StepCounter, "Get Summary Document ID-", "DB value found. Summary Document ID:"&db_bill_summary_docid & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Billing Document ID-", "DB value found.  Billing Document ID:"&db_bill_docid & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Billing Session ID-", "DB value found.  Billing Session ID:"&db_bill_sessionid & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Invoice Customer ID-", "DB value found.  Invoice Customer ID:"&db_bill_invoice_cust_id & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Billing Document Number-", "DB value found. Billing Document Number:"&db_bill_doc_no & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Billing Document Date-", "DB value found. Billing Document date:"&db_bill_doc_date & "for transaction", "PASSED"
		
	else
		Append_TestHTML StepCounter, "Get Sales Transaction", "DB value not found. different Summary Document ID1:" & db_bill_summary_docid  , "FAILED"
	bFlag = False
	End If
	
	query4 = "Select * from SummaryBillingDocument where SummaryBillingDocumentID = '" & Trans_bill_summary_doc_id & "';"
	Append_TestHTML StepCounter, "Get Summary Billing Document details", query4, "PASSED"
	Set dbRecordSet = execute_db_query(query4, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_summary_docid1 = dbRecordSet("SummaryBillingDocumentID")
	
	db_bill_sum_doc_no = dbRecordSet("SummaryDocumentNumber")
	db_bill_sessionid1 = dbRecordSet("BillingSessionID")

	dx350_Summary_Docno = db_bill_sum_doc_no
	dx350_Balance =  dbRecordSet("Balance")
	dx350_Paymentduedate = dbRecordSet("PaymentDueDate") 
	Set dbRecordSet = Nothing
	Trans_bill_session_id = db_bill_sessionid1
	
	If db_bill_summary_docid1 <> "" and db_bill_sum_doc_no <> "" and db_bill_sessionid1 <> "" Then
		
		Append_TestHTML StepCounter, "Get Summary Document ID-", "DB value found. Summary Document ID:"&db_bill_summary_docid1 & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get SummaryDocumentNumber-", "DB value found.  SummaryDocumentNumber:"&db_bill_sum_doc_no & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Billing Session ID-", "DB value found.  Billing Session ID:"&db_bill_sessionid1 & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Billing Payment due date", "DB value found.  Billing PaymentDue date:"&dx350_Paymentduedate & "for transaction", "PASSED"
		
		
	else
		Append_TestHTML StepCounter, "Get Summary Billing Document deetails", "DB value not found. different Summary Billing Document ID1:" & db_bill_summary_docid1  , "FAILED"
	bFlag = False
	End If
	query5 = "Select * from BillingSession where BillingSessionID = '" & Trans_bill_session_id & "';"
	Append_TestHTML StepCounter, "Get Summary Documents", query5, "PASSED"
	Set dbRecordSet = execute_db_query(query5, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_no_summary_docs = dbRecordSet("NoOfSummaryDocuments")
	db_bill_no_docs = dbRecordSet("NoOfDocuments")
	Set dbRecordSet = Nothing
	
		If db_bill_no_summary_docs <> "" and db_bill_no_docs <> ""  Then
		
		Append_TestHTML StepCounter, "Get Number of Summary Documents -", "DB value found. No of Summary Documents are :"&db_bill_no_summary_docs & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get number of documents-", "DB value found.  No Of Documents:"&db_bill_no_docs & "for transaction", "PASSED"
				
	else
		Append_TestHTML StepCounter, "Get Number of Document details", "DB value not found."  , "FAILED"
	bFlag = False
	End If
	
'	query6_1="Select count(jobtypeid) as No_of_rows from BillingSessionoutput where BillingSessionID = '"& Trans_bill_session_id &"';"
'	Set dbRecordSet = execute_db_query(query6_1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
'	no_of_jobs = dbRecordSet("No_of_rows")
'	Set dbRecordSet = Nothing
	
	query6 = "Select * from BillingSessionoutput where BillingSessionID = '" & Trans_bill_session_id & "';"
	Append_TestHTML StepCounter, "Get Job Types for billing session", query6, "PASSED"
	Set dbRecordSet = execute_db_query(query6, 3,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_bill_job_type_id = dbRecordSet("JobTypeID")
	
	
	jobtypeid_arr = split(db_bill_job_type_id,"|")
	db_bill_jobtype_id1 = jobtypeid_arr(0)
	db_bill_jobtype_id2 = jobtypeid_arr(1)
	db_bill_jobtype_id3 = jobtypeid_arr(2)
	
	db_bill_job_id = dbRecordSet("JobID")
	jobid_arr = split(db_bill_job_id,"|")
	
	db_bill_job_id1 = jobid_arr(0)
	db_bill_job_id2 = jobid_arr(1)
	db_bill_job_id3 = jobid_arr(2)
	
	Set dbRecordSet = Nothing
	If db_bill_jobtype_id1 = "211"  Then
		db_211_job_id = db_bill_job_id1
	ElseIf  db_bill_jobtype_id2 = "211" Then
		db_211_job_id = db_bill_job_id2
	ElseIf db_bill_jobtype_id3 = "211" Then
		db_211_job_id = db_bill_job_id3
	Else
		db_211_job_id = ""
	End If
	If db_bill_jobtype_id1 = "264"  Then
		db_264_job_id = db_bill_job_id1
	ElseIf  db_bill_jobtype_id2 = "264" Then
		db_264_job_id = db_bill_job_id2
	ElseIf db_bill_jobtype_id3 = "264" Then
		db_264_job_id = db_bill_job_id3
	Else
		db_264_job_id = ""
	End If
	If db_bill_jobtype_id1 = "305"  Then
		db_305_job_id = db_bill_job_id1
	ElseIf  db_bill_jobtype_id2 = "305" Then
		db_305_job_id = db_bill_job_id2
	ElseIf db_bill_jobtype_id3 = "305" Then
		db_305_job_id = db_bill_job_id3
	Else
		db_305_job_id = ""
	End If
	
	
	If db_bill_jobtype_id1 <> "" or db_bill_jobtype_id2 <> "" or db_bill_jobtype_id3<> ""  Then
		
		Append_TestHTML StepCounter, "Job Triggered Type details-", "DB value found. Job TypeID1 :"&db_bill_jobtype_id1& VBCRLF&"job TypeID2 :"&db_bill_jobtype_id2& VBCRLF &"job TypeID3 :"&db_bill_jobtype_id3& VBCRLF   &  "for this transaction", "PASSED"
			
	else
		Append_TestHTML StepCounter, "Job Triggered Type details", "DB value not found."  , "FAILED"
	bFlag = False
	End If
	If db_bill_job_id1 <> "" or db_bill_job_id2 <> "" or db_bill_job_id3<> ""  Then
		
		Append_TestHTML StepCounter, "Job IDs -", "DB value found. Job ID1 :"&db_bill_job_id1& VBCRLF&"job ID2 :"&db_bill_job_id2& VBCRLF &"job ID3 :"&db_bill_job_id3& VBCRLF   &  "for this transaction", "PASSED"
		Append_TestHTML StepCounter, "Job IDs -", "DB value found. Job ID1 :"&db_211_job_id& VBCRLF&"job ID2 :"&db_264_job_id& VBCRLF &"job ID3 :"&db_305_job_id& VBCRLF   &  "for this transaction", "PASSED"
			
	else
		Append_TestHTML StepCounter, "Job Ds", "DB value not found."  , "FAILED"
	bFlag = False
	End If
Else	
		Append_TestHTML StepCounter, "Get Billing document details for file validation", "Unable to get the details because transaction and job not performed"  , "FAILED"
	bFlag = False

End If		
End Function





Function setGlobaldataforFilesValidationforSettlement(PAN_num)
On error resume next
	bFlag = True
If PAN_num <> "" Then

	query1 = "Select SettlementReportDate,BatchID,* from SalesItemDelco where PAN = '" & PAN_num & "' order by 1 desc;"
	Append_TestHTML StepCounter, "Get SettlementReportDate & BatchID from SalesItemDelco", "Select * from SalesItemDelco where PAN = '" & PAN_num & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_settlement_report_date = dbRecordSet("SettlementReportDate")
	db_batch_id = dbRecordSet("BatchID")
	date_arr = split(db_settlement_report_date,"|")
	batch_arr = split(db_batch_id,"|")
	db_settlement_report_date1 = date_arr(0)
	db_settlement_report_date2 = date_arr(1)
	db_batch_item_id1 = batch_arr(0)
	db_batch_item_id2 = batch_arr(1)
	Set dbRecordSet = Nothing
	Trans_settlement_report_date = db_settlement_report_date1
	Trans_sales_batchid = db_batch_item_id1
	If cdate(db_settlement_report_date1) = cdate( db_settlement_report_date2) Then
		
		Append_TestHTML StepCounter, "Validate Settlement reportdate", "DB value found. Same Settlement report date:"&Trans_settlement_report_date & VBCRLF & "for transaction", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Settlement reportdate", "DB value not found. different Settlement report date:" & db_settlement_report_date1 & VBCRLF & "settlement Report Date2: " & db_settlement_report_date2 & VBCRLF , "FAILED"
	bFlag = False
	
	End If
	If db_batch_item_id1 =  db_batch_item_id2 Then
		
		Append_TestHTML StepCounter, "Validate Batch ID", "DB value found. Same Batch ID:"&Trans_sales_batchid & VBCRLF & "for transaction", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Batch ID", "DB value not found. different Batch ID1:" & db_batch_item_id1 & VBCRLF & "Batch ID2: " & db_batch_item_id2 & VBCRLF , "FAILED"
	bFlag = False
	
	End If
	
	'query2 = "Select SettlementDocumentID,* from SalesItemDelco where SettlementReportDate = '" & Trans_settlement_report_date & "'"
	query2 = "Select SettlementDocumentID,* from SalesItemDelco where BatchID = '" & Trans_sales_batchid & "'"
	
	Append_TestHTML StepCounter, "Get SettlementDocumentID from SalesItemDelco", query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_settlement_report_date = dbRecordSet("SettlementDocumentID")
	
	docid_arr = split(db_settlement_report_date,"|")
	
	db_settlement_doc_id1 = docid_arr(0)
	db_settlement_doc_id2= docid_arr(1)

	Set dbRecordSet = Nothing
	Trans_settlement_doc_id = db_settlement_doc_id1
	
	If db_settlement_doc_id1 =  db_settlement_doc_id2 Then
		
		Append_TestHTML StepCounter, "Validate Settlement Document ID", "DB value found. Same Settlement Document ID:"&Trans_settlement_doc_id & VBCRLF & "for transaction", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Settlement Document ID", "DB value not found. different Settlement Document ID1:" & db_settlement_doc_id1 & VBCRLF & "settlement Document ID2: " & db_settlement_doc_id2 & VBCRLF , "FAILED"
	bFlag = False
	
	End If
	
	query3 = "Select * from SettlementDocument where SettlementDocumentID = '" & Trans_settlement_doc_id & "'"
	Append_TestHTML StepCounter, "Get Sales Transaction details", query3, "PASSED"
	Set dbRecordSet = execute_db_query(query3, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_settlement_summary_docid = dbRecordSet("SummarySettlementDocumentID")
	
	db_settlement_docid = dbRecordSet("SettlementDocumentID")
	db_settlement_sessionid = dbRecordSet("SettlementSessionID")
	db_settlement_delco_id = dbRecordSet("DelcoID")
	db_settlement_doc_no = dbRecordSet("SettlementDocumentNumber")
	db_settlement_siteowner_id= dbRecordSet("SiteOwnerID")

	Set dbRecordSet = Nothing
	Trans_settlement_summary_doc_id = db_settlement_summary_docid
	
	If db_settlement_summary_docid <> "" and db_settlement_docid <> "" and db_settlement_sessionid <> "" and db_settlement_delco_id<> "" and  db_settlement_doc_no<> ""  and db_settlement_siteowner_id <> ""Then
		
		Append_TestHTML StepCounter, "Get Summary Document ID-", "DB value found. Summary Document ID:"&db_settlement_summary_docid & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Settlement Document ID-", "DB value found.  Settlement Document ID:"&db_settlement_docid & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Settlement Session ID-", "DB value found.  Settlement Session ID:"&db_settlement_sessionid & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Invoice Customer ID-", "DB value found.  Delco ID:"&db_settlement_delco_id & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Settlement Document Number-", "DB value found. Settlement Document Number:"&db_settlement_doc_no & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Settlement Document Date-", "DB value found. Settlement Siteowner id:"&db_settlement_siteowner_id & "for transaction", "PASSED"
		
	else
		Append_TestHTML StepCounter, "Get Sales Transaction", "DB value not found. different Summary Document ID1:" & db_settlement_summary_docid  , "FAILED"
	bFlag = False
	
	End If
	
	query4 = "Select * from SummarySettlementDocument where SummarySettlementDocumentID = '" & Trans_settlement_summary_doc_id & "'"
	Append_TestHTML StepCounter, "Get Summary Settlement Document details", query4, "PASSED"
	Set dbRecordSet = execute_db_query(query4, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_settlement_summary_docid1 = dbRecordSet("SummarySettlementDocumentID")
	
	db_settlement_sum_doc_no = dbRecordSet("SummarySettlementDocumentNumber")
	db_settlement_sessionid1 = dbRecordSet("SettlementSessionID")


	Set dbRecordSet = Nothing
	Trans_settlement_session_id = db_settlement_sessionid1
	
	If db_settlement_summary_docid1 <> "" and db_settlement_sum_doc_no <> "" and db_settlement_sessionid1 <> "" Then
		
		Append_TestHTML StepCounter, "Get Summary Document ID-", "DB value found. Summary Document ID:"&db_settlement_summary_docid1 & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get SummaryDocumentNumber-", "DB value found.  SummaryDocumentNumber:"&db_settlement_sum_doc_no & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get Settlement Session ID-", "DB value found.  Settlement Session ID:"&db_settlement_sessionid1 & "for transaction", "PASSED"
		
		
	else
		Append_TestHTML StepCounter, "Get Summary Settlement Document deetails", "DB value not found. different Summary Settlement Document ID1:" & db_settlement_summary_docid1  , "FAILED"
	bFlag = False
	
	End If
	query5 = "Select * from SettlementSession where SettlementSessionID = '" & Trans_settlement_session_id & "'"
	Append_TestHTML StepCounter, "Get Summary Documents", query5, "PASSED"
	Set dbRecordSet = execute_db_query(query5, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_settlement_no_summary_docs = dbRecordSet("NoOfSummaryDocuments")
	db_settlement_no_docs = dbRecordSet("NoOfDocuments")
	Set dbRecordSet = Nothing
	
	
	If db_settlement_no_summary_docs <> "" and db_settlement_no_docs <> ""  Then
		
		Append_TestHTML StepCounter, "Get Number of Summary Documents -", "DB value found. No of Summary Documents are :"&db_settlement_no_summary_docs & "for transaction", "PASSED"
		Append_TestHTML StepCounter, "Get number of documents-", "DB value found.  No Of Documents:"&db_settlement_no_docs & "for transaction", "PASSED"
				
	else
		Append_TestHTML StepCounter, "Get Number of Document details", "DB value not found."  , "FAILED"
	bFlag = False
	
	End If
	query6 = "Select * from SettlementSessionoutput where SettlementSessionID = '" & Trans_settlement_session_id & "';"
	Append_TestHTML StepCounter, "Get Job Types for Settlement session", query6, "PASSED"
	Set dbRecordSet = execute_db_query(query6, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_settlement_job_type_id = dbRecordSet("JobTypeID")
	
	
	jobtypeid_arr = split(db_settlement_job_type_id,"|")
	
	db_settlement_jobtype_id1 = jobtypeid_arr(0)
	db_settlement_jobtype_id2 = jobtypeid_arr(1)
	
	db_settlement_job_id = dbRecordSet("JobID")
	jobid_arr = split(db_settlement_job_id,"|")
	
	db_settlement_job_id1 = jobid_arr(0)
	db_settlement_job_id2 = jobid_arr(1)
	
	Set dbRecordSet = Nothing
	
	If db_settlement_jobtype_id1 = "289" Then
		db_289_job_id = db_settlement_job_id1
		
	ElseIf db_settlement_jobtype_id2 = "289" Then
		db_289_job_id = db_settlement_job_id2
	Else
		db_289_job_id = ""
	End If
	If db_settlement_jobtype_id1 = "380" Then
		db_380_job_id = db_settlement_job_id1
		
	ElseIf db_settlement_jobtype_id2 = "380" Then
		db_380_job_id = db_settlement_job_id2
	Else
		db_380_job_id = ""
	End If
	
	
	
	If db_settlement_jobtype_id1 <> "" and db_settlement_jobtype_id2 <> ""  Then
		
		Append_TestHTML StepCounter, "Job Triggered Type details-", "DB value found. Job TypeID1 :"&db_settlement_jobtype_id1& VBCRLF&"job TypeID2 :"& db_settlement_jobtype_id2 & "for this transaction", "PASSED"
			
	else
		Append_TestHTML StepCounter, "Job Triggered Type details", "DB value not found."&db_settlement_jobtype_id1& VBCRLF&"job TypeID2 :"& db_settlement_jobtype_id2 & "for this transaction"  , "FAILED"
	bFlag = False
	
	End If
	If db_settlement_job_id1 <> "" and db_settlement_job_id2 <> ""   Then
		
		Append_TestHTML StepCounter, "Job IDs -", "DB value found. Job ID1 :"&db_settlement_job_id1& VBCRLF&"job ID2 :"&db_settlement_job_id2& VBCRLF &  "for this transaction", "PASSED"
		Append_TestHTML StepCounter, "Job IDs -", "DB value found. Job ID1 :"&db_289_job_id& VBCRLF&"job ID2 :"&db_380_job_id& VBCRLF &  "for this transaction", "PASSED"
		
			
	else
		Append_TestHTML StepCounter, "Job Ds", "DB value not found."&db_settlement_job_id1& VBCRLF&"job ID2 :"&db_settlement_job_id2& VBCRLF &  "for this transaction"  , "FAILED"
	bFlag = False
	
	End If
Else
		Append_TestHTML StepCounter, "Get Settlement document details for file validations", "Unable to get the details because no settlement and job not initiated"  , "FAILED"
	bFlag = False

	
End If	
End Function



Function validateJobonestatus()
On error resume next
	bFlag = True
bRunFlag = True
	
	query = "Select * from Job where JobTypeID = '1' order by 1 desc;"
	Append_TestHTML StepCounter, "Validate Job details", query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_job_statusID = dbRecordSet("StatusID")
	db_outputfile = dbRecordSet("InputXml")
	db_generated_date = dbRecordSet("DateCreated")
	
	Set dbRecordSet = Nothing
	
	If db_job_statusID = "3"  or  instr(db_outputfile,NULL) > 0  Then
		
		Append_TestHTML StepCounter, "Validate Status and inputxml -", "DB value matched successfully. Job Status completed with status as:- "&db_job_statusID & " inputxml file:-"& db_outputfile , "PASSED"
					
	else
		Append_TestHTML StepCounter, "Validate Status and inputxml -", "DB values are not matched. Job Status  with Null values/No entry:- "&db_job_statusID & " inputxml file:-"& db_outputfile , "FAILED"
	bFlag = False
	bRunFlag = False
		
	End If
End Function

Function validatenormalJobStatus(jobtypeid,jfilename)
On error resume next
bRunFlag = True
				bFlag = True
	query = "Select * from Job where JobTypeID = '" & jobtypeid & "';"
	Append_TestHTML StepCounter, "Validate Job details", query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_job_statusID = dbRecordSet("StatusID")
	db_outputfile = dbRecordSet("InputXml")
	db_generated_date = dbRecordSet("DateCreated")
	
	Set dbRecordSet = Nothing
	
	If db_job_statusID = "4"  or  instr(db_outputfile,jfilename) > 0  Then
		
		Append_TestHTML StepCounter, "Validate Status and inputxml -", "DB value matched successfully. Job Status completed with status as:- "&db_job_statusID & " inputxml file:-"& db_outputfile , "PASSED"
					
	else
		Append_TestHTML StepCounter, "Validate Status and inputxml -", "DB values not matched. Job Status completed with status as:- "&db_job_statusID & " inputxml file:-"& db_outputfile , "FAILED"
		bRunFlag = False
				bFlag = False
	End If
End Function

Function validateJobStatus(jobtypeid,jfilename)
On error resume next
bRunFlag = True
				bFlag = True
If jobtypeid = "211" Then
	db_job_id = db_211_job_id
ElseIf jobtypeid = "264" Then
	db_job_id = db_264_job_id
ElseIf jobtypeid = "305" Then
	db_job_id = db_305_job_id
ElseIf  jobtypeid = "289" Then
	db_job_id = db_289_job_id
ElseIf  jobtypeid = "380" Then
	db_job_id = db_380_job_id
End If
If jobtypeid <> "" and  db_job_id <> "" Then
	query = "Select * from Job where JobTypeID = '" & jobtypeid & "' and ID = '"& db_job_id &"'"
	Append_TestHTML StepCounter, "Validate Job details", query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_job_statusID = dbRecordSet("StatusID")
	db_outputfile = dbRecordSet("InputXml")
	db_generated_date = dbRecordSet("DateCreated")
	
	Set dbRecordSet = Nothing
	
	If db_job_statusID = "4" or  instr(db_outputfile,jfilename) > 0  Then
		
		Append_TestHTML StepCounter, "Validate Status and inputxml -", "DB value matched successfully. Job Status completed with status as:- "&db_job_statusID & " inputxml file:-"& db_outputfile , "PASSED"
			validateJobStatus = True		
	else
		Append_TestHTML StepCounter, "Validate Status and inputxml -", "DB values not matched. Job Status completed with status as:- "&db_job_statusID & " inputxml file:-"& db_outputfile , "FAILED"
		validateJobStatus = False
		bRunFlag = False
		bFlag = False
	End If
Else
		Append_TestHTML StepCounter, "Verify " & jobtypeid & "status", "Unable to find the job id for DB validation jobid-"& db_job_id , "FAILED"
bRunFlag = False
		bFlag = False
End If
End Function

'******************************* HEADER ******************************************
' Description : 
' Creator : 
' Date : 
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Function verifyFileExistance(filePath,filetype)
On error resume next
	bFlag = True
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	
	If fileSysObj.FolderExists(filePath)Then
	
'		'''''msgbox "Exist"
		Set sFolder = fileSysObj.GetFolder(filePath)
		Set oFileColl = sFolder.Files
		Set RecentFile = Nothing
	
		fileLocation = ""
		For each  otmpFile in oFileColl
'			'''''msgbox otmpFile.Name
'			'''''msgbox otmpFile.DateLastModified
'			'''''msgbox  otmpFile.Type
			name = otmpFile.Name
			dateLModified = otmpFile.DateLastModified
			ftype = otmpFile.Type
				
			If (instr(ucase(name),ucase(filetype))>0 or  instr(otmpFile.Type,filetype)>0 )and instr(otmpFile.DateLastModified,JobRundate )>0 Then
				Print name
				If RecentFile is Nothing Then
					Set RecentFile = otmpFile
					fileLocation = filePath & "\" & name
					
				ElseIf otmpFile.DateLastModified>=RecentFile.DateLastModified Then
					Set RecentFile = otmpFile
					fileLocation = filePath & "\" & name
				End If			
			End If
'			Exit for
		Next
		
		If fileLocation<>"" Then
			''''''msgbox fileLocation
			Append_TestHTML StepCounter, "File Identification", "Identified file in the path:- "&fileLocation , "PASSED"
			
			verifyFileExistance = fileLocation
		Else
			'Append_TestHTML StepCounter, "File Identification", "File not generated in the source path:- "&filePath , "FAILED"
			verifyFileExistance=""		
	'bFlag = False
			
		End If
		
	Else
			Append_TestHTML StepCounter, "Folder path Identification", "Folder path not exist- "&filePath , "FAILED"
			verifyFileExistance=""
	bRunFlag = False
		bFlag = False	
	End If
End Function

'******************************* HEADER ******************************************
' Description : 
' Creator : 
' Date : 
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateFile(filePath,data)
       On error resume next
       
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(filePath) Then
		Set fileRead = fileSysObj.OpenTextFile(filePath,1)
		content = fileRead.ReadAll
		If instr(data,";") = 0 Then
			If instr(content,data)>0 Then
				Append_TestHTML StepCounter, "File data validation", "Data " & data & "matched in the path "&filePath , "PASSED"
				
			Else				Append_TestHTML StepCounter, "File data validation", "Data " & data & "not matched in the path "&filePath , "FAILED"
			
			End If
		Else
		dataval = split(data,";")
			For iitr = 0 To ubound(dataval) Step 1
				If instr(content,dataval(iitr))>0 Then
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "matched in the path "&filePath , "PASSED"
					
				Else
					Append_TestHTML StepCounter, "File data validation", "Data " & dataval(iitr) & "not matched in the path "&filePath , "FAILED"
				
				End If
			Next
		End If
	End If
End Function


Public Function validateFiledata(filePath,data,ftype)
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
	End If
End Function



Public Function validateCardStatus(card_no,status)

	On error resume next
	
	bFlag = True
'	wait 3   ' waiting for the DB update
	If card_no <> "" Then
		query = "select * from Card where PAN="& card_no &" order by 1 desc;"
		Append_TestHTML StepCounter, "Card Startus query", query, "PASSED"
		
		set dbRecordSet =execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		cardidno = dbRecordSet("CardID")
		statusid = dbRecordSet("StatusID")
		''''''msgbox cardidno
		set dbRecordSet = Nothing
		If cardidno <> ""  and cint(statusid) = cint(status) Then
			If cint(status) = cint("10") Then
				Append_TestHTML StepCounter,"Validate Card "& card_no &" Status ","Card status matched with new status and Card ID is: " & cardidno &VBCRLF & "and status is: " & statusid   ,"PASSED"
			ElseIf cint(status) =cint("1") Then	
				Append_TestHTML StepCounter,"Validate Card "& card_no &" Status ","Card status matched with Active status and Card ID is: " & cardidno &VBCRLF & "and status is: " & statusid   ,"PASSED"
			ElseIf cint(statusid) =cint(status) Then	
				Append_TestHTML StepCounter,"Validate Card "& card_no &" Status ","Card status matched with Active status and Card ID is: " & cardidno &VBCRLF & "and status is: " & statusid   ,"PASSED"
			Else
				Append_TestHTML StepCounter,"Validate Card "& card_no &" Status ","Card status matched with Active status and Card ID is: " & cardidno &VBCRLF & "and status Expected: " & status & VBCRLF & "Actual : " & statusid  ,"FAILED"
			bRunFlag = False
			
			End If
		
		Else
			Append_TestHTML StepCounter,"Validate Card "& card_no &" Status ","Card status not matched and Card ID is: " & cardidno &VBCRLF & "and status Expected: " & status & VBCRLF & "Actual : " & statusid  ,"FAILED"
			bRunFlag = False
			
		End If
		validateCardStatus = cardidno	
	Else
			Append_TestHTML StepCounter,"Validate Card "& card_no &" Status ","Card not exist and unable to do status validation"  ,"FAILED"
			bRunFlag = False
	End If
	
End Function



Public Function RunJob196()

	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 196 and StatusID=0 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 196", query, "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	NextDate = dbRecordSet("NextRunDate")
	Nextrun = FormatDateTime(NextDate,2)
	
	Set dbRecordSet = Nothing
	If db_id <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 196", "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 196", "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "FAILED"
	End If
	If db_status_id = 0 Then
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is not successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
	End If
	If Nextrun = FormatDateTime(Date,2) OR Nextrun = FormatDateTime(Date+1,2)  Then
		Append_TestHTML StepCounter, "Validated Next date validation", "Validation of status ID before update is successful. Expected Value: "& Nextrun  & VBCRLF & "Actual Value: " & NextDate, "PASSED"
	End  If
	query1 = "Update Job Set NextRunDate = null where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Update Job 196", "Update Job Set NextRunDate = null where ID = '" & db_id & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	Set dbRecordSet = Nothing
	wait 10
	query2 = "Select * from Job where JobTypeId = 196 and ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 196",query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_status_id = dbRecordSet("StatusID")
	
	If db_status_id <> "" Then
		
		Append_TestHTML StepCounter, "Verify Job 196", "Job ID-"& dbRecordSet("ID") & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Verify Job 196", "DB values not found. Status ID: " & db_status_id, "FAILED"
	End If
	If db_status_id = 4 Then
		Append_TestHTML StepCounter, "Validate Job Status", "Validation of status ID after update is successful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
		JobRundate = Date()
	else
		Append_TestHTML StepCounter, "Validate Job Status", "Validation of status ID after update is not successful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
		bRunFlag = False
	End If
	Set dbRecordSet = Nothing
End  Function

Public Function createandmoveDX026File(fPath,dPath,fileInitialName,cardNum,cardExpdate)

On error resume next	

If cardPANNum <> ""  Then
	
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(fPath) Then
		'fileInitialName = "DX026_GFN_TRX_03601_000094"
		dtvalue = getDateandTimestamp()
		dtvalue = Replace(dtvalue," ","_")
		dtvalue = Replace(dtvalue,"-","")
		nFname =  fileInitialName&"_"&dtvalue & ".dat"
		newfile = dPath & fileInitialName&"_"&dtvalue & ".dat"
		
		fileSysObj.CreateTextFile(newfile)
		Set DXread = fileSysObj.OpenTextFile(fPath,1)
		content = DXread.ReadAll
		DXread.Close
		
		DXarray = Split(content,"|")
		If cardPANNum = "" Then
			cardPANNum = cardNum
		End If
		If cardExpiry_date = "" Then
			cardExpiry_date = cardExpdate
		End If
		datebuildval = split(cardExpiry_date,"-")
		expdate = datebuildval(1) &  right(datebuildval(0),2)
		
		currentdatetimeval = getDateandTimestamp()
		
		ntimeval = getDateandTimestamp()
		nstimeval = split(ntimeval," ")
		newtimeval = nstimeval(1)
			sRandNo = Int((10 - 1 + 1) * Rnd + 1)
'			DX26SessionBillDate
'			DX26SessionSettleDate
'			If cdate(DX26SessionBillDate) <= cdate(DX26SessionSettleDate) Then
'				DX26SessionDate = DX26SessionSettleDate
'			else

				DX26SessionDate = DX26SessionBillDate
'			End If
		newsessiondateval = cdate(DX26SessionDate) +	sRandNo
		
		nd = day(newsessiondateval)
		nm = month(newsessiondateval)
		ny = year(newsessiondateval)
		If nd<10 Then
			nd = "0"&nd
		End If
		If nm<10 Then
			nm = "0"& nm
		End If
		newsessiondateval = ny&nm&nd
		salesdateval = newsessiondateval & newtimeval
		
		
		'''msgbox salesdateval
		'salesdateval="20211114172326"
		existdatelen = len(salesdateval)
		existdateval = mid(salesdateval,1,existdatelen-6)
		curredateval = split(currentdatetimeval," ")
		changedateval1 = existdateval & curredateval(1)
		
		
'		changedateval2 = changedateval1+1
'		vocherval1 = curredateval(1)
'		vocherval2= curredateval(1) + 1
		timedataval = curredateval(1)
		secondval =  right(timedataval,2)
		hmval =  left(timedataval,4)
		
		If secondval >= 59 Then
			secondsval =  secondval - 1
		Else
			secondsval =  secondval + 1
		End If
		If len(secondsval) = 1 Then
			secondsval = "0"& secondsval
		End If
		currenttimeval = hmval & secondsval
		changedateval2 = existdateval & currenttimeval
		vocherval1 = curredateval(1)
		vocherval2= currenttimeval
	
		modifyData = "10-"&cardPANNum&";11-"& expdate & ";12-"& changedateval1 &";22-"& vocherval1 & ";41-"& cardPANNum&";42-"& expdate & ";43-"& changedateval2 &";53-"& vocherval2
	
		Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
		sdata = Split(modifyData,";")
		Newcontent = content
		For itr = 0 To ubound(sdata) Step 1
			datavalue = sdata(itr)
			spos =Split(datavalue,"-")(0)
			svalue= Split(datavalue,"-")(1)
			Newcontent = Replace(Newcontent, DXarray(spos),svalue)
		Next
		Newcontent = Replace(Newcontent, DXarray(63),nFname)
		DXwrite.Write Newcontent
		DXwrite.Close
		Set fileSysObj = Nothing
		JobRundate=date()
	End If
	If nFname <> "" Then
		Append_TestHTML StepCounter, "DX026 File creation and move to GLobal location", "File created at"& newfile , "PASSED"
		dx026FinalFName = nFname
		createandmoveDX026File = nFname
	Else
		createandmoveDX026File = ""
	End If

Else
	Append_TestHTML StepCounter, "DX026 File creation and move to GLobal location", "Unable to crete a File without any customer card" , "FAILED"
End If	
End Function

Function verifyDX026FileStatus(inboundfilepath,DX26newFilename)
On error resume next	

	For itr = 1 To 15 Step 1
		wait 5
		fileopstatus = verifyFileExistance(inboundfilepath,DX26newFilename)
		If fileopstatus = "" Then
				
			Exit for
		Else
			wait 5
		End If
	Next
	
	If fileopstatus = "" Then
		Append_TestHTML StepCounter, "Verify File in Global Folder", DX26newFilename&"File got processed from "&inboundfilepath , "PASSED"
		successfolderpath = inboundfilepath & "\Archive"
		Errorfolderpath = inboundfilepath & "\Error"
		
		fileopstatus = verifyFileExistance(successfolderpath,DX26newFilename)
		
		If fileopstatus<> "" Then
		Append_TestHTML StepCounter, "Verify File in Archive Folder", "File Successfully Identified in the path:- "&fileopstatus & " in Archive", "PASSED"
		Else
			fileopstatus = verifyFileExistance(Errorfolderpath,DX26newFilename)
			If fileopstatus <> "" Then
				Append_TestHTML StepCounter, "Verify File in Error Folder", "File having errors and moved to path:- "&fileopstatus & "in Error" , "FAILED"
				bRunFlag = FALSE
			End If
		End If
	Else
				Append_TestHTML StepCounter, "Verify File in Folder", "File not processed"&inboundfilepath&"\"&DX26newFilename , "FAILED"
				bRunFlag = FALSE
	End If
End Function



Public Function VerifySettlementTxn(settlement_report_date, PAN_num)
'''''msgbox settlement_report_date

	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
		'PAN_num = "7002851000050000016"
	End If
	
''	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_BillPreview") = False Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Start", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Transactions", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_BillAcceptance", "Click", ""
		
		wait 10
'	Else
	
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebLink_Actions", "Click", ""
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebLink_SettlementAceeptance", "Click", ""
		wait 12
		
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_SettlementPreview") Then
			Append_TestHTML StepCounter, "Settlement Prview Acceptance", "Navigated to Settlement Transactions page", "PASSED"
		else
			Append_TestHTML StepCounter, "Settlement Prview Acceptance", "Navigation to Settlement Transactions page failed", "FAILED"
			bRunFlag = False
			bFlag = False
		End If
'			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_SettlementPreview", "Click", ""
'		''msgbox  "Next will Settlement Preview"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlPreview", "Click", ""
		wait 4
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
			Append_TestHTML StepCounter, "Settellement Preview", "Clicked on Setllement preview", "PASSED"
		else
			Append_TestHTML StepCounter, "Settelement Preview", "Clicking on settlement preview button failed", "FAILED"
		End If
		
		Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_StlRefresh") = False
			wait 1
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlRefresh", "Click", ""
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
				Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
			else
'				Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
			preview_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledPreview", "GetROProperty", "innertext")
				breqreportdate = reqdateFormt()
				If instr(preview_msg, "Settlement previewed") <> 0  or  instr(preview_msg, breqreportdate) <> 0 Then
					Append_TestHTML StepCounter, "Validate Settlement preview", "Settlement previewed successfully. Expected Value: " & settll_report_date_req & VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "PASSED"
				else
					Append_TestHTML StepCounter, "Validate Settlement preview", "Settlement preview failed. Expected Value: " & settll_report_date_req & VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "FAILED"
					bRunFlag = False
						bFlag = False
					
				End If
			End If	
		Loop
		settll_report_date_req = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_SettlementReportdate", "GetROProperty", "default value")
		''''msgbox settll_report_date_req
		settl_report_date_arr = split(settll_report_date_req,"/")
		settl_report_date = settl_report_date_arr(2) & "-" & settl_report_date_arr(1) & "-" & settl_report_date_arr(0)
		settllReportDate = settl_report_date
		'''''msgbox settllReportDate
		wait 20
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_SettlementReportdate") Then
			Append_TestHTML StepCounter, "Settlement report date", "Captured Settlement report date displayed: "& settl_report_date, "PASSED"
		else
			Append_TestHTML StepCounter, "Settlement report date", "Settlement report date not found", "FAILED"
				bRunFlag = False
				bFlag = False
		End If
		
	If settl_report_date<> "" and Trans_BatchID <> "" Then
		
		'Perform DB operations
			query = "Select * from SalesItemDelcoUnsettled where batchid = '" & Trans_BatchID & "'"
			Append_TestHTML StepCounter, "Check actual Settlement date",query, "PASSED"
			Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
			wait 2
			db_actual_report_date = dbRecordSet("SettlementReportDate")
			date_arr = split(db_actual_report_date,"|")
			db_actual_report_date1 = date_arr(0)
			db_actual_report_date2 = date_arr(1)
			Set dbRecordSet = Nothing
			If db_actual_report_date1 <> "" and db_actual_report_date2 <> "" Then
				Append_TestHTML StepCounter, "Actual Settlement ReportDate Report date", "Actual Settlement report date1: " & db_actual_report_date1 & VBCRLF & "Actual Settlement report date2: " & db_actual_report_date2, "PASSED"
			else
				Append_TestHTML StepCounter, "Actual Settlement Report date", "DB values not found. Actual Settlement report date1: " & db_actual_report_date1 & VBCRLF & "Actual Settlement report date2: " & db_actual_report_date2, "FAILED"
			bRunFlag = False
				bFlag = False
			End If
			
			query1 = "Select * from Job where JobTypeId = 19 order by 1 desc;"
			Append_TestHTML StepCounter, "Run Job 19", "Select * from Job where JobTypeId = 19 order by 1 desc;", "PASSED"
			Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
			wait 2
			db_id_19 = dbRecordSet("ID")
			db_status_id_19 = dbRecordSet("StatusID")
			Set dbRecordSet = Nothing
			If db_id_19 <> "" and db_status_id_19 <> "" Then
				Append_TestHTML StepCounter, "Check Job 19 ID", "ID: " & db_id_19 & VBCRLF & "Status: " & db_status_id_19, "PASSED"
			else
				Append_TestHTML StepCounter, "Check Job 19 ID", "DB values not found. ID: " & db_id_19 & VBCRLF & "Status: " & db_status_id_19, "FAILED"
				bRunFlag = False
				bFlag = False
			
			End If
			
			query2 = "Select * from InfoProviderNextRunDate where InfoProviderID = 8"
			Append_TestHTML StepCounter, "Check InfoProviderNextRunDate", "Select * from InfolProviderNextRunDate where InfoProviderID = 8", "PASSED"
			Set dbRecordSet = execute_db_query(query2, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
			wait 2
			db_next_report_date = dbRecordSet("NextReportDate")
			report_date_arr = split(db_next_report_date,"|")
			db_next_report_date1 = report_date_arr(0)
			db_next_report_date2 = report_date_arr(1)
			db_next_run_date = dbRecordSet("NextRunDate")
			run_date_arr = split(db_next_run_date,"|")
			db_next_run_date1 = run_date_arr(0)
			db_next_run_date2 = run_date_arr(1)
			Set dbRecordSet = Nothing
			If db_next_report_date1 <> "" and db_next_report_date2 <> ""  and db_next_run_date1 <> "" and db_next_run_date2 <> "" Then
				Append_TestHTML StepCounter, "Check dates in InfoProviderNextRunDate", "Next Report Date1: " & db_next_report_date1 & VBCRLF & "Next Report Date2: " & db_next_report_date2 & VBCRLF & "Next Run Date1: " & db_next_run_date1 & VBCRLF & "Next Run Date2: " & db_next_run_date2, "PASSED"
			else
				bRunFlag = False
			
			bFlag = False
				Append_TestHTML StepCounter, "Check dates in InfoProviderNextRunDate", "DB values not found. Next Report Date1: " & db_next_report_date1 & VBCRLF & "Next Report Date2: " & db_next_report_date2 & VBCRLF & "Next Run Date1: " & db_next_run_date1 & VBCRLF & "Next Run Date2: " & db_next_run_date2, "FAILED"
			
			End If
			
			query3 = "Select * from SettlementAcceptance order by 1 desc;"
			Append_TestHTML StepCounter, "Check SettlementAcceptance", "Select * from SettlementAcceptance order by 1 desc;", "PASSED"
			Set dbRecordSet = execute_db_query(query3, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
			wait 2
			db_settllement_id = dbRecordSet("SettlementAcceptanceID")
			db_settlement_report_date = dbRecordSet("SettlementReportDate")
			Set dbRecordSet = Nothing
			If db_settllement_id <> "" and db_settlement_report_date <> "" Then
				Append_TestHTML StepCounter, "Check SettlementAcceptance", "Settlement ID: " & db_settllement_id & VBCRLF & "Settlement Report Date: " & db_settlement_report_date, "PASSED"
			else
				Append_TestHTML StepCounter, "Check SettlementAcceptance", "DB values not found. Settlement ID: " & db_settllement_id & VBCRLF & "Settlement Report Date: " & db_settlement_report_date, "FAILED"
			bFlag = False
				bRunFlag = False
			
			End If
			
'			query4 = "Update SalesItemDelcoUnsettled Set SettlementReportDate = '" & db_settlement_report_date & "',SettlementDiscountReportDate ='"& db_settlement_report_date &"' where batchid = '" & Trans_BatchID & "'"
'			Append_TestHTML StepCounter, "Update SettlementReportDate in SalesItemDelcoUnsettled", query4, "PASSED"
'			Set dbRecordSet = execute_db_query(query4 ,1,"SFN_SHELL_SPRINTQA_ID_BATCH")
'			wait 2
'			Set dbRecordSet = Nothing
			
			Set dbRecordSet = execute_db_query(query, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
			Append_TestHTML StepCounter, "Check updated Settlement date", "Select * from SalesItemDelcoUnsettled where batchid = '" & Trans_BatchID & "'", "PASSED"
			wait 2
			db_updated_report_date = dbRecordSet("SettlementReportDate")
			date_updated_arr = split(db_updated_report_date,"|")
			db_updated_report_date1 = date_updated_arr(0)
			db_updated_report_date2 = date_updated_arr(1)
			Set dbRecordSet = Nothing
			If db_updated_report_date1 <> "" and db_updated_report_date2 <> "" Then
				Append_TestHTML StepCounter, "Updated Settlement Report date", "Updated Settlement report date1: " & db_updated_report_date1 & VBCRLF & "Updated Settlement report date2: " & db_updated_report_date2, "PASSED"
			else
				Append_TestHTML StepCounter, "Updated Settlement Report date", "DB values not found. Updated Settlement report date1: " & db_updated_report_date1 & VBCRLF & "Updated Settlement report date2: " & db_updated_report_date2, "FAILED"
			bFlag = False
				bRunFlag = False
			
			End If
			
			If instr(db_updated_report_date1, settl_report_date) <> 0 Then
				Append_TestHTML StepCounter, "Validate Settlement Report date", "Settlement report date1 is updated successfully. Expected Value: " & db_updated_report_date1 & VBCRLF & "Actual Value: " & settl_report_date, "PASSED"
			else
				Append_TestHTML StepCounter, "Validate Settlement Report date", "Settlement report date1 update failed. Expected Value: " & db_updated_report_date1 & VBCRLF & "Actual Value: " & settl_report_date, "FAILED"
			bFlag = False
				bRunFlag = False
			
			End If
			
'			If instr(db_updated_report_date2, settl_report_date) <> 0 Then
'				Append_TestHTML StepCounter, "Validate Settlement Report date", "Settlement report date2 is updated successfully. Expected Value: " & db_updated_report_date2 & VBCRLF & "Actual Value: " & settl_report_date, "PASSED"
'			else
'				Append_TestHTML StepCounter, "Validate Settlement Report date", "Settlement report date2 update failed. Expected Value: " & db_updated_report_date2 & VBCRLF & "Actual Value: " & settl_report_date, "FAILED"
'				bRunFlag = False
'			End If
			'----------------Updating InfoProviderNextrundate table data	
				
				settl_next_report_date = nextdateDBFormt(settl_report_date)
				
'				query_Info_next_run = "update InfoProviderNextRunDate set NextReportDate='"& settl_report_date &"',NextRunDate='"& settl_next_report_date &"' where infoproviderid = 8;"
'				Append_TestHTML StepCounter, "Update NextReportdate in InfoProviderNextRunDate", query_Info_next_run, "PASSED"
'				Set dbRecordSet = execute_db_query(query_Info_next_run ,1,"SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				Set dbRecordSet = Nothing
				query_Info_run = "Select * from InfoProviderNextRunDate where infoproviderid = 8;"
				Set dbRecordSet = execute_db_query(query_Info_run, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
				Append_TestHTML StepCounter, "Check updated reported dates", query_Info_run , "PASSED"
				wait 2
				db_Info_report_date = dbRecordSet("NextReportDate")
				Info_report_date = split(db_Info_report_date,"|")
				db_info_report_date1 = Info_report_date(0)
				db_info_report_date2 = Info_report_date(1)
				
					db_Info_nrun_date = dbRecordSet("NextRunDate")
				Info_report_nrdate = split(db_Info_nrun_date,"|")
				db_info_report_ndate1 = Info_report_nrdate(0)
				db_info_report_ndate2 = Info_report_nrdate(1)
				
				Set dbRecordSet = Nothing
				If db_info_report_date1 <> "" and db_info_report_date2 <> "" Then
					Append_TestHTML StepCounter, "Updated next Report date", "Updated next Info report date1: " & db_info_report_date1 & VBCRLF & "Updated next info report date2: " & db_info_report_date2, "PASSED"
				else
					Append_TestHTML StepCounter, "Updated next Report date", "DB values are not Updated. next Info report date1: " & db_info_report_date1 & VBCRLF & "Updated next info report date2: " & db_info_report_date2, "FAILED"
				bFlag = False
				bRunFlag = False
				
				End If
				
				If instr(db_info_report_date1, settl_report_date) <> 0 Then
					Append_TestHTML StepCounter, "Validate next Report date", "Next report date1 is updated successfully. Expected Value: " & db_info_report_date1 & VBCRLF & "Actual Value: " & settl_report_date, "PASSED"
				else
					Append_TestHTML StepCounter, "Validate next Report date", "Next report date1 update failed. Expected Value: " & db_info_report_date1 & VBCRLF & "Actual Value: " & settl_report_date, "FAILED"
					bFlag = False
				bRunFlag = False
				
				End If
'				If instr(db_info_report_ndate2,settl_next_report_date) <> 0 Then
'					Append_TestHTML StepCounter, "Validate next Report run ate", "Next run report date is updated successfully. Expected Value: " & db_info_report_ndate2 & VBCRLF & "Actual Value: " & bill_next_report_date, "PASSED"
'				else
'					Append_TestHTML StepCounter, "Validate next Report run date", "Next run report date update failed. Expected Value: " & db_info_report_ndate2 & VBCRLF & "Actual Value: " & bill_next_report_date, "FAILED"
'					bFlag = False
'				bRunFlag = False
'						
'				End If
			'-------------------------------------------	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_SettlementPreview", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
		Append_TestHTML StepCounter, "Settllement Preview", "Clicked on Settllement preview", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Settllement Preview", "Clicking on Settllement preview button failed", "FAILED"
'			bFlag = False
'				bRunFlag = False
	
	End If
	
	Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_StlRefresh") = False
		wait 1
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlRefresh", "Click", ""
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
			Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
		else
'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
			preview_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledPreview", "GetROProperty", "innertext")
				breqreportdate = reqdateFormt()
				If instr(preview_msg, "Settlement previewed") <> 0  or  instr(preview_msg, breqreportdate) <> 0 Then
					Append_TestHTML StepCounter, "Validate Settlement preview", "Settlement previewed successfully. Expected Value: " & settll_report_date_req & VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "PASSED"
				else
					Append_TestHTML StepCounter, "Validate Settlement preview", "Settlement preview failed. Expected Value: " & settll_report_date_req & VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "FAILED"
					bRunFlag = False
						bFlag = False
					
				End If
		End If
	Loop
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Company") Then
		Append_TestHTML StepCounter, "Validate Txns", "Transactions updated with bill report date", "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Txns", "Transaction details not displayed", "FAILED"
			bFlag = False
	
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledFees", "Click", ""
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Currency") Then
		Append_TestHTML StepCounter, "Validate Fee details", "Fee details updated withSettled report date", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Validate Fee details", "Fee details not displayed", "FAILED"
'			bFlag = False
'				bRunFlag = False
	
	End If
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledTxns", "Click", ""
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlCutoff", "Click", ""
	wait 4
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
		Append_TestHTML StepCounter, "Cutoff", "Clicked on cutoff", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Cutoff", "Clicking on cutoff failed", "FAILED"
'			bFlag = False
'				bRunFlag = False
	
	End If
	
	'add refresh in between and validate side text
	Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_StlRefresh") = False
		wait 1
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlRefresh", "Click", ""
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
			Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
		else
'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
			cutoff_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledCutoff", "GetROProperty", "innertext")
				
			If instr(cutoff_msg, "Settlement cut off") <> 0 or  instr(cutoff_msg, breqreportdate) <> 0 Then
				Append_TestHTML StepCounter, "Validate Settlement cutoff", "Settlement cutoff successful. Expected Value: " & settll_report_date_req& VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & cutoff_msg, "PASSED"
			else
				Append_TestHTML StepCounter, "Validate Settlement cutoff", "Settlement cutoff failed. Expected Value: " & settll_report_date_req& VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & cutoff_msg, "FAILED"
				bRunFlag = False
					bFlag = False
				
			End If
		End If
	Loop
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlSignoff", "Click", ""
	wait 4
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_OK") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "NoFrame", "WebElement_OK", "Click", ""
		Append_TestHTML StepCounter, "Signoff", "Clicked on signoff", "PASSED"
	else
'		Append_TestHTML StepCounter, "Signoff", "Clicking on signoff failed", "FAILED"
	End If
'	
	'''''msgbox "singoff alert"
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
		Apend_TestHTML StepCounter, "Signoff", "Signoff started", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Signoff", "Signoff not started", "FAILED"
'			bFlag = False
'	
	End If
	
	'again refresh and validate
	Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_StlRefresh") = False
		wait 1
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_StlRefresh", "Click", ""
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebButton_StlRefresh") Then
			Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
'		else
'			Append_TestHTML StepCounter, "Refresh", "Clicking on refresh button failed", "FAILED"
		End If
	Loop
	
	preview_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledPreview", "GetROProperty", "innertext")
	breqreportdate = reqdateFormt()
	If instr(preview_msg, "Settlement previewed") <> 0  or  instr(preview_msg, breqreportdate) <> 0 Then
		Append_TestHTML StepCounter, "Validate Settlement preview", "Settlement previewed successfully. Expected Value: " & settll_report_date_req & VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Settlement preview", "Settlement preview failed. Expected Value: " & settll_report_date_req & VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & preview_msg, "FAILED"
		bRunFlag = False
			bFlag = False
		
	End If
	
	cutoff_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledCutoff", "GetROProperty", "innertext")
		
	If instr(cutoff_msg, "Settlement cut off") <> 0 or  instr(cutoff_msg, breqreportdate) <> 0 Then
		Append_TestHTML StepCounter, "Validate Settlement cutoff", "Settlement cutoff successful. Expected Value: " & settll_report_date_req& VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & cutoff_msg, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Settlement cutoff", "Settlement cutoff failed. Expected Value: " & settll_report_date_req& VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & cutoff_msg, "FAILED"
		bRunFlag = False
			bFlag = False
		
	End If
	
	signoff_msg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_SettledSignoff", "GetROProperty", "innertext")

	If instr(signoff_msg, "Settlement signed off") <> 0 or  instr(signoff_msg, breqreportdate) <> 0 Then
		Append_TestHTML StepCounter, "Validate Settlement signoff", "Settlement signoff successful. Expected Value: " & settll_report_date_req& VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & signoff_msg, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Settlement signoff", "Settlement signoff failed. Expected Value: " & settll_report_date_req& VBCRLF& breqreportdate & VBCRLF & "Actual Value: " & signoff_msg, "FAILED"
		bRunFlag = False
			bFlag = False
		
	End If
	
	next_preview_date = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_SettledNextPreview", "GetROProperty", "value")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_SettledNextPreview") Then
		Append_TestHTML StepCounter, "Validate Next preview date", "Next preview date displayed: " & next_preview_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Next preview date", "Next preview date not displayed: " & next_preview_date, "FAILED"
		bRunFlag = False
			bFlag = False
		
	End If
	
	'DB operations	
	Set dbRecordSet = execute_db_query(query2, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	Append_TestHTML StepCounter, "Check updated InfoProviderNextRunDate", "Select * from InfoProviderNextRunDate where InfoProviderID = 8", "PASSED"
	wait 2
	db_next_report_date1 = dbRecordSet("NextReportDate")
	report_date_arr1 = split(db_next_report_date1,"|")
	db_next_report_date3 = report_date_arr1(0)
	db_next_report_date4 = report_date_arr1(1)
	db_next_run_date = dbRecordSet("NextRunDate")
	run_date_arr1 = split(db_next_run_date,"|")
	db_next_run_date3 = run_date_arr1(0)
	db_next_run_date4 = run_date_arr1(1)
	Set dbRecordSet = Nothing
	If db_next_report_date3 <> "" and db_next_report_date4 <> ""  and db_next_run_date3 <> "" and db_next_run_date4 <> "" Then
		Append_TestHTML StepCounter, "Check updated dates in InfoProviderNextRunDate", "Updated Next Report Date1: " & db_next_report_date3 & VBCRLF & "Updated Next Report Date2: " & db_next_report_date4 & VBCRLF & "Updated Next Run Date1: " & db_next_run_date3 & VBCRLF & "Updated Next Run Date2: " & db_next_run_date4, "PASSED"
	else
		Append_TestHTML StepCounter, "Check updated dates in InfoProviderNextRunDate", "DB values not found. Updated Next Report Date1: " & db_next_report_date3 & VBCRLF & "Updated Next Report Date2: " & db_next_report_date4 & VBCRLF & "Updated Next Run Date1: " & db_next_run_date3 & VBCRLF & "Updated Next Run Date2: " & db_next_run_date4, "FAILED"
			bFlag = False
		bRunFlag = False
	
	End If
	
	Set dbRecordSet = execute_db_query(query3, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	Append_TestHTML StepCounter, "Check updated SettlementAcceptance", "Select * from SettlementAcceptance order by 1 desc;", "PASSED"
	wait 2
	db_Settlement_id1 = dbRecordSet("SettlementAcceptanceID")
	db_Settlement_report_date1 = dbRecordSet("SettlementReportDate")
	Set dbRecordSet = Nothing
	If db_Settlement_id1 <> "" and db_Settlement_report_date1 <> "" Then
		Append_TestHTML StepCounter, "Check updated dates in SettlementAcceptance", "Updated Settlement ID: " & db_Settlement_id & VBCRLF & "Updated Settlement Report Date: " & db_Settlement_report_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Check updated dates in SettlementAcceptance", "DB values not found. Updated Settlement ID: " & db_Settlement_id & VBCRLF & "Updated Settlement Report Date: " & db_Settlement_report_date, "FAILED"
		bRunFlag = False
			bFlag = False
		
	End If	
Else
		Append_TestHTML StepCounter, "Check updated dates in SettlementAcceptance and preview", "Unable to process the settlement because report date/batch transaction got not updated", "FAILED"
			bFlag = False
		bRunFlag = False

End  If

	
	
End  Function


Public Function VerifySettlementJob19()

	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 19 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 19", "Select * from Job where JobTypeId = 19 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_id <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 19", "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 19", "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "FAILED"
		bRunFlag = False
			bFlag = False
	End If
	If db_status_id = 0 Then
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is not successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
		bRunFlag = False
			bFlag = False
	End If
	
	query1 = "Update Job Set NextRunDate = null where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Update Job 19", "Update Job Set NextRunDate = null where ID = '" & db_id & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	Set dbRecordSet = Nothing
	JobRundate = Date()
	wait 10
	query2 = "Select * from Job where JobTypeId = 19 and ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 19", "Select * from Job where JobTypeId = 19 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_status_id <> "" Then
		Append_TestHTML StepCounter, "Verify Job 19", "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Verify Job 19", "DB values not found. Status ID: " & db_status_id, "FAILED"
		bRunFlag = False
			bFlag = False
	End If
	If db_status_id = 4 Then
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID after update is successful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID after update is not successful. Expected Value: 4" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
		bRunFlag = False
			bFlag = False
	
	End If
	
End  Function


'******************************* HEADER ******************************************
' Description : The function to verify billing job 64 in GFN application
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function VerifySettlementJob155(stl_report_date, PAN_num)
''''msgbox "settjob64"
	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
	End If
	
	If stl_report_date = ""  or stl_report_date = empty Then
		stl_report_date = settllReportDate
	End If
	wait 10
	query = "Select * from Job where JobTypeId = 155 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 155", "Select * from Job where JobTypeId = 155 order by 1 desc;", "PASSED"
	wait 10
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_id = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	If db_id <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 155", "ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 155", "DB values not found. ID: " & db_id & VBCRLF & "Status ID: " & db_status_id, "FAILED"
		bRunFlag = False
			bFlag = False
	End If
'	If db_status_id = 0 Then
'		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "PASSED"
'	else
'		Append_TestHTML StepCounter, "Validated status ID", "Validation of status ID before update is not successful. Expected Value: 0" & VBCRLF & "Actual Value: " & db_status_id, "FAILED"
'	End If
	wait 15
	query1 = "Select SettlementReportdate,SalesItemID from SalesItemDelco where PAN = '" & PAN_num & "' order by 1 desc"
	Append_TestHTML StepCounter, "Verify SalesItemDelco entries after Job 155", "Select * from SalesItemDelco where PAN = '" & PAN_num & "'", "PASSED"
'	query1 = "Select * from SalesItemDelco where batchid = '" & Trans_BatchID & "'"
'	Append_TestHTML StepCounter, "Verify SalesItemDelco Report Dates", "Select * from SalesItemDelco where batchid = '" & Trans_BatchID & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 2,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_s_report_date = dbRecordSet("SettlementReportdate")
	db_sales_item_id = dbRecordSet("SalesItemID")
	date_arr_stl = split(db_s_report_date,"|")
	db_stl_report_date1 = date_arr_stl(0)
	db_stl_report_date2 = date_arr_stl(1)
'	''msgbox db_stl_report_date
'	''msgbox date_arr_stl
'	''msgbox db_stl_report_date1
'	''msgbox db_stl_report_date2
	
	sales_arr = split(db_sales_item_id,"|")
	db_sales_item_id1 = sales_arr(0)
	db_sales_item_id2 = sales_arr(1)
	Set dbRecordSet = Nothing
	If db_stl_report_date1 <> "" and db_stl_report_date2 <> ""  and db_sales_item_id1 <> "" and db_sales_item_id2 <> "" Then
		Append_TestHTML StepCounter, "Verify Sales Item table", "Settlement Report Date1: " & db_stl_report_date1 & VBCRLF & "Settlement Report Date2: " & db_stl_report_date2 & VBCRLF & "Sales Item ID1: " & db_sales_item_id1 & VBCRLF & "Sales Item ID2: " & db_sales_item_id2, "PASSED"
	else
		Append_TestHTML StepCounter, "Verify Sales Item table", "DB values not found. Settlement Report Date1: " & db_stl_report_date1 & VBCRLF & "Settlement Report Date2: " & db_stl_report_date2 & VBCRLF & "Sales Item ID1: " & db_sales_item_id1 & VBCRLF & "Sales Item ID2: " & db_sales_item_id2, "FAILED"
	bRunFlag = False
			bFlag = False
	End If
	If instr(db_stl_report_date1, stl_report_date) <> 0 Then
		Append_TestHTML StepCounter, "Validate Sales ItemDelco table", "Txns moved to Sales ItemDelco table with correct Settlement date. Expected Value: " & db_stl_report_date1 & VBCRLF & "Actual Value: " & stl_report_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Sales ItemDelco table", "Txns not found in Sales ItemDelco table. Expected Value: " & db_stl_report_date1 & VBCRLF & "Actual Value: " & stl_report_date, "FAILED"
		bRunFlag = False
	bFlag = False
		
	End If
	If instr(db_stl_report_date2, stl_report_date) <> 0 Then
		Append_TestHTML StepCounter, "Validate Sales ItemDelco table", "Txns moved to Sales ItemDelco table with correct Settlement date. Expected Value: " & db_stl_report_date2 & VBCRLF & "Actual Value: " & stl_report_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Sales ItemDelco table", "Txns not found in Sales ItemDelco table. Expected Value: " & db_stl_report_date2 & VBCRLF & "Actual Value: " & stl_report_date, "FAILED"
		bRunFlag = False
	bFlag = False
		
	End If
	
End  Function



Public Function backdateToCustomerDetails(cust_erp)

	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
	End If
If cust_erp <> "" Then
	'Get customer ID
	Set dbRecordSet = execute_db_query("Select * from Customer where CustomerERP = '" & cust_erp & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Customer ID", "Select * from Customer where CustomerERP = '" & cust_erp & "'", "PASSED"
	wait 2
	db_cust_id = dbRecordSet("CustomerID")
	Set dbRecordSet = Nothing
	If db_cust_id <> "" Then
		Append_TestHTML StepCounter, "Customer ID", "Customer ID: " & db_cust_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Customer ID", "DB values not found. Customer ID: " & db_cust_id, "FAILED"
	End If

	'Get actual date established and start date
	query = "Select * from Customer where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Get Dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_established_date = dbRecordSet("DateEstablished")
	db_start_date = dbRecordSet("StartDate")
	
	est_date_arr = split(db_established_date," ")	
	start_date_arr = split(db_start_date," ")
	'''''msgbox start_date_arr
	new_est_date = "2021-01-10 " & est_date_arr(1) 
	new_start_date = "2021-01-10 " & start_date_arr(1)
	'''''msgbox new_start_date
	Set dbRecordSet = Nothing
	If db_established_date <> "" and db_start_date <> "" Then
		Append_TestHTML StepCounter, "Get Dates", "DateEstablished: " & db_established_date & VBCRLF & "StartDate: " & db_start_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Get Dates", "DB values not found. DateEstablished: " & db_established_date & VBCRLF & "StartDate: " & db_start_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update Customer Set DateEstablished = '" & new_est_date & "', StartDate = '" & new_start_date & "' where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Update dates", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	query2 = "Select * from CustomerBankAccount where CustomerID=" & db_cust_id & ";"
	Append_TestHTML StepCounter, "Get Dates", query2, "PASSED"
	set dbRecordSet = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_BankAccountID = dbRecordSet("BankAccountID")
	
	Set dbRecordSet = Nothing
	Set query2 = Nothing
	
	query1 = "Update BankAccount Set DateEffective = '" & new_est_date & "' where BankAccountID = '" & db_BankAccountID & "'"
	Append_TestHTML StepCounter, "Update dates", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	query1 = "Update BankMandates Set DirectDebitDateEffective = '" & new_est_date & "' where BankAccountID = '" & db_BankAccountID & "'"
	Append_TestHTML StepCounter, "Update dates", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	Set query1 = Nothing
	
	'Get updated date established and start date
	query2 = "Select * from Customer where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Get Dates", query2, "PASSED"
	set dbRecordSet = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_updated_established_date = dbRecordSet("DateEstablished")
	db_updated_start_date = dbRecordSet("StartDate")
	Set dbRecordSet = Nothing
	If db_updated_established_date <> "" and db_updated_start_date <> "" Then
		Append_TestHTML StepCounter, "Get Dates", "Updated DateEstablished: " & db_updated_established_date & VBCRLF & "Updated StartDate: " & db_updated_start_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Get Dates", "DB values not found. Updated DateEstablished: " & db_updated_established_date & VBCRLF & "Updated StartDate: " & db_updated_start_date, "FAILED"
	End If
	
	timestamp_arr = split(db_updated_established_date," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	db_updated_established_date = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If db_updated_established_date =new_est_date  Then
		Append_TestHTML StepCounter, "Established date", "Updated established date successfully. Expected Value: " & new_est_date & VBCRLF & "Actual Value: " & db_updated_established_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Established date", "Update of established date failed. Expected Value: " & new_est_date & VBCRLF & "Actual Value: " & db_updated_established_date, "FAILED"
	End If
	
	timestamp_arr = split(db_updated_start_date," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	db_updated_start_date = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If db_updated_start_date =new_start_date  Then
		Append_TestHTML StepCounter, "Start date", "Updated start date successfully. Expected Value: " & new_start_date & VBCRLF & "Actual Value: " & db_updated_start_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Start date", "Update of start date failed. Expected Value: " & new_start_date & VBCRLF & "Actual Value: " & db_updated_start_date, "FAILED"
		bRunFlag = False
	End If
	
	'---------------------------------------------------------------------------------------
	'Get Info Subscription ID
	Set dbRecordSet = execute_db_query("Select * from CustomerInfoSubscription where CustomerID = '" & db_cust_id & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Info Subscription ID", "Select * from CustomerInfoSubscription where CustomerID = '" & db_cust_id & "'", "PASSED"
	wait 2
	db_infosub_id = dbRecordSet("InfoSubscriptionID")
	Set dbRecordSet = Nothing
	If db_infosub_id <> "" Then
		Append_TestHTML StepCounter, "Info Subscription ID", "Info Subscription ID: " & db_infosub_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Info Subscription ID", "DB values not found. Info Subscription ID: " & db_infosub_id, "FAILED"
	End If
	
	'Get actual date established and start date
	query = "Select * from InfoSubscription where InfoSubscriptionID = '" & db_infosub_id & "'"
	Append_TestHTML StepCounter, "Info Subscription", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update InfoSubscription Set DateEffective = '" & new_effective_date & "' Where InfoSubscriptionID = '" & db_infosub_id & "'"
	Append_TestHTML StepCounter, "Update DateEffective", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from InfoSubscription where InfoSubscriptionID = '" & db_infosub_id & "'"
	Append_TestHTML StepCounter, "Verify updated dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If

	'-----------------------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from CustomerPriceRule where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Get Customer PriceRule", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CustomerPriceRule Set DateEffective = '" & new_effective_date & "' Where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Update DateEffective", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from CustomerPriceRule where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Verify updated dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If
	



'Get actual date established and start date
	query = "Select * from CustomerFeeRule where CustomerID = '" & db_cust_id & "';"	' and FeeRuleID = 21"
	Append_TestHTML StepCounter, "Get CustomerFeeRule", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	db_customer_feerule = dbRecordSet("CustomerFeeRuleID")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CustomerFeeRule Set DateEffective = '" & new_effective_date & "' Where CustomerFeeRuleID = '" & db_customer_feerule & "'"
	Append_TestHTML StepCounter, "Update CustomerFeeRule", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from CustomerFeeRule where CustomerID = '" & db_cust_id & "';"			' and FeeRuleID = 21"
	Append_TestHTML StepCounter, "Verify updated dates", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If
	'-------------------------------------------------------------------------------
Else

		Append_TestHTML StepCounter, "Back date Customer details", "Customer ERP value is empty and not perform back date action", "FAILED"

End If	


	
	
End Function




Function backdateToCarddetails(cust_erp, PAN_num)

	On error resume next
	
	bFlag = True
	
	If PAN_num = ""  or PAN_num = empty Then
		PAN_num = cardPANNum
	End If
If cust_erp <> "" and PAN_num <> ""  Then
	'Get customer ID
	Set dbRecordSet = execute_db_query("Select * from Customer where CustomerERP = '" & cust_erp & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Customer ID", "Select * from Customer where CustomerERP = '" & cust_erp & "'", "PASSED"
	wait 2
	db_cust_id = dbRecordSet("CustomerID")
	Set dbRecordSet = Nothing
	If db_cust_id <> "" Then
		Append_TestHTML StepCounter, "Customer ID", "Customer ID: " & db_cust_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Customer ID", "DB values not found. Customer ID: " & db_cust_id, "FAILED"
	End If
	
	'Get PAN ID
	Set dbRecordSet = execute_db_query("Select * from Card where PAN = '" & PAN_num & "'", 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	Append_TestHTML StepCounter, "Get Card ID", "Select * from Card where PAN = '" & PAN_num & "'", "PASSED"
	wait 2
	db_card_id = dbRecordSet("CardID")
	Set dbRecordSet = Nothing
	If db_card_id <> "" Then
		Append_TestHTML StepCounter, "Card ID", "Card ID: " & db_card_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Card ID", "DB values not found. Card ID: " & db_card_id, "FAILED"
	End If


	'Get actual date established and start date
	query = "Select * from CustomerCard where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Customer Card", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_effective_date = dbRecordSet("DateEffective")
	new_effective_date = "2021-01-10"
	Set dbRecordSet = Nothing
	If db_effective_date <> "" Then
		Append_TestHTML StepCounter, "DateEffective", "DateEffective: " & db_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "DateEffective", "DB values not found. DateEffective: " & db_effective_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CustomerCard Set DateEffective = '" & new_effective_date & "' Where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Update Customer Card", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from CustomerCard where CustomerID = '" & db_cust_id & "'"
	Append_TestHTML StepCounter, "Validate Customer Card", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	updated_effective_date = dbRecordSet("DateEffective")
	Set dbRecordSet = Nothing
	If updated_effective_date <> "" Then
		Append_TestHTML StepCounter, "Updated DateEffective", "DateEffective: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateEffective", "DB values not found. DateEffective: " & updated_effective_date, "FAILED"
	End If
	
	If updated_effective_date =new_effective_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated effective date successfully. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating effective date failed. Expected Value: " & new_effective_date & VBCRLF & "Actual Value: " & updated_effective_date, "FAILED"
		bRunFlag = False
	End If
	'-----------------------------------------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from Card where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Get Card table", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_date = dbRecordSet("InitialCreationDateTime")
	date_arr = split(db_date," ")	
	new_date = "2021-01-10 " & date_arr(1) 
	Set dbRecordSet = Nothing
	If db_date <> "" Then
		Append_TestHTML StepCounter, "InitialCreationDateTime", "InitialCreationDateTime: " & db_date, "PASSED"
	else
		Append_TestHTML StepCounter, "InitialCreationDateTime", "DB values not found. InitialCreationDateTime: " & db_date, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update Card Set InitialCreationDateTime = '" & new_date & "' Where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Update Card table", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	'Get actual date established and start date
	query = "Select * from Card where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Validate card table", query, "PASSED"
	set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_updated_date = dbRecordSet("InitialCreationDateTime")
	Set dbRecordSet = Nothing
	If db_updated_date <> "" Then
		Append_TestHTML StepCounter, "Validate updated date", "InitialCreationDateTime: " & db_updated_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "DB values not found. InitialCreationDateTime: " & db_updated_date, "FAILED"
	End If
	
	timestamp_arr = split(db_updated_date," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	db_updated_date = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If db_updated_date =new_date  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated date successfully. Expected Value: " & new_date & VBCRLF & "Actual Value: " & db_updated_date, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating date failed. Expected Value: " & new_date & VBCRLF & "Actual Value: " & db_updated_date, "FAILED"
		bRunFlag = False
	End If
	'-----------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
	
	
	'Get actual date established and start date
	query = "Select * from CardStatusHistory where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Get CardStatusHistory", query, "PASSED"
'	and ModifiedBy = 'EFN'"
	set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_date = dbRecordSet("DateModified")
	db_date2 = dbRecordSet("DateModifiedNextStatus")
	date_rows = split(db_date,"|")
	date_arr1 = split(date_rows(0)," ")	
	date_arr2 = split(date_rows(1)," ")
	actual_date1 = 	date_rows(0)
	actual_date2 = 	date_rows(1)
	new_date1 = "2021-01-10 " & date_arr1(1) 
	new_date2 = "2021-01-10 " & date_arr2(1) 
	
	date_rows2 = split(db_date2,"|")
	date_arr3 = split(date_rows2(0)," ")
	new_date3 = "2021-01-10 " & date_arr3(1) 
	actual_date3 = 	date_rows2(0)
	
	cardStatusID = split(dbRecordSet("CardStatusHistoryID"),"|")
	cardStatusID1 = cardStatusID(0)
	cardStatusID2 = cardStatusID(1)
	Set dbRecordSet = Nothing
	If actual_date1 <> "" and  actual_date2 <> "" and  actual_date3 <> "" Then
		Append_TestHTML StepCounter, "DateModified & DateModifiedNextStatus", "DateModified1: " & actual_date1 & VBCRLF & "DateModified2: " & actual_date2 & VBCRLF & "DateModifiedNextStatus: " & actual_date3 , "PASSED"
	else
		Append_TestHTML StepCounter, "DateModified & DateModifiedNextStatus", "DB values not found. DateModified1: " & actual_date1 & VBCRLF & "DateModified2: " & actual_date2 & VBCRLF & "DateModifiedNextStatus: " & actual_date3, "FAILED"
	End If
	If cardStatusID1 <> "" and  cardStatusID2 <> "" Then
		Append_TestHTML StepCounter, "CardStatusHistoryID", "CardStatusHistoryID1: " & cardStatusID1 & VBCRLF & "CardStatusHistoryID2: " & cardStatusID2, "PASSED"
	else
		Append_TestHTML StepCounter, "CardStatusHistoryID", "DB values not found. CardStatusHistoryID1: " & cardStatusID1 & VBCRLF & "CardStatusHistoryID2: " & cardStatusID2, "FAILED"
	End If
	
	'Update with new date established and start date
	query1 = "Update CardStatusHistory Set DateModified = '" & new_date1 & "', DateModifiedNextStatus = '" & new_date3 & "' Where CardStatusHistoryID = '" & cardStatusID1 & "'"
	Append_TestHTML StepCounter, "Update CardStatusHistory", query1, "PASSED"
	set dbRecordSet = execute_db_query(query1, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	query2 = "Update CardStatusHistory Set DateModified = '" & new_date2 & "' Where CardStatusHistoryID = '" & cardStatusID2 & "'"
		Append_TestHTML StepCounter, "Update CardStatusHistory", query2, "PASSED"
	set dbRecordSet = execute_db_query(query2, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	Set dbRecordSet = Nothing
	
	
	'Get actual date established and start date
	query = "Select * from CardStatusHistory where CardID = '" & db_card_id & "'"
	Append_TestHTML StepCounter, "Validate CardStatusHistory", query, "PASSED"
'	and ModifiedBy = 'EFN'"
	set dbRecordSet = execute_db_query(query, 2, "SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_date = dbRecordSet("DateModified")
	db_date2 = dbRecordSet("DateModifiedNextStatus")
	date_rows = split(db_date,"|")
	
	date_rows2 = split(db_date2,"|")
	updated_date1 = date_rows(0)
	updated_date2 = date_rows(1)
	updated_date3 = date_rows2(0)
	Set dbRecordSet = Nothing
	If updated_date1 <> "" and  updated_date2 <> "" and  updated_date3 <> "" Then
		Append_TestHTML StepCounter, "Updated DateModified & DateModifiedNextStatus", "DateModified1: " & updated_date1 & VBCRLF & "DateModified2: " & updated_date2 & VBCRLF & "DateModifiedNextStatus: " & updated_date3 , "PASSED"
	else
		Append_TestHTML StepCounter, "Updated DateModified & DateModifiedNextStatus", "DB values not found. DateModified1: " & updated_date1 & VBCRLF & "DateModified2: " & updated_date2 & VBCRLF & "DateModifiedNextStatus: " & updated_date3, "FAILED"
	End If
	
	timestamp_arr = split(updated_date1," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	updated_date1 = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If updated_date1 =new_date1  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated DateModified1 successfully. Expected Value: " & new_date1 & VBCRLF & "Actual Value: " & updated_date1, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating DateModified1 failed. Expected Value: " & new_date1 & VBCRLF & "Actual Value: " & updated_date1, "FAILED"
		bRunFlag = False
	End If
	
	timestamp_arr = split(updated_date2," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	updated_date2 = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If updated_date2 =new_date2  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated DateModified2 successfully. Expected Value: " & new_date2 & VBCRLF & "Actual Value: " & updated_date2, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating DateModified2 failed. Expected Value: " & new_date2 & VBCRLF & "Actual Value: " & updated_date2, "FAILED"
		bRunFlag = False
	End If
	
	timestamp_arr = split(updated_date3," ")
	date_arr = split(timestamp_arr(0),"/")
	If int(date_arr(0))<10 Then
		date_arr(0) = "0"&date_arr(0)
	End If
	updated_date3 = date_arr(2) & "-" & date_arr(0) & "-" & date_arr(1) & " " & timestamp_arr(1)
	If updated_date3 =new_date3  Then
		Append_TestHTML StepCounter, "Validate updated date", "Updated DateModifiedNextStatus successfully. Expected Value: " & new_date3 & VBCRLF & "Actual Value: " & updated_date3, "PASSED"
	else
		Append_TestHTML StepCounter, "Validate updated date", "Updating DateModifiedNextStatus failed. Expected Value: " & new_date3 & VBCRLF & "Actual Value: " & updated_date3, "FAILED"
		bRunFlag = False
	End If
Else
		Append_TestHTML StepCounter, "Card Activation", "Card not created and unable to perform card backdate actions", "FAILED"
	
End If

''''msgbox "unknownbackdatefunction"
End Function


Public Function VerifyDX300C()
	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 291 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 291", "Select * from Job where JobTypeId = 203 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id = dbRecordSet("ID")
	Set dbRecordSet = Nothing
	
	If db_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 291", "ID: " & db_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 291", "DB values not found. ID: " & db_id, "FAILED"
	bFlag = False
	End If
	
	query1 = "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 291", "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	wait 10
	Set dbRecordSet = Nothing
	query2 = "select * from Job where JobTypeID = 292 order by 1 desc ;"
	Append_TestHTML StepCounter, "Verify Job 292",  query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id1 = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	
	If db_id1 <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 292", "ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "PASSED"
		If db_status_id = 4 Then
			Append_TestHTML StepCounter, "Validate Job status", "Expected:  " & db_status_id & VBCRLF & "Actual Status :4 ", "PASSED"
		
			''''''msgbox "Job 243 Successful"
			''''''msgbox db_id1
		End If
	else
		Append_TestHTML StepCounter, "Get ID from Job 292", "DB values not found. ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "FAILED"
	bFlag = False
		
	End If
	
End Function


'******************************* HEADER ******************************************
' Description : Function to run Job 203 and verify status of Job 243.
' Creator : Naman Joshi
' Date : 20th December,2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function VerifyDX300A()
	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 203 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 203", "Select * from Job where JobTypeId = 203 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id = dbRecordSet("ID")
	Set dbRecordSet = Nothing
	
	If db_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 203", "ID: " & db_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 203", "DB values not found. ID: " & db_id, "FAILED"
	bFlag = False
	End If
	
	query1 = "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 203", "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	wait 10
	Set dbRecordSet = Nothing
	query2 = "select * from Job where JobTypeID = 243 order by 1 desc ;"
	Append_TestHTML StepCounter, "Verify Job 243",  query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id1 = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	
	If db_id1 <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 243", "ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "PASSED"
		If db_status_id = 4 Then
			Append_TestHTML StepCounter, "Validate Job status", "Expected:  " & db_status_id & VBCRLF & "Actual Status :4 ", "PASSED"
		
			''''''msgbox "Job 243 Successful"
			''''''msgbox db_id1
		End If
	else
		Append_TestHTML StepCounter, "Get ID from Job 243", "DB values not found. ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "FAILED"
	bFlag = False
		
	End If
	
End Function

'******************************* HEADER ******************************************
' Description : Function to run Job 203 and verify status of Job 243.
' Creator : Naman Joshi
' Date : 21th December,2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function VerifyDX300()
	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 202 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 202", "Select * from Job where JobTypeId = 202 order by 1 desc;", "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id = dbRecordSet("ID")
	Set dbRecordSet = Nothing
	
	If db_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 202", "ID: " & db_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 202", "DB values not found. ID: " & db_id, "FAILED"
	End If
	
	query1 = "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 202", "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'", "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	wait 10
	Set dbRecordSet = Nothing
	query2 = "select * from Job where JobTypeID = 242 order by 1 desc;"
	Append_TestHTML StepCounter, "Verify Job 242",  "select * from Job where JobTypeID = 242 order by 1 desc", "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id1 = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	
	If db_id1 <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 242", "ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "PASSED"
		If db_status_id = 4 Then
			Append_TestHTML StepCounter, "Validate Job status", "Expected:  " & db_status_id & VBCRLF & "Actual Status :4 ", "PASSED"
			''''''msgbox "Job 242 Successful"
			''''''msgbox db_id1
		End If
	else
		Append_TestHTML StepCounter, "Get ID from Job 242", "DB values not found. ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "FAILED"
	End If
	
End Function
Public Function VerifyDX350data(custERP_value,dx350_Summary_Docno_val,dx350_Balance_val)
	On error resume next
	
	bFlag = True
	
	query1 = "Select * from PaymentItemImport where IncomingPartnerNumber= '" & custERP_value & "'"
	Append_TestHTML StepCounter, query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_patNo = dbRecordSet("IncomingPartnerNumber")
	db_DocNo= dbRecordSet("IncomingDocumentNumber")
	db_PayAmount = dbRecordSet("IncomingPaymentAmount")
	Set dbRecordSet = Nothing
		
	If db_patNo = custERP_value and db_DocNo =  dx350_Summary_Docno_val and db_PayAmount =  dx350_Balance_val Then
		Append_TestHTML StepCounter, "Validate DX350 PaymentItemImport table data", "Data saved with billing Document Number and details are - IncomingPartnerNumber  " & db_patNo & VBCRLF & "IncomingDocumentNumber: " & db_DocNo & VBCRLF & "IncomingPaymentAmount:"&db_PayAmount , "PASSED"
		
	else
		Append_TestHTML StepCounter, "Validate DX350 PaymentItemImport table data", "Miss match Data and details are - IncomingPartnerNumber  " & db_patNo & VBCRLF & "IncomingDocumentNumber: " & db_DocNo & VBCRLF & "IncomingPaymentAmount:"&db_PayAmount , "FAILED"
		bFlag=False
	End If
	
End Function


Public Function createandmoveDX602File(fPath,dPath,custERP)
	On error resume next
	bFlag = True

If customerERP_id <> "" Then

	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(fPath) Then
'		fileInitialName = "DX602_CREDIT_Data_086_000000"
		fileInitialName = DX602_FilePrefixNameval
		
		vd1= getmydate(now()-60)
		vd1_val = split(vd1," ")
		vd2= getmydate(now()-45)
		vd2_val = split(vd2," ")
		vd3= getmydate(now()-40)
		vd3_val = split(vd3," ")
		vd4= getmydate(now()-30)
		vd4_val = split(vd4," ")
		
		v1_date= vd1_val(0)
		v2_date = vd2_val(0)
		v3_date = vd3_val(0)
		v4_date=vd4_val(0)
		v4_time = vd4_val(1)

			
		nFname =  fileInitialName&"_"&v4_date&"_"&v4_time & ".dat"
		newfile = dPath & nFname
		
		fileSysObj.CreateTextFile(newfile)
		Set DXread = fileSysObj.OpenTextFile(fPath,1)
		content = DXread.ReadAll
		DXread.Close
		
		DXarray = Split(content,"|")
		If customerERP_id = "" Then
			customerERP_id = custerp
		End If
		
		modifyData = "2-"&custerp&";10-"& v1_date & ";11-"& v2_date &";17-"& v3_date
	
		Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
		sdata = Split(modifyData,";")
		Newcontent = content
		For itr = 0 To ubound(sdata) Step 1
			datavalue = sdata(itr)
			spos =Split(datavalue,"-")(0)
			svalue= Split(datavalue,"-")(1)
			Newcontent = Replace(Newcontent, DXarray(spos),svalue)
		Next
		Newcontent = Replace(Newcontent, DXarray(16),nFname)
		DXwrite.Write Newcontent
		DXwrite.Close
		JobRundate=date()
		
	End If
	If nFname <> "" Then
		Append_TestHTML StepCounter, "DX602 File creation and move to ID location", "File created at"& newfile , "PASSED"
		dx026FinalFName = nFname
		createandmoveDX602File = nFname
	Else
		createandmoveDX602File = ""
	bFlag = False
		
	End If

Else
		Append_TestHTML StepCounter, "DX602 File creation and move to ID location", "File will not process because no customer ERP " , "FAILED"
	bFlag = False
	
End If
		
End Function

'Public Function createandmoveDX602File(fPath,dPath,custerp)
'	On error resume next
'	bFlag = True
'If customerERP_id <> "" Then
'
'	Set fileSysObj = createObject("Scripting.FileSystemObject")
'	If fileSysObj.FileExists(fPath) Then
'		fileInitialName = "DX602_CREDIT_Data_085_000000"
'		
'		vd1= getmydate(now()-60)
'		vd1_val = split(vd1," ")
'		vd2= getmydate(now()-45)
'		vd2_val = split(vd2," ")
'		vd3= getmydate(now()-40)
'		vd3_val = split(vd3," ")
'		vd4= getmydate(now()-30)
'		vd4_val = split(vd4," ")
'		
'		v1_date= vd1_val(0)
'		v2_date = vd2_val(0)
'		v3_date = vd3_val(0)
'		v4_date=vd4_val(0)
'		v4_time = vd4_val(1)
'
'			
'		nFname =  fileInitialName&"_"&v4_date&"_"&v4_time & ".dat"
'		newfile = dPath & nFname
'		
'		fileSysObj.CreateTextFile(newfile)
'		Set DXread = fileSysObj.OpenTextFile(fPath,1)
'		content = DXread.ReadAll
'		DXread.Close
'		
'		DXarray = Split(content,"|")
'		If customerERP_id = "" Then
'			customerERP_id = custerp
'		End If
'		
'		modifyData = "2-"&custerp&";10-"& v1_date & ";11-"& v2_date &";17-"& v3_date
'	
'		Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
'		sdata = Split(modifyData,";")
'		Newcontent = content
'		For itr = 0 To ubound(sdata) Step 1
'			datavalue = sdata(itr)
'			spos =Split(datavalue,"-")(0)
'			svalue= Split(datavalue,"-")(1)
'			Newcontent = Replace(Newcontent, DXarray(spos),svalue)
'		Next
'		Newcontent = Replace(Newcontent, DXarray(16),nFname)
'		DXwrite.Write Newcontent
'		DXwrite.Close
'		JobRundate=date()
'		
'	End If
'	If nFname <> "" Then
'		Append_TestHTML StepCounter, "DX602 File creation and move to ID location", "File created at"& newfile , "PASSED"
'		dx026FinalFName = nFname
'		createandmoveDX602File = nFname
'	Else
'		createandmoveDX602File = ""
'	bFlag = False
'		
'	End If
'
'Else
'		Append_TestHTML StepCounter, "DX602 File creation and move to ID location", "File will not process because no customer ERP " , "FAILED"
'	bFlag = False
'	
'End If
'		
'End Function
'
Function getmydate(vdate)
	Dim dd, mm, yy, hh, nn, ss
	Dim datevalue, timevalue, dtsnow, dtsvalue
	'Store DateTimeStamp once.
	dtsnow = vdate
	'Individual date components
	dd = Right("00" & Day(dtsnow), 2)
	mm = Right("00" & Month(dtsnow), 2)
	yy = Year(dtsnow)
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00" & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	'Build the date string in the format yyyy-mm-dd
	datevalue = yy &  mm & dd
	'Build the time string in the format hh:mm:ss
	timevalue = hh & nn  & ss
	'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss
	dtsvalue = datevalue & " " & timevalue
	getmydate = dtsvalue
End Function



Function valuetrailzeros(seqno)
	lenno= len(seqno)
	For i = 1 To 6-leno Step 1
		seqno = "0" & seqno 
	Next
	valuetrailzeros = seqno
End Function
	

Public Function createandmoveDX350File(fPath,dPath,custerp,d_Summary_Docno,d_Balance,d_Paymentduedate )
	On error resume next
	bFlag = True
	
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	
If customerERP_id<>"" Then
	
	If fileSysObj.FileExists(fPath) Then
		query = "Select * from FileSeq where FileSeqID='DX350';"
	'	set dictDbResultSet = execute_db_query(query)
		set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		
		dx350_Filesequenceno= dictDbResultSet("LastSequence")
		seqnum_val = valuetrailzeros(dx350_Filesequenceno+1)
	
		Set dictDbResultSet = Nothing
	
		fileInitialName = DX350_FilePrefixNameval
		newpay_date =Split(dx350_Paymentduedate,"-")
		ndateval = newpay_date(2)&"-"&newpay_date(1)&"-"&newpay_date(0)
		payrolldate = cdate(ndateval)-7
		transrolldate = cdate(payrolldate) + 2
			transpaysubdate = reqdateforDBfileFormt(transrolldate)
		Paymentddate =   reqdateforfileFormt(payrolldate)
		paysubmitdate = reqdateforfileFormt(transrolldate)
		timeval = getDateandTimestamp()
		timeval_arr = split(timeval," ")
		timevaltext = timeval_arr(1)
			
		nFname =  fileInitialName&"_"&seqnum_val&"_"&transpaysubdate&"_"&timevaltext & ".dat"
		newfile = dPath & nFname
		
		fileSysObj.CreateTextFile(newfile)
		Set DXread = fileSysObj.OpenTextFile(fPath,1)
		content = DXread.ReadAll
		DXread.Close
		
		DXarray = Split(content,"|")
		If customerERP_id = "" Then
			customerERP_id = custerp
		End If
		
		modifyData = "1-"&custerp&";2-"& d_Summary_Docno & ";4-"& d_Balance &";6-"& Paymentddate &";18-"&d_Balance
	
		Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
		sdata = Split(modifyData,";")
		Newcontent = content
		For itr = 0 To ubound(sdata) Step 1
			datavalue = sdata(itr)
			spos =Split(datavalue,"-")(0)
			svalue= Split(datavalue,"-")(1)
			If svalue <> "" Then
				Newcontent = Replace(Newcontent, DXarray(spos),svalue)
			Else
				'Append_TestHTML StepCounter, "DX350 File creation and move to ID location", "Unable to generated file" , "FAILED"
				'bFlag = False
				
				createandmoveDX350File = ""
			End If
		Next
		Newcontent = Replace(Newcontent, DXarray(16),nFname)
		DXwrite.Write Newcontent
		DXwrite.Close
		JobRundate=date()
		
	End If
	If nFname <> "" Then
		Append_TestHTML StepCounter, "DX350 File creation and move to ID location", "File created at"& newfile , "PASSED"
		dx026FinalFName = nFname
		createandmoveDX350File = nFname
	Else
		createandmoveDX350File = ""
	bFlag = False
		
	End If
Else
		Append_TestHTML StepCounter, "DX350 File creation and move to ID location", "Unable to crete new file creation because no customer ERP value" , "FAILED"
	bFlag = False
	
End If		
End Function	


Public Function createandmoveDX451File(fPath,dPath)
	On error resume next
	bFlag = True
	
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	
If customerERP_id<>"" Then
	
	If fileSysObj.FileExists(fPath) Then
		query = "Select * from FileSeq where FileSeqID='DX451_SITE';"
	'	set dictDbResultSet = execute_db_query(query)
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_MASTER")
		wait 2
		
		dx451_Filesequenceno= dictDbResultSet("LastSequence")
		seqnum_val = valuetrailzeros(dx451_Filesequenceno+1)
	
		Set dictDbResultSet = Nothing
	
		fileInitialName = DX451_FilePrefixNameval
'		newpay_date =Split(dx350_Paymentduedate,"-")
'		ndateval = newpay_date(2)&"-"&newpay_date(1)&"-"&newpay_date(0)
'		payrolldate = cdate(ndateval)-7
'		transrolldate = cdate(payrolldate) + 2
'			transpaysubdate = reqdateforDBfileFormt(transrolldate)
'		Paymentddate =   reqdateforfileFormt(payrolldate)
'		paysubmitdate = reqdateforfileFormt(transrolldate)
		timeval = getDateandTimestamp()
		timeval_arr = split(timeval," ")
		timedateltext = replace(timeval_arr(0),"-","")
		timevaltext = timeval_arr(1)
			
		nFname =  fileInitialName&"_"&seqnum_val&"_"&timedateltext&"_"&timevaltext & ".dat"
		newfile = dPath & nFname
		
		fileSysObj.CreateTextFile(newfile)
	
		Set DXread = fileSysObj.OpenTextFile(fPath,1,False,-1)
		content = DXread.ReadAll
		DXread.Close
		
		DXarray = Split(content,"|")
		If customerERP_id = "" Then
			customerERP_id = custerp
		End If
		
		filePath = sCurrentDirectory & "Test Data\site_"& countryCode & ".txt"
	Set fileSysObj12 = createObject("Scripting.FileSystemObject")
	Set fileRead12 = fileSysObj.OpenTextFile(filePath,1)
	sitecontent = fileRead12.ReadAll
	fileRead12.Close
	Set DXwrite12 = fileSysObj12.OpenTextFile(filePath,2)
	Newsitecontent = sitecontent + 1
	registrationnumber = Newsitecontent
	siteERPname = countryCode & registrationnumber
	site451Name = siteERPname
	DXwrite12.Write Newsitecontent
	DXwrite12.Close
	Set fileSysObj12 = Nothing
		siteunqNo = site451Name
		siteFName = countryCode & registrationnumber & " AutomationFName"
		siteSName = countryCode & registrationnumber & " AutomationSName"
		soneunqNo = right(registrationnumber,6)
		modifyData = "3-"&siteunqNo&";5-"& siteFName & ";6-"& siteSName & ";23-"& soneunqNo 
	
		Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
		sdata = Split(modifyData,";")
		Newcontent = content
		For itr = 0 To ubound(sdata) Step 1
			datavalue = sdata(itr)
			spos =Split(datavalue,"-")(0)
			svalue= Split(datavalue,"-")(1)
			
			If svalue <> "" Then
				Newcontent = Replace(Newcontent, DXarray(spos),svalue)
			Else
				'Append_TestHTML StepCounter, "DX350 File creation and move to ID location", "Unable to generated file" , "FAILED"
				'bFlag = False
				
				createandmoveDX451File = ""
			End If
		Next
		Newcontent = Replace(Newcontent, DXarray(26),nFname)
		DXwrite.Write Newcontent
		DXwrite.Close
		JobRundate=date()
		
	End If
	If nFname <> "" Then
		Append_TestHTML StepCounter, "DX451 File creation and move to ID location", "File created at"& newfile , "PASSED"
		'dx026FinalFName = nFname
		createandmoveDX451File = nFname
	Else
		createandmoveDX451File = ""
	bFlag = False
		
	End If
Else
		Append_TestHTML StepCounter, "DX350 File creation and move to ID location", "Unable to crete new file creation because no customer ERP value" , "FAILED"
	bFlag = False
	
End If		
End Function	



Function valuetrailzeros(seqno)
	lenno= len(seqno)
	totaldigits = 6-lenno

	For i = 1 To totaldigits Step 1
		seqno = "0" & seqno 
	Next
	valuetrailzeros = seqno
End Function


Function reqdateforfileFormt(original_date)
	On error resume next
	bFlag = True
			n_db_original_date = cdate(original_date)
			n_day =  datepart("d", n_db_original_date )
			n_month =  datepart("m", n_db_original_date )
			n_year = datepart("yyyy", n_db_original_date )
			If len(n_day) = 1  Then
				n_day = "0" & n_day 
			End If
			If len(n_month) = 1  Then
				n_month = "0" & n_month 
			End If
			next_date =n_day&n_month&n_year
			reqdateforfileFormt =  next_date
End Function

Function reqdateforDBfileFormt(original_date)
	On error resume next
	bFlag = True
			n_db_original_date = cdate(original_date)
			n_day =  datepart("d", n_db_original_date )
			n_month =  datepart("m", n_db_original_date )
			n_year = datepart("yyyy", n_db_original_date )
			If len(n_day) = 1  Then
				n_day = "0" & n_day 
			End If
			If len(n_month) = 1  Then
				n_month = "0" & n_month 
			End If
			next_date = n_year&n_month&n_day
			reqdateforDBfileFormt =  next_date
End Function
Function reqdateFormt()
On error resume next
	bFlag = True
			original_date = Date()
			n_db_original_date = cdate(original_date)
			n_day =  datepart("d", n_db_original_date )
			n_month =  datepart("m", n_db_original_date )
			n_year = datepart("yyyy", n_db_original_date )
			If len(n_day) = 1  Then
				n_day = "0" & n_day 
			End If
			If len(n_month) = 1  Then
				n_month = "0" & n_month 
			End If
			next_date =n_day&"/"&n_month&"/"&n_year
			reqdateFormt =  next_date
		End Function

Function nextdateDBFormt(original_date)
On error resume next
	bFlag = True
			n_db_original_date = cdate(original_date)+1
			n_day =  datepart("d", n_db_original_date )
			n_month =  datepart("m", n_db_original_date )
			n_year = datepart("yyyy", n_db_original_date )
			If len(n_day) = 1  Then
				n_day = "0" & n_day 
			End If
			If len(n_month) = 1  Then
				n_month = "0" & n_month 
			End If
			next_date = n_year&"-"&n_month&"-"&n_day
			nextdateDBFormt =  next_date
End Function

Function getDateandTimestamp()
On error resume next
	bFlag = True
	Dim dd, mm, yy, hh, nn, ss
	Dim datevalue, timevalue, dtsnow, dtsvalue
	'Store DateTimeStamp once.
	dtsnow = Now()
	'Individual date components
	dd = Right("00" & Day(dtsnow), 2)
	mm = Right("00" & Month(dtsnow), 2)
	yy = Year(dtsnow)
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00" & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	'Build the date string in the format yyyy-mm-dd
	datevalue = yy & "-"& mm & "-" & dd
	'Build the time string in the format hh:mm:ss
	timevalue = hh & nn  & ss
	'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss
	dtsvalue = datevalue & " " & timevalue
	getDateandTimestamp = dtsvalue
End Function
Function getoutputtabledata(cno,value)
On error resume next
	bFlag = True

	Set outputtblobj = Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("class:=FlexGrid fg_pb hasFloatHead","html id:=ctl00_CPH.*_Table","html tag:=TABLE")
	If outputtblobj.Exist Then
		out_data = outputtblobj.GetCellData(2,cno)
		If instr(out_data,value)>0 Then
				Append_TestHTML StepCounter, "Validate "& value, "Entry added into the outpu table", "PASSED"
		Else
				Append_TestHTML StepCounter, "Validate "& value, "Fail to add into the outpu table", "FAILED"
				bFlag=False
		End If
	End If
	getoutputtabledata = bFlag
	Set outputtblobj = Nothing
End Function
Function menuSelection(mainlink,sublink)
	On error resume next
	bFlag = True
	Set linkobj = Browser("creationTime:=1").Page("creationTime:=1").Link("html tag:=A","innertext:="&mainlink)
	Set linkobj1 = Browser("creationTime:=1").Page("creationTime:=1").Link("html tag:=A","innertext:="&sublink)
	linkobj.Highlight
	wait 2
	Setting.webPackage("ReplayType")=2
	wait 2
	linkobj.FireEvent "onmouseover"
	wait 2
	linkobj.Click
	wait 2
	linkobj1.Highlight
	wait 2
	Append_TestHTML StepCounter,"Navigate "& sublink & "screen","Options are displayed correctly under customer-"& customerERP_id ,"PASSED"
	wait 5
	linkobj1.Click
	wait 5
'	'''''msgbox sublink
	'linkobj2.Click
	'wait 2
	Setting.webPackage("ReplayType")=1
	If Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("creationTime:=1").Page("creationTime:=1").WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If
	wait 2
Set linkobj = Nothing
Set linkobj1 = Nothing
End Function

Public Function navigateStartMenu(mainoptionobj,suboptionobj,verifyoptionobj)
	On error resume next
	bFlag = True
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Start", "Mover", ""
	If suboptionobj = "" Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", mainoptionobj, "Click", ""
	Else
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "",mainoptionobj, "Mover", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", suboptionobj, "Click", ""
	End If
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", verifyoptionobj)  Then
		Print "Navigated to"&verifyoptionobj
	Else
		bFlag = False
	End IF		
	
'	'''''msgbox "end start"
	navigateStartMenu = bFlag
End Function
Public Function customerSearch()
	On error resume next
	bFlag = True
	Call  navigateStartMenu("Link_Customers","Link_SearchforCustomer","WebLIST_Role")
'	'''''msgbox "customerSearch"
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLIST_Role")  Then
		Append_TestHTML StepCounter,"Search for Customer","Insert newly created Custer ERP value: "& customerERP_id ,"PASSED"
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"
	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", customerERP_id
			
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	Else
		Append_TestHTML StepCounter,"Search for Customer","Fail to get customer search page to insert "& customerERP_id ,"FAILED"
		bFlag = False		
	End  If
'	'''''msgbox "end customersearch"
	customerSearch = bFlag
End Function
Public Function navigateCustomerSummaryMenuoption(selectionobj,verifyobj)
	
'	Customer Summary
On error resume next
	bFlag = True
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary")  Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CustomerSummary", "Mover", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_CustomerSummary", "Click", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", selectionobj, "Click", ""
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", verifyobj)  Then
				Print "Navigated to"&verifyobj
			Else
				bFlag = False
			End IF	
	End If
	navigateCustomerSummaryMenuoption=bFlag

End Function

Public Function createSession()


On error resume next
	bFlag = True
	prev_month_date_dbval =  previousmonthDBFormt(date())
	'prev_run_date_dbval =  previousmonthDBFormt(prev_month_date_dbval)
	crr_date_dbval =  nextdateDBFormt(date()-1)
	'Select * from BillingAcceptance where BillingReportDate >= prev_month_date_dbval and BillingReportDate <= crr_date_dbval 
	'"2021-12-12"
'	prev_month_date_dbval="2022-01-12"
'	crr_date_dbval = "2022-01-12"
	Billquery = "Select min(BillingAcceptanceID) as MinBillaccid from BillingAcceptance where BillingReportDate >='" & prev_month_date_dbval & "'and BillingReportDate <='" & crr_date_dbval  & "';"
				Append_TestHTML StepCounter, "Get minimum reportdate id", Billquery, "PASSED"
	
	'	set dictDbResultSet = execute_db_query(query)
		set dictDbResultSet = execute_db_query(Billquery, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
		wait 2
		
		minbillingaccID= dictDbResultSet("MinBillaccid")
		''''msgbox isNull(minbillingaccID)
		set dictDbResultSet =Nothing
		
		If  isNull(minbillingaccID) Then
				Append_TestHTML StepCounter, "Validate reportdate id and perform delete entries", "No past one month entries are exist in BillingAcceptance Table and delete action", "PASSED"
			Else
				query1 = "Delete from BillingAcceptance Where BillingAcceptanceID > "& minbillingaccID &";"
				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query1, "PASSED"
				set dictDbResultSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
				query2 = "Delete from BillingAcceptanceInfoProvider Where BillingAcceptanceID > "& minbillingaccID &";"
				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query2, "PASSED"
				set dictDbResultSet = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
				query3 = "Delete from BillingAcceptanceCustomer Where BillingAcceptanceID > "& minbillingaccID &";"
				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query3, "PASSED"
				set dictDbResultSet = execute_db_query(query3, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
				query4 = "Delete from BillingAcceptanceTotal Where BillingAcceptanceID >"& minbillingaccID &";"
				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query4, "PASSED"
				set dictDbResultSet = execute_db_query(query4, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
				query5 = "Delete from BillingAcceptanceProductTotal Where BillingAcceptanceID > "& minbillingaccID &";"
				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query5, "PASSED"
				set dictDbResultSet = execute_db_query(query5, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
				query6 = "Delete from BillingAcceptanceFeeTotal Where BillingAcceptanceID > "& minbillingaccID &";"
				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query6, "PASSED"
				set dictDbResultSet = execute_db_query(query6, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
		End If		
				wait 4
				query = "Select * from BillingAcceptance order by 1 desc;"
				Append_TestHTML StepCounter, "Get last billing report date from BillingAcceptance Table",query, "PASSED"
				set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				currentBillinRepDate= dictDbResultSet("BillingReportDate")
				''''msgbox currentBillinRepDate
				set dictDbResultSet = Nothing
				currentNextBillRepDate = nextdateDBFormt(currentBillinRepDate)
				PrevNextBillRepDate = previousmonthDBFormt(currentNextBillRepDate)
				DX26SessionBillDate = PrevNextBillRepDate
				Append_TestHTML StepCounter, "Current Billing Report Date","Billing Report Date is:"& currentBillinRepDate  , "PASSED"
				
				query = "Update InfoProviderNextRunDate set NextReportDate = '"& currentNextBillRepDate &"', PreviousReportDate = '"& PrevNextBillRepDate &"' where InfoProviderID=4;"
				Append_TestHTML StepCounter, "Updating InfoProviderNextRunDate table entries for Billing",query, "PASSED"
				set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
			
				feetypequery = "Update FeeNextBillingReportDate set NextBillingReportDate = '" & currentNextBillRepDate &"' where FeeTypeID =1002;"		'FeeTypeID =38 or 
				Append_TestHTML StepCounter, "Update FeeNextBillingReportdate details",feetypequery, "PASSED"
				set dictDbResultSet = execute_db_query(feetypequery, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
				
				feerulequery = "Update FeeNextCreationDate set NextFeeCreationDate = '" & currentNextBillRepDate &"',PreviousFeeCreationDate='"& PrevNextBillRepDate &"' where Feeruleid=1002;"	'100025 '5 or Feeruleid=14;"
				Append_TestHTML StepCounter, "Update FeeNextBillingReportdate details",feerulequery, "PASSED"
				set dictDbResultSet = execute_db_query(feerulequery, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				set dictDbResultSet = Nothing
		
'		prev_month_date_dbval =  previousmonthDBFormt(date())
'	crr_date_dbval =  nextdateDBFormt(date()-1)
'	settelquery = "Select min(SettlementAcceptanceID) as MinSetlaccid from SettlementAcceptance where SettlementReportDate >='" & prev_month_date_dbval & "'and SettlementReportDate <='" & crr_date_dbval  & "';"
'				Append_TestHTML StepCounter, "Get minimum reportdate id", settelquery, "PASSED"
'	
'		set dictDbResultSet = execute_db_query(settelquery, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'		wait 2
'		
'		minsetlaccID= dictDbResultSet("MinSetlaccid")
'				set dictDbResultSet = Nothing
'		
'		If  isNull(minsetlaccID) Then
'				Append_TestHTML StepCounter, "Validate reportdate id and perform delete entries", "No past one month entries are exist in SettlementAcceptance Table and delete action", "PASSED"
'			Else
'				query1 = "Delete from SettlementAcceptance Where SettlementAcceptanceID > "& minsetlaccID &";"
'				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query1, "PASSED"
'				set dictDbResultSet = execute_db_query(query1, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
'				query2 = "Delete from SettlementAcceptanceInfoProvider Where SettlementAcceptanceID > "& minsetlaccID &";"
'				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query2, "PASSED"
'				set dictDbResultSet = execute_db_query(query2, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
'				query3 = "Delete from SettlementAcceptanceSiteOwner Where SettlementAcceptanceID > "& minsetlaccID &";"
'				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query3, "PASSED"
'				set dictDbResultSet = execute_db_query(query3, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
'				query4 = "Delete from SettlementAcceptanceTotal Where SettlementAcceptanceID >"& minsetlaccID &";"
'				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query4, "PASSED"
'				set dictDbResultSet = execute_db_query(query4, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
'				query5 = "Delete from SettlementAcceptanceProductTotal Where SettlementAcceptanceID > "& minsetlaccID &";"
'				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query5, "PASSED"
'				set dictDbResultSet = execute_db_query(query5, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
'				query6 = "Delete from SettlementAcceptanceFeeTotal Where SettlementAcceptanceID > "& minsetlaccID &";"
'				Append_TestHTML StepCounter, "Delete entries from BillingAcceptance Table",query6, "PASSED"
'				set dictDbResultSet = execute_db_query(query6, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
'		End If		
'				wait 4
'				query = "Select * from SettlementAcceptance order by 1 desc;"
'				Append_TestHTML StepCounter, "Get last billing report date from SettlementAcceptance Table",query, "PASSED"
'				set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				currentSettleRepDate= dictDbResultSet("SettlementReportDate")
'				'''msgbox currentSettleRepDate
'				set dictDbResultSet = Nothing
'				currentNextSettleRepDate = nextdateDBFormt(currentSettleRepDate)
'				PrevNextSettleRepDate = previousmonthDBFormt(currentNextSettleRepDate)
'				DX26SessionSettleDate = PrevNextBillRepDate
'				query = "Update InfoProviderNextRunDate set NextReportDate = '"& currentNextSettleRepDate &"', PreviousReportDate = '"& PrevNextSettleRepDate &"' where InfoProviderID=8;"
'				Append_TestHTML StepCounter, "Updating InfoProviderNextRunDate table entries for Billing",query, "PASSED"
'				set dictDbResultSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_BATCH")
'				wait 2
'				set dictDbResultSet = Nothing
		
End Function





Function previousmonthDBFormt(original_date)
On error resume next
	bFlag = True
	
			n_db_original_date = cdate(original_date)-30
			n_day =  datepart("d", n_db_original_date )
			n_month =  datepart("m", n_db_original_date )
			n_year = datepart("yyyy", n_db_original_date )
			If len(n_day) = 1  Then
				n_day = "0" & n_day 
			End If
			If len(n_month) = 1  Then
				n_month = "0" & n_month 
			End If
			next_date = n_year&"-"&n_month&"-"&n_day
			previousmonthDBFormt =  next_date
End Function
		
'/*******************		New Functions ******************************/
Function addAddressEmailofPinadvanceType()
	
On Error Resume Next
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "Card_WebList_PinadvType") Then
				Append_TestHTML StepCounter, "Create card", "Added details on create new card page", "PASSED"
				pinAdviceTypeVal = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Card_WebList_PinadvType","GetROProperty","selection")
				
				If ucase(pinAdviceTypeVal) = ucase("Email") Then
					OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_overridetab", "Click", ""
					wait 4
					UCDACheckval = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_uCDAobj","GetROProperty","checked")
						If UCDACheckval = "0" Then
							OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_APDEmail", "Set", "testnewemail@shell.com"
							wait 2
							emailaddress =   OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_APDEmail","GetROProperty","value")
						Else	
							OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_ACDEmail", "Set", "testnewemail@shell.com"
							wait 2
							emailaddress =   OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_ACDEmail","GetROProperty","value")
					
						End If
				
							If ucase(emailaddress) =ucase("testnewemail@shell.com") Then
								Append_TestHTML StepCounter, "Verify Card Deliver Address screen Fields", "Address Email Field is filled with value", "PASSED"
		'						OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		'							If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") = True  Then
		'								Append_TestHTML StepCounter, "Verify Email address is mandatary ", " Email address field is mandatary.","PASSED"
		'								Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebElement("class:=button ctl00_ctl12","html tag:=SPAN","innertext:=Save").Click	
		'								
		'							Else
		'							
		'								Append_TestHTML StepCounter, "Verify Email address is mandatary ", " Card Address Email field is empty","FAILED"
		'							bRunFlag = False
		'								
		'							End If
							Else
									Append_TestHTML StepCounter, "Verify PIN Advice Type Value", " Value is not set into Card Address email field","FAILED"
									bRunFlag = False
										
							End If 
				
				Else
						Append_TestHTML StepCounter, "Verify PIN Advice Type Default Value", " Default Value is not set into Email field","PASSED"
'									bRunFlag = False
				End If
		else
			Append_TestHTML StepCounter, "Verify Pin Advice Type filed in New Card Screen", "Field not exist in the screen", "FAILED"
		End If
End Function



Function fillCardAddressEmailofPinadvanceType()
	
On Error Resume Next

					If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebEdit_CardDelieveryAddress") = True  Then
						OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_ACDEmail", "Set", "testnewemail@shell.com"
						wait 2
							card_emailaddress =   OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Epin_WE_ACDEmail","GetROProperty","value")
							Append_TestHTML StepCounter, "Navigate to Override Screen", " Successfully navigated","PASSED"
							card_address =   OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_overrideAddress","GetROProperty","value")
							If ucase(card_emailaddress) =ucase("testnewemail@shell.com")  Then
								Append_TestHTML StepCounter, "Verify Card Deliver Address screen Fields", "Card Address Email Field is filled with value", "PASSED"
		'						OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		'							If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") = True  Then
		'								Append_TestHTML StepCounter, "Verify Email address is mandatary ", " Email address field is mandatary.","PASSED"
		'								Browser("name:=.*Shell Indonesia").Page("title:=.*Shell Indonesia").WebElement("class:=button ctl00_ctl12","html tag:=SPAN","innertext:=Save").Click	
		'								
		'							Else
		'							
		'								Append_TestHTML StepCounter, "Verify Email address is mandatary ", " Card Address Email field is empty","FAILED"
		'							bRunFlag = False
		'								
		'							End If
							Else
									Append_TestHTML StepCounter, "Verify PIN Advice Type Value", " Value is not set into Card Address email field","FAILED"
									bRunFlag = False
										
							End If 
						Else
								Append_TestHTML StepCounter, "Verify Carddelivery Address field in Override tab", " Fail to navigate override screen tab","FAILED"
									bRunFlag = False
								
						End  IF
			
End Function



'/*************************  Customer G


	Public Function VerifyCustomerGuaranteeCreation(cust_data)

	On error resume next
	
	bFlag = True
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Search for Customers"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchForCust", "Click", ""
	wait 1
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"

	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", cust_data("custERP")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""


	If Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Exist Then
		Browser("name:="&browserProp).Page("title:="&pageProp).WebElement("class:=button ct.*","html tag:=SPAN","innertext:=No").Click
		wait 2
	End If

		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", "Customer Guarantee"
	wait 2
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_CustomerGuarantee", "Click", ""
		
		wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewCustomerGuarantee") Then
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewCustomerGuarantee", "Click", ""
			wait 2
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_GuaranteeType", "RadioSelect", cust_data("oGuarantType")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_GuaranteeCurrency", "RadioSelect", cust_data("oGCurrency")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_GExpiry_SDate", "Set", cust_data("oGEDate")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_GuaranteeDetails", "Set", cust_data("oGDetails")
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_GuaranteeAdjustmentValue", "Set", cust_data("oGAdjustmentval")
		Append_TestHTML StepCounter,"Create new manual Fee details","All details are maintained" ,"PASSED"
		
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		
		wait 4   ' waiting for the UI update
		Append_TestHTML StepCounter,"Customer Guarantee details ","Details are saved" ,"PASSED"
		wait 15
	
	End If	
	wait 12   ' waiting for the UI update
	If customerERP_id <> "" Then
	
		query = "select * from Customer where CustomerERP='" &cust_data("custERP")  & "';"
	
		Append_TestHTML StepCounter,"Get Customer ID ",query,"PASSED"
	
		set dbRecordSet = execute_db_query(query, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_Cust_ID = dbRecordSet("CustomerID")
		set dbRecordSet = Nothing
	
		query_fee_item = "Select * from CustomerGuarantees where CustomerID ='" & db_Cust_ID &  "' order by 1 desc;"
	'	set dbRecordSet_Fee_Item = execute_db_query(query_fee_item)
		Append_TestHTML StepCounter,"Customer Guarantees details ",query_fee_item,"PASSED"
	
		set dbRecordSet_Fee_Item = execute_db_query(query_fee_item, 1, "SFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
		db_GTID = dbRecordSet_Fee_Item("GuaranteeTypeID")
		db_GValue = dbRecordSet_Fee_Item("GuaranteeValue")
		db_GDetails = dbRecordSet_Fee_Item("GuaranteeDetails")
		db_GExpirydate= dbRecordSet_Fee_Item("GuaranteeExpiryDate")
		set dbRecordSet_Fee_Item = Nothing
		
		If  db_GTID <> ""  Then  ' Fee type 38 = Manual Fee
			Append_TestHTML StepCounter,"Customer Guarantee Type","Expected Value: " & "1" & VBCRLF & "Actual Value: " & db_GTID ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Guarantee Type","Expected Value: " & "1" & VBCRLF & "Actual Value: " & db_GTID ,"FAILED"
			bFlag = False
		End If
		
		If  db_GValue <> ""  Then
			Append_TestHTML StepCounter,"Customer Guarantee Adjustment Value","Expected Value: " & cust_data("oGAdjustmentval") & VBCRLF & "Actual Value: " & db_GValue ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Guarantee Adjustment Value","Expected Value: " & cust_data("oGAdjustmentval") & VBCRLF & "Actual Value: " &  db_GValue ,"FAILED"
			bFlag = False
		End If
		
		If db_GDetails <> ""  Then
			Append_TestHTML StepCounter,"Customer Guarantee Details","Expected Value: " & cust_data("oGDetails") & VBCRLF & "Actual Value: " & db_GDetails ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Guarantee Details","Expected Value: " & cust_data("oGDetails") & VBCRLF & "Actual Value: " & db_GDetails ,"FAILED"
			bFlag = False
		End If
		
		If db_GExpirydate  <> "" Then
			Append_TestHTML StepCounter,"Customer Expirydate","Expected Value: " & cust_data("oGEDate") & VBCRLF & "Actual Value: " & db_GExpirydate ,"PASSED"
		else
			Append_TestHTML StepCounter,"Customer Expirydate","Expected Value: " & cust_data("oGEDate") & VBCRLF & "Actual Value: " & db_manual_fee_text ,"FAILED"
			bFlag = False
		End If
	
	End If
	
	If bFlag = True Then
		Append_TestHTML StepCounter,"Validation successful","Validation successful for verifying Customer Guarantee Creation details","PASSED"
	else
		Append_TestHTML StepCounter,"Validation failed","Validation failed for for verifying  Customer Guarantee Creation details details","FAILED"
		bRunFlag = False
	End If

End Function
	


'******************************* HEADER ******************************************
' Description : Function to run Job 212 and verify status of Job 262.
' Creator : Venkata Srinivasa Rao
' Date : 20th December,2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function VerifyDX300B()
	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 212 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 212", query , "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id = dbRecordSet("ID")
	Set dbRecordSet = Nothing
	
	If db_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 203", "ID: " & db_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 203", "DB values not found. ID: " & db_id, "FAILED"
	bFlag = False
	End If
	
	query1 = "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 212", query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	wait 10
	Set dbRecordSet = Nothing
	Set query1= Nothing
	query1 = "select * from Job where JobTypeID = 212 and ID = '" & db_id & "';"
	Append_TestHTML StepCounter, "Verify Job 212",  query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_212_id1 = dbRecordSet("ID")
	db_212_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	Set query1= Nothing
	If db_212_id1 <> "" and db_212_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 212", "ID: " & db_212_id1 & VBCRLF & "Status ID: " & db_212_status_id, "PASSED"
		If db_212_status_id = 4 Then
			Append_TestHTML StepCounter, "Validate Job status", "Expected:  " & db_212_status_id & VBCRLF & "Actual Status :4 ", "PASSED"
		
			''''''msgbox "Job 243 Successful"
			''''''msgbox db_id1
		End If
	else
		Append_TestHTML StepCounter, "Get ID from Job 212", "DB values not found. ID: " & db_212_id1 & VBCRLF & "Status ID: " & db_212_status_id, "FAILED"
	bFlag = False
		
	End If
	
	query2 = "select * from Job where JobTypeID = 262 order by 1 desc ;"
	Append_TestHTML StepCounter, "Verify Job 262",  query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id1 = dbRecordSet("ID")
	db_262_status_id = dbRecordSet("StatusID")
	Set dbRecordSet = Nothing
	
	If db_id1 <> "" and db_262_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 262", "ID: " & db_id1 & VBCRLF & "Status ID: " & db_262_status_id, "PASSED"
		If db_262_status_id = 4 Then
			Append_TestHTML StepCounter, "Validate Job status", "Expected:  " & db_262_status_id & VBCRLF & "Actual Status :4 ", "PASSED"
		
			''''''msgbox "Job 243 Successful"
			''''''msgbox db_id1
		End If
	else
		Append_TestHTML StepCounter, "Get ID from Job 262", "DB values not found. ID: " & db_id1 & VBCRLF & "Status ID: " & db_262_status_id, "FAILED"
	bFlag = False
		
	End If
	
End Function



Public Function VerifyDX600()
	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 206 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 206", query , "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id = dbRecordSet("ID")
	Set dbRecordSet = Nothing
	
	If db_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 206", "ID: " & db_id, "PASSED"
	else
		Append_TestHTML StepCounter, "Get ID from Job 206", "DB values not found. ID: " & db_id, "FAILED"
	bFlag = False
	End If
	
	query1 = "Update Job set NextRunDate = NULL  where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 291",query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	wait 10
	Set dbRecordSet = Nothing
	query2 = "select * from Job where ID = '" & db_id & "'"
	Append_TestHTML StepCounter, "Verify Job 206",  query2, "PASSED"
	Set dbRecordSet = execute_db_query(query2, 1,"SFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	db_id1 = dbRecordSet("ID")
	db_status_id = dbRecordSet("StatusID")
	
	If db_id1 <> "" and db_status_id <> "" Then
		Append_TestHTML StepCounter, "Get ID from Job 206", "ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "PASSED"
		If db_status_id = 4 Then
			Append_TestHTML StepCounter, "Validate Job status", "Expected:  " & db_status_id & VBCRLF & "Actual Status :4 ", "PASSED"
		
			''''''msgbox "Job 243 Successful"
			''''''msgbox db_id1
		End If
	else
		Append_TestHTML StepCounter, "Get ID from Job 206", "DB values not found. ID: " & db_id1 & VBCRLF & "Status ID: " & db_status_id, "FAILED"
	bFlag = False
		
	End If
	
End Function



Public Function VerifyDX451data()
	On error resume next
	
	bFlag = True
	
	query1 = "Select * from Site where ERPSiteNumber = '" & siteunqNo & "'"
	Append_TestHTML StepCounter, query1, "PASSED"
	Set dbRecordSet = execute_db_query(query1, 1,"SFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_SiteID = dbRecordSet("SiteID")
	db_DocNo= dbRecordSet("DateEstablished")
	
	Set dbRecordSet = Nothing
		
	If db_SiteID <> ""  Then
		Append_TestHTML StepCounter, "Validate Site entry table", db_SiteID & "- New entry created" , "PASSED"
		
	else
		Append_TestHTML StepCounter, "Validate Site entry table", "Fail to create entry" , "FAILED"
		bFlag=False
	End If
	
End Function



Public Function VerifyDX451Jobs(fileNameval)
	On error resume next
	
	bFlag = True
	
	query = "Select * from Job where JobTypeId = 294 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 294", query , "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_MASTER")
	wait 2
	
	db_id = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	
	Set dbRecordSet = Nothing
	
	If db_id = "4" and instrdb_inputxml(db_inputxml,fileNameval)>0 Then
		Append_TestHTML StepCounter, "Validate Job 294", "ID: " & db_id & "Input XML: " & db_inputxml , "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Job 294", "ID: " & db_id & "Input XML: " & db_inputxml , "FAILED"
		bFlag = False
	End If
	
	query = "Select * from Job where JobTypeId = 230 order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 230", query , "PASSED"
	Set dbRecordSet = execute_db_query(query, 1,"SFN_SHELL_MASTER")
	wait 2
	
	db_id = dbRecordSet("StatusID")
	db_inputxml = dbRecordSet("InputXml")
	
	Set dbRecordSet = Nothing
	
	If db_id = "3" and db_inputxml <> "" Then
		Append_TestHTML StepCounter, "Validate Job 230", "ID: " & db_id & "Input XML: " & db_inputxml , "PASSED"
	else
		Append_TestHTML StepCounter, "Validate Job 230", "ID: " & db_id & "Input XML: " & db_inputxml , "FAILED"
		bFlag = False
	End If
	
End Function



