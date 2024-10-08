  
  Function getPreItemForReplication(myQuery,dbName,myColVal)
 	
On Error Resume Next
bFlag = True
	query_customer = myQuery 
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	 preItemval = dictDbResultSet(myColVal)
	wait 3
	
	set dictDbResultSet = Nothing
	Set query_customer = Nothing
	
	If preItemval <> "" Then
		getPreItemForReplication = preItemval
	Else
	getPreItemForReplication = ""
	End If
	
 End Function
 Function getCustomerForReplication()
 	
On Error Resume Next
	query_customer = "Select * from Customer;"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_NL_OLTP")
	 custCERP = dictDbResultSet("CustomerERP")
	wait 3
	
	set dictDbResultSet = Nothing
	Set query_customer = Nothing
	
	If custCERP <> "" Then
		getCustomerForReplication = custCERP
	Else
	getCustomerForReplication = ""
	End If
	
 End Function
 Public originalcustSName,updateCSName
Function customer_dbValidation(myERP,dbName,UpdateCheck)
On Error Resume Next
bFlag = True
	query_customer = "Select * from Customer where CustomerERP = '"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
	wait 3
	 custCID = dictDbResultSet("CustomerID")
	custSName = dictDbResultSet("ShortName")
	
	set dictDbResultSet = Nothing
	Set query_customer = Nothing
	
	If UpdateCheck Then
		originalcustSName = custSName
		query_customer = "Update Customer Set ShortName ='" & updateCSName & "' where CustomerERP = '"&myERP&"';"
		Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
		set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
		wait 3
		set dictDbResultSet = Nothing
		Set query_customer = Nothing
		
		
		query_customer = "Select * from Customer where CustomerERP = '"&myERP&"'"
		Append_TestHTML StepCounter, "Execute Query",query_customer, "PASSED"
		set dictDbResultSet = execute_db_query(query_customer, 1, dbName)
		wait 3
		
		custCID = dictDbResultSet("CustomerID")
		custSName = dictDbResultSet("ShortName")
		set dictDbResultSet = Nothing
		Set query_customer = Nothing
		If custSName = updateCSName  Then
					Append_TestHTML StepCounter, "Validate Short Name of Customer "&myERP  ,"Expected Value: " & updateCSName & VBCRLF & "Actual Value: " &  custSName, "PASSED"
			
		Else
				Append_TestHTML StepCounter, "Validate Short Name of Customer "&myERP  ,"Fail to  update", "FAILED"
			bRunFlag = False
		End If
		
	End If
		
	If custSName <> "" Then
		customer_dbValidation = custSName
		
	Else
		customer_dbValidation = ""	
	End If
	
End Function


Function getSalesItemIDForReplication()
 	
On Error Resume Next
	query_num = "Select * from SalesItemUnbilled"
	Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
	set dictDbResultSet = execute_db_query(query_num, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
	wait 5
	 mysIID = dictDbResultSet("SalesItemID")
	
	set dictDbResultSet = Nothing
	Set query_num = Nothing
	
	If mysIID <> "" Then
		getSalesItemIDForReplication = mysIID
	Else
	getSalesItemIDForReplication = ""
	End If
	
 End Function
Public originalSaleNum,updateSNum
Function Voucher_dbValidation(myERP,dbName,UpdateCheck)
On Error Resume Next
bFlag = True
	query_num = "Select * from SalesItemUnbilled where SalesItemID = '"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
	set dictDbResultSet = execute_db_query(query_num, 1, dbName)
	wait 5
	 btID = dictDbResultSet("BatchID")
	custVNum = dictDbResultSet("VoucherNumber")
	
	set dictDbResultSet = Nothing
	Set query_num = Nothing
	
	If UpdateCheck Then
		originalSaleNum = custVNum
		query_num = "Update  SalesItemUnbilled Set VoucherNumber ='" & updateSNum & "' where SalesItemID = '"&myERP&"';"
		Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
		set dictDbResultSet = execute_db_query(query_num, 1, dbName)
		wait 5
		set dictDbResultSet = Nothing
		Set query_num = Nothing
		
		
		query_num = "Select * from SalesItemUnbilled where SalesItemID= '"&myERP&"'"
		Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
		set dictDbResultSet = execute_db_query(query_num, 1, dbName)
		wait 5
		
		btID = dictDbResultSet("BatchID")
		custVNum = dictDbResultSet("VoucherNumber")
		set dictDbResultSet = Nothing
		Set query_num = Nothing
		
		If custVNum = updateSNum  Then
				Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled   "&myERP  ,"Expected Value: " & updateSNum & VBCRLF & "Actual Value: " & custVNum, "PASSED"
			
		Else
				Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "&myERP  ,"Expected Value: " & updateSNum & VBCRLF & "Actual Value: " & custVNum, "FAILED"
			bRunFlag = False
		End If	
	End If
	
	If custVNum <> "" Then
		Voucher_dbValidation = custVNum
	Else
		Voucher_dbValidation = ""	
	End If
	
End Function

Public originalsiteSName,updateSSName
Function site_dbValidation(myERP,dbName,UpdateCheck)
On Error Resume Next
bFlag = True
	query_site = "Select * from Site where SiteID= '"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_site, "PASSED"
	set dictDbResultSet = execute_db_query(query_site, 1, dbName)
	wait 3
	' custCID = dictDbResultSet("CustomerID")
	siteSName = dictDbResultSet("ShortName")
	
	set dictDbResultSet = Nothing
	Set query_site = Nothing
	
	If UpdateCheck Then
		originalsiteSName = siteSName
		query_site = "Update Site Set ShortName ='" & updateSSName & "' where SiteID= '"&myERP&"';"
		Append_TestHTML StepCounter, "Execute Query",query_site, "PASSED"
		set dictDbResultSet = execute_db_query(query_site, 1, dbName)
		wait 3
		set dictDbResultSet = Nothing
		Set query_site = Nothing
		
		
		query_site = "Select * from Site where SiteID= '"&myERP&"'"
		Append_TestHTML StepCounter, "Execute Query",query_site, "PASSED"
		set dictDbResultSet = execute_db_query(query_site, 1, dbName)
		wait 3
		
		'custCID = dictDbResultSet("CustomerID")
		siteSName = dictDbResultSet("ShortName")
		set dictDbResultSet = Nothing
		Set query_site = Nothing
		
		If siteSName = updateSSName  Then
				Append_TestHTML StepCounter, "Validate Short Name of site "&myERP  ,"Expected Value: " & updateSSName & VBCRLF & "Actual Value: " &  siteSName, "PASSED"
			
		Else
				Append_TestHTML StepCounter, "Validate Short Name of site "&myERP  ,"Fail to  update", "FAILED"
			bRunFlag = False
		End If	
	End If
	
	If siteSName <> "" Then
		site_dbValidation = siteSName
	Else
		site_dbValidation = ""	
	End If
	
End Function

Public originalStatusID

Function Job_VerifyJobID(myERP,dbName,UpdateCheck)
On Error Resume Next
bFlag = True
	query_job = "Select * from Job where JobTypeID ='"&myERP&"'Order by 1 desc; "
	Append_TestHTML StepCounter, "Execute Query",query_job, "PASSED"
	set dictDbResultSet = execute_db_query(query_job, 1, dbName)
	wait 3
	existID = dictDbResultSet("StatusID")
	'set dictDbResultSet = Nothing
	'Set  query_job= Nothing	
	If UpdateCheck Then
		If existID = "3" Then
			Append_TestHTML StepCounter,  "Verify the job is in running state"," Job '"&myERP&"' is in running  state with StatusID='"&existID&"'", "PASSED"
		else
			Append_TestHTML StepCounter, "Verify the job is in running state"," Job '"&myERP&"' is not in running  state with StatusID='"&existID&"'", "FAILED"
			bRunFlag = False
		end if 
	End If
	If existID <> "" Then
		Job_VerifyJobID = existID
	Else
		Job_VerifyJobID = ""
	End If
	
End  Function


Function VerifyRGX_Value(myERP,dbName,UpdateCheck)
On Error Resume Next
bFlag = True
	query_rgx = "Select * from RegEx Where RegExID='"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_rgx, "PASSED"
	set dictDbResultSet = execute_db_query(query_rgx, 1,dbName)
	wait 3
	value = dictDbResultSet("RegExValue")
	If value = UpdateCheck  Then
		Append_TestHTML StepCounter,  "Verify  value"," value '"&myERP&"' Successfully verified with REgExValue='"&value&"'", "PASSED"
	else
		Append_TestHTML StepCounter, "Verify value"," value '"&myERP&"' Successfully not verified with REgExValue "&value&" -" & UpdateCheck , "FAILED"
		bRunFlag = False
	End if
end function
'Function VerifyRGX_Value1(myERP,dbName)
'On Error Resume Next
'	query_rgx = "select * from RegEx Where RegExID='"&myERP&"'"
'	Append_TestHTML StepCounter, "Execute Query",query_rgx, "PASSED"
'	wait 3
'	set dictDbResultSet = execute_db_query(query_rgx, 1, "GFN_SHELL_SPRINTQA_MASTER")
'	value1 = dictDbResultSet("RegExValue")
'	If value1 = "(?i)DX053_([0-9]{3})_([0-9]{6}_([0-9]{8})_([0-9]{6}))(.json)"  Then
'		Append_TestHTML StepCounter,  "Verify  value"," value '"&myERP&"' Successfully verified with REgExValue='"&value1&"'", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Verify value"," value '"&myERP&"' Successfully not verified with REgExValue='"&value1&"'", "FAILED"
'	End if
'End  Function
'Function VerifyRGX_Value3(myERP,dbName)
'On Error Resume Next
'	query_rgx = "select * from RegEx Where RegExID='"&myERP&"'"
'	Append_TestHTML StepCounter, "Execute Query",query_rgx, "PASSED"
'	wait 3
'	set dictDbResultSet = execute_db_query(query_rgx, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
'	value2 = dictDbResultSet("RegExValue")
'	If value2 = "(?i)DX350_([0-9]{8})_([0-9]{3,4}_([0-9]{6})_([0-9]{8})_([0-9]{6}))(.dat)"  Then
'		Append_TestHTML StepCounter,  "Verify  value"," value '"&myERP&"' Successfully verified with REgExValue='"&value2&"'", "PASSED"
'	else
'		Append_TestHTML StepCounter, "Verify value"," value '"&myERP&"' Successfully not verified with REgExValue='"&value2&"'", "FAILED"
'	End if
'End  Function

'Public  originalSaleNum,updateSNum
Function Batch_dbValidation(myERP,dbName,UpdateCheck)
On Error Resume Next
bFlag = True
	query_num = "Select* from TranBatchFile where BatchID= '"&myERP&"'"
	Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
	set dictDbResultSet = execute_db_query(query_num, 1, dbName)
	wait 5
	 btID = dictDbResultSet("BatchID")
	custVNum = dictDbResultSet("TranBatchTypeDesc")
	
	set dictDbResultSet = Nothing
	Set query_num = Nothing
	
	If UpdateCheck Then
		originalSaleNum = custVNum
		query_num = "Update TranBatchFile Set TranBatchTypeDesc ='" & updateSNum & "' where BatchID = '"&myERP&"';"
		Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
		set dictDbResultSet = execute_db_query(query_num, 1, dbName)
		wait 5
		set dictDbResultSet = Nothing
		Set query_num = Nothing
		
		
		query_num = "Select* from TranBatchFile where BatchID= '"&myERP&"'"
		Append_TestHTML StepCounter, "Execute Query",query_num, "PASSED"
		set dictDbResultSet = execute_db_query(query_num, 1, dbName)
		wait 5
		
		btID = dictDbResultSet("BatchID")
		custVNum = dictDbResultSet("TranBatchTypeDesc")
		set dictDbResultSet = Nothing
		Set query_num = Nothing
		
		If custVNum = updateSNum  Then
				Append_TestHTML StepCounter, "Validate TranBatchTypeDesc of  TranBatchFile   "&myERP  ,"Expected Value: " & updateSNum & VBCRLF & "Actual Value: " & custVNum, "PASSED"
			
		Else
				Append_TestHTML StepCounter, "Validate TranBatchTypeDesc of  TranBatchFile  "&myERP  ,"Expected Value: " & updateSNum & VBCRLF & "Actual Value: " & custVNum, "FAILED"
			bRunFlag = False
		End If	
	End If
	
	If custVNum <> "" Then
		Batch_dbValidation = custVNum
	Else
		Batch_dbValidation = ""	
	End If
	
End Function





	











