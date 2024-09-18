Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	cERP = Get_Dictionary(ParamValDict,"ItemID" & "_" & iRowCount)
SNum="4719"

updateSNum = "2001"
batchdbName = "GFN_SHELL_SPRINTQA_NL_BATCH"
oltpdbName = "GFN_SHELL_SPRINTQA_NL_OLTP"
wwwdbName = "GFN_SHELL_SPRINTQA_NL_WWW"
reportsdbName = "GFN_SHELL_SPRINTQA_NL_REPORTS"

'msgbox "yes"
	'If strExecute = "Yes_"& countryName Then
	If strExecute = "Yes" Then
	cERP = getPreItemForReplication("Select * from SalesItemUnbilled;","GFN_SHELL_SPRINTQA_NL_BATCH","SalesItemID")
		'Call  DB_check((ShortName,cERP,"GFN_SHELL_SPRINTQA_NL_OLTP")
		'
	newvoucherNumber = Voucher_dbValidation(cERP,batchdbName,True)
		'msgbox newvoucherNumber
		oltpSNum = Voucher_dbValidation(cERP,oltpdbName,False)
		if newvoucherNumber = oltpSNum Then
			Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "& cERP & " in oltp DB"  ,"Expected Value: " & newvoucherNumber & VBCRLF & "Actual Value: " & oltpSNum, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "& cERP & " in oltp DB"  ,"Expected Value: " & newvoucherNumber & VBCRLF & "Actual Value: " & oltpSNum, "FAILED"		
		End If
		wwwSNum = Voucher_dbValidation(cERP,wwwdbName,False)
		if newvoucherNumber = wwwSNum Then
			Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "& cERP & " in WWW DB"  ,"Expected Value: " & newvoucherNumber & VBCRLF & "Actual Value: " & wwwSNum, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "& cERP & " in WWW DB"  ,"Expected Value: " & newvoucherNumber & VBCRLF & "Actual Value: " & wwwSNum, "FAILED"		
		End If
		reportsSNum = Voucher_dbValidation(cERP,reportsdbName,False)
		if newvoucherNumber = reportsSNum Then
			Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "& cERP & " in REPORTS DB"  ,"Expected Value: " & newvoucherNumber & VBCRLF & "Actual Value: " & reportsSNum, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate VoucherNumber of  SalesItemUnbilled  "& cERP & " in REPORTS DB"  ,"Expected Value: " & newvoucherNumber & VBCRLF & "Actual Value: " & reportsSNum, "FAILED"		
		End If
		updateSNum = originalSaleNum
		newvoucherNumber = Voucher_dbValidation(cERP,batchdbName,True)

	End If
next