Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")

For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	cERP = Get_Dictionary(ParamValDict,"ItemID" & "_" & iRowCount)
SNum="4719"

updateSNum = "DX028Job"
importdbName="GFN_SHELL_SPRINTQA_IMPORT"
batchdbName = "GFN_SHELL_SPRINTQA_NL_BATCH"


'msgbox "yes"
	'If strExecute = "Yes_"& countryName Then
	If strExecute = "Yes" Then
		'Call  DB_check((ShortName,cERP,"GFN_SHELL_SPRINTQA_NL_OLTP")
		'.
	cERP = getPreItemForReplication("Select* from TranBatchFile;",batchdbName,"BatchID")
	newTranBatchTypeDesc = Batch_dbValidation(cERP,importdbName,True)
		'msgbox newTranBatchTypeDesc
		wait 10
		batchSNum = Batch_dbValidation(cERP,batchdbName,False)
		if newTranBatchTypeDesc = batchSNum Then
			Append_TestHTML StepCounter, "Validate TranBatchTypeDesc of  TranBatchFile  "& cERP & " in batch DB"  ,"Expected Value: " & newTranBatchTypeDesc & VBCRLF & "Actual Value: " & batchSNum, "PASSED"
		Else
			Append_TestHTML StepCounter, "Validate TranBatchTypeDesc of  TranBatchFile  "& cERP & " in batch DB"  ,"Expected Value: " & newTranBatchTypeDesc & VBCRLF & "Actual Value: " & batchSNum, "FAILED"		
			bRunFlag = False
		End If

		updateSNum = originalSaleNum
		newTranBatchTypeDesc = Batch_dbValidation(cERP,importdbName,True)

	End If
next