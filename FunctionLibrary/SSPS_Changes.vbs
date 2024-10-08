Function sspsChanges_fieldEditable_cardParameter(useFleetPIN,selectedPIN)
On error resume next
	bFlag = True
	Call pageNavigation("Card Parameters","Link_CardParameters")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebCheckbox_SelectedPIN")  Then
		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Card Parameters' screen", "PASSED"		
		Call operateOnCheckBox("WebCheckbox_SelectedPIN","Selected PIN",selectedPIN)
		Call validate_checkBox_enabled("WebCheckbox_UseFleetPIN","UseFleetPIN")
		Call operateOnCheckBox("WebCheckbox_UseFleetPIN","UseFleetPIN",useFleetPIN)
		Call validate_textbox_enabled("WedEdit_CustomerFleetPIN","Customer Fleet PIN")
		Call validate_textbox_enabled("WedEdit_SSPSFormNumber","SSPS Form Number")		
	else
		Append_TestHTML StepCounter, "Fee Rule Tier", "User does not navigates to 'Card Parameters' screen", "FAILED"
	End  If
End  Function

Function sspsChanges_fieldNotVisible_cardParameter(useFleetPIN,selectedPIN)
On error resume next
	bFlag = True
	Call pageNavigation("Card Parameters","Link_CardParameters")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebCheckbox_SelectedPIN")  Then
		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Card Parameters' screen", "PASSED"		
		Call operateOnCheckBox("WebCheckbox_SelectedPIN","Selected PIN",selectedPIN)
		Call validate_checkBox_disabled("WebCheckbox_UseFleetPIN","UseFleetPIN")
		Call validate_textbox_NotVisible("WedEdit_CustomerFleetPIN","Customer Fleet PIN")
		Call validate_textbox_NotVisible("WedEdit_SSPSFormNumber","SSPS Form Number")		
		
	else
		Append_TestHTML StepCounter, "Fee Rule Tier", "User does not navigates to 'Card Parameters' screen", "FAILED"
	End  If
End  Function

Function sspsChanges_fieldNotEditable_cardParameter(useFleetPIN,selectedPIN)
On error resume next
	bFlag = True
	Call pageNavigation("Card Parameters","Link_CardParameters")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebCheckbox_SelectedPIN")  Then
		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Card Parameters' screen", "PASSED"		
		Call operateOnCheckBox("WebCheckbox_SelectedPIN","Selected PIN",selectedPIN)
		Call validate_checkBox_enabled("WebCheckbox_UseFleetPIN","UseFleetPIN")
		Call operateOnCheckBox("WebCheckbox_UseFleetPIN","UseFleetPIN",useFleetPIN)

		Call validate_textbox_disabled("WedEdit_CustomerFleetPIN","Customer Fleet PIN")
		Call validate_textbox_disabled("WedEdit_SSPSFormNumber","SSPS Form Number")		
	else
		Append_TestHTML StepCounter, "Fee Rule Tier", "User does not navigates to 'Card Parameters' screen", "FAILED"
	End  If
End  Function

Function sspsChanges_DB_queries(hideSSPSDetails,SysVarID)
On Error Resume Next
	flag = True
	query_sysvarcolco = "Select * from SysVarColco where SysVarID = '"&SysVarID&"'"
	Append_TestHTML StepCounter, "Validate 'Value' column in 'SysVarColco' table",query_sysvarcolco, "PASSED"
	set dictDbResultSet = execute_db_query(query_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_value = dictDbResultSet("Value")
	Select Case cdbl(hideSSPSDetails)
		Case 0:
			If db_value <> cdbl(hideSSPSDetails) Then
				updatequery_sysvarcolco = "Update SysVarColco Set Value = 0 where SysVarID = "&SysVarID&";"
				Append_TestHTML StepCounter, "Update 'Value' column 'SysVarColco' table",updatequery_sysvarcolco, "PASSED"
				Call update_db_query(updatequery_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
				
'				db_value = dictDbResultSet("Value")
				wait 2
			else
				Append_TestHTML StepCounter, "SysVarColco table","'Value' column value is '0' in 'SysVarColco' table, as expected", "PASSED"
			End If
		Case 1:
			If db_value <> cdbl(hideSSPSDetails) Then
				updatequery_sysvarcolco = "Update SysVarColco Set Value = 1 where SysVarID = "&SysVarID&";"
				Append_TestHTML StepCounter, "Update 'Value' column 'SysVarColco' table",updatequery_sysvarcolco, "PASSED"
				Call update_db_query(updatequery_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'				db_value = dictDbResultSet("Value")
				wait 2
			else
				Append_TestHTML StepCounter, "SysVarColco table","'Value' column value is '1' in 'SysVarColco' table, as expected", "PASSED"
			End If
			wait 10
	End Select
	
End Function

Function preCondition_colcoPinAdvice(ColcoID)

On Error Resume Next
	bFlag = True
	dbquery = "Select * from SysVarColCo where SysvarID = 145;"
'	Append_TestHTML StepCounter, "Execute Card Addreess query",dbquery, "PASSED"
	Set dbRecordSet = execute_db_query(dbquery, 1,SFN_SHELL_SPRINTQA_ID_OLTP)
	wait 2
	db_ColcoID = dbRecordSet("ColcoID")
	Set dbRecordSet = Nothing
	Set dbquery = Nothing
	updatequery_ColcoPINAdvice = "Update ColcoPINAdvice Set DefaultForNewTopLevelCustomer = 1 where PINAdviceTypeID =2 and ColcoID = "&db_ColcoID &";"
	Append_TestHTML StepCounter, "Update 'Value' column 'SysVarColco' table",updatequery_ColcoPINAdvice, "PASSED"
	Call update_db_query(updatequery_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	db_value = dictDbResultSet("Value")
	wait 2
	updatequery_SysVarColco = "Update SysVarColco Set value = 1 where SysvarID = 145"
	Append_TestHTML StepCounter, "Update 'Value' column in 'SysVarColco' table",updatequery_SysVarColco, "PASSED"
	Call update_db_query(updatequery_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	db_value = dictDbResultSet("Value")
	wait 2
	updatequery_SysVarColco = "Update SysVarColco Set value = 1 where SysvarID = 146"
	Append_TestHTML StepCounter, "Update 'Value' column in 'SysVarColco' table",updatequery_SysVarColco, "PASSED"
	Call update_db_query(updatequery_sysvarcolco, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
'	db_value = dictDbResultSet("Value")
	wait 2
End Function