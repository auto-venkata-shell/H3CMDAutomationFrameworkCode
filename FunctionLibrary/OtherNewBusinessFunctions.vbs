
'-------------------------------------------
Public Function fieldPresent(field,name)
On error resume next
	bFlag = True
	If VerifyWebObjectExist ("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  field) Then
		Append_TestHTML StepCounter, "Site Details page", "Expected "&name&" field is present", "PASSED"				
	else
		Append_TestHTML StepCounter, "Site Details page", "Expected "&name&" field is not present", "FAILED"	
	End If
End Function



Public Function launchApplicationnewGFN()
	On error resume next
	If bSingleSignOnFlag = True  Then		
		systemutil.CloseProcessByName "msedge.exe"	
		wait 2		
		systemutil.Run "msedge.exe", url
		wait 4
		If Browser("name:=ntp\.msn\.com").Exist or (Window("micClass:=Browser","name:=Home Page").Exist = False) Then
		Browser("creationTime:=0").WinObject("abs_x:=4","abs_y:=104").Click
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
		End If
		Append_TestHTML StepCounter,"Application launch","Launch the URL '"&url&"'","PASSED"
		wait 2		
		Window("micClass:=Browser","name:=Home Page").Maximize
		wait 2
		flag = False
		rcount = Browser("name:=Home Page").Page("title:=Home Page").WebTable("html id:=MainContent_grdServers").Rowcount
		For i = 1 To rcount Step 1
			countryTab = Browser("name:=Home Page").Page("title:=Home Page").WebTable("html id:=MainContent_grdServers").GetCellData(i,1)
			If Instr(countryTab, countryName) > 0 Then
				Browser("name:=Home Page").Page("title:=Home Page").WebTable("html id:=MainContent_grdServers").childItem(i,1,"WebElement",0).click
				browserProp = Browser("creationTime:=1").GetROProperty("name")
				dynamicBrowser = "name:"&browserProp
				strBrowser = "name:="&browserProp
				pageProp =  Browser("creationTime:=1").Page("creationTime:=1").GetROProperty("title")
				dynamicPage = "title:"& pageProp
				strPage = "title:="& pageProp
				Window("micClass:=Browser","name:="&browserProp).Maximize
				flag = True
				Exit For
			End If
		Next
		If flag = False Then
			Append_TestHTML StepCounter,"Country Name","Expected country name is not found in the application","FAILED"				
		End If
		If objORDict.Exists("WebBrowser_PTShell_INDO") Then
		  objORDict.item("WebBrowser_PTShell_INDO") = dynamicBrowser
		End If
		If objORDict.Exists("WebPage_PTShell_INDO") Then
		  objORDict.item("WebPage_PTShell_INDO") = dynamicPage 
		End If
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","WebEdit_Search") Then		
			Append_TestHTML StepCounter,"Application launch","Application is launched successfully","PASSED"
			bSignOnFlag = True
		else
			Append_TestHTML StepCounter,"Application launch failed","Application launch failed","FAILED"
			bRunFlag = False
		End If 		
	End if 	
End Function


'----------------------------------------New Functions


Function pageNavigation(page_name,link)
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Search", "Set", page_name
	Append_TestHTML StepCounter, "Page navigation", "Enter the page name '"&page_name&"' in Search Menu", "PASSED"	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", link, "Click", ""
	Append_TestHTML StepCounter, "Page navigation", "Click on page name '"& page_name&"' in the Search Menu drop_down", "PASSED"	
End Function

Function enterTextbox_value(webEdit,field_name,value)
On error resume next
	bFlag = True
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "Highlight", ""
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "Click", ""
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "Set", value
	Append_TestHTML StepCounter, "Enter_textbox_value", field_name&" '"&value&"' is entered successfully", "PASSED"
End  Function

Function enterWebList_value(webList,field_name,value)
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webList, "RadioSelect", value
	Append_TestHTML StepCounter, "Enter_listbox_value", field_name&": "&value&" value is selected successfully", "PASSED"	
	
End  Function

Function operateOnCheckBox(chkBoxObj,chkBoxName,chkPropVal)
On error resume next
	bFlag = True
	ui_checkProp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "",chkBoxObj, "GetROProperty", "checked")
	Select Case chkPropVal 
		Case "0":
			If ui_checkProp Then
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", chkBoxObj, "Click", ""	
			End If
			Append_TestHTML StepCounter, "Disable checkBox", "Disable the '"&chkBoxName&"' check box", "PASSED"	
		Case "1":
			If ui_checkProp = 0 Then
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", chkBoxObj, "Click", ""	
			End If
			Append_TestHTML StepCounter, "Enable checkBox", "Enable the '"&chkBoxName&"' check box", "PASSED"	
	End Select	
End Function

Function operateOnTblElement(tblID,tblName,cellData)
On error resume next
	bFlag = True
	Flag  = False
	wait 2
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:="&tblID)
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Select table item", "User is displayed with '"&tblName & "' table", "PASSED"				
		rcount = tableObj.Rowcount
		If rcount > 1 Then
'			cellData = tableObj.GetCelldata(2,1)
			Set desc = description.Create
			desc("micclass").value = "WebElement"
			Set childObj = tableObj.ChildObjects(desc)
			chdCount =  childObj.count
			For i = 1 To chdCount Step 1
				ui_cellData =  childObj(i).GetRoproperty("innertext")
				If Trim(cellData) = Trim(ui_cellData) Then
					Append_TestHTML StepCounter, "Select table item", "Search the table data '"&cellData&"' in '"&tblName&"' table", "PASSED"
					childObj(i).highlight
					 childObj(i).click
					 Append_TestHTML StepCounter, "Select table item", "Click on table data '"&cellData&"' in '"&tblName&"' table", "PASSED"
					 Flag  = True
					 Exit For
				End If
			Next
			If Flag  = False Then
				Append_TestHTML StepCounter, "Select table item", "Expected table data '"&cellData&"' is not found in '"&tblName&"' table", "FAILED"
			End If
		End If
	else
		Append_TestHTML StepCounter, "Select table item", "User is not displayed with '"&tblName&"' table",  "FAILED"
	end if
	wait 3
End Function



Public Function navigateWithoutStartMenu(mainoptionobj,suboptionobj,verifyoptionobj,mainMenuName,subMenuName)
	On error resume next
	bFlag = True
	If suboptionobj = "" Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", mainoptionobj, "Click", ""
		Append_TestHTML StepCounter, "Navigate Menu","Click on '"&mainMenuName&"'", "PASSED"
	Else
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "",mainoptionobj, "Mover", ""
		Append_TestHTML StepCounter, "Navigate Menu","Mouse hover on '"&mainMenuName&"' link", "PASSED"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "",suboptionobj, "Mover", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", suboptionobj, "Click", ""
		Append_TestHTML StepCounter, "Navigate Menu","Click on '"&subMenuName&"' link", "PASSED"
	End If
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", verifyoptionobj)  Then
		Append_TestHTML StepCounter, "Navigate Menu","User navigates to '"&subMenuName&"' screen", "PASSED"
	Else
		Append_TestHTML StepCounter, "Navigate Menu","User does not navigates to '"&subMenuName&"' screen", "FAILED"
		bFlag = False
	End IF		
'	'''''msgbox "end start"
	navigateWithoutStartMenu = bFlag
	wait 3
End Function



Function click_on_save()
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"	
	If VerifyWebObjectExist ("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebELement_Yes") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_Yes", "Highlight", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebELement_Yes", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Yes' in popup ", "PASSED"
	End  If	
End Function

Public Function customerSearch1(customerERP_id)
	On error resume next
	bFlag = True
	Call  navigateStartMenu("Link_Customers","Link_SearchforCustomer","WebLIST_Role")
'	'''''msgbox "customerSearch"

	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLIST_Role")  Then
		
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLIST_Role", "RadioSelect", "Customer ERP Number:"	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_EndValueBox", "Set", customerERP_id
			Append_TestHTML StepCounter,"Search for Customer","Enter the Customer ERP ID: "& customerERP_id ,"PASSED"			
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""			
			Append_TestHTML StepCounter,"Search for Customer","Click on Search" ,"PASSED"
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "Link_CustomerSummary")  Then
				Append_TestHTML StepCounter,"Customer Summary","User successfully navigated to 'Customer Summary' screen" ,"PASSED"
			End  If
	Else
		Append_TestHTML StepCounter,"Search for Customer","Fail to get customer search page to insert "& customerERP_id ,"FAILED"
		bFlag = False		
	End  If
'	'''''msgbox "end customersearch"
	customerSearch1 = bFlag
End Function


Public Function CreateCustPriceRule1(cust_erp)

	On error resume next
	
	bFlag = True
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_NewPriceRule", "Click", ""
	wait 2
		Append_TestHTML StepCounter, "Create new price rule", "Navigating to new price rule page", "PASSED"
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_SearchPriceRule", "Click", ""
	wait 2
		Append_TestHTML StepCounter, "Search for existing price rule", "Clicked on search button", "PASSED"
	
'	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_PriceRuleDisc", "Click", ""
'
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebList_PriceRuleCat", "RadioSelect", "#1"
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Search", "Click", ""
	wait 2
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_RetailDefault", "Click", ""

	wait 2
		Append_TestHTML StepCounter, "Click on Price Rule Description", "Rule selected", "PASSED"
	
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_Save") Then
		Append_TestHTML StepCounter, "Create price rule", "Adding price rule", "PASSED"
	else
		Append_TestHTML StepCounter, "Create price rule", "Adding price rule failed", "FAILED"
	End If
	
'	refPrice = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RefPrice","GetROProperty","default value")
	
	refPrice = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_PriceRuleread","GetROProperty","default value")
	
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	
	wait 3   ' waiting for the DB update
	
	query = "select * from CustomerPriceRule where CustomerID = (select customerid from Customer where CustomerERP = '" & cust_erp & "');"
	dbName = appName & "_SHELL_SPRINTQA_" & countryCode & "_OLTP"
	set dbRecordSet = execute_db_query(query, 1, dbName)
	wait 2
	priceRuleID = dbRecordSet("PriceRuleID")
	
	query2 = "select * from PriceRule where PriceRuleID = '" & priceRuleID & "';"
	dbName = appName & "_SHELL_SPRINTQA_" & countryCode & "_OLTP"
	set dbRecordSet2 = execute_db_query(query2, 1, dbName)
	wait 2
	priceRuleDesc = dbRecordSet2("PriceRuleDescription")
	
	If instr(priceRuleDesc, refPrice)<>0 Then
		Append_TestHTML StepCounter,"Price Rule Validation","Expected Value: " & priceRuleDesc & VBCRLF & "Actual Value: " & refPrice ,"PASSED"
	else
		Append_TestHTML StepCounter,"Price Rule Validation","Expected Value: " & priceRuleDesc & VBCRLF & "Actual Value: " & refPrice ,"FAILED"
		bRunFlag = False
	End If
	
End Function

Function objectClick(objlink,objName)
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", objlink, "Click", ""
	Append_TestHTML StepCounter, "Click on object", "Click on '"&objName&"'", "PASSED"
End Function

Function validate_resultTable(strDesc)
On error resume next
	bFlag = True
	Call pageNavigation("Search Price Rule","WebLink_SearchPriceRule")
	Call objectClick("WebLink_Refresh","Refresh link")
	Call enterTextbox_value("WedEdit_PriceRuleDescription","Price Rule Description",strDesc)
	Call objectClick("WebLink_Search","Search link")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebTable_ResultTable") Then
		Append_TestHTML StepCounter, "Create price rule", "User is displayed with 'Search result table'", "PASSED"
		Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*mvPriceRules_grdResults")
		Set desc = Description.Create
		desc("micclass").value = "WebElement"
		Set chdObj = tableObj.ChildObjects(desc)
		For i = 1 To chdObj.count Step 1
			ui_strDesc = chdObj(i).GetRoProperty("innertext")
			If ui_strDesc = strDesc Then
				Append_TestHTML StepCounter, "Create price rule", "New Price Rule with the Price Rule description '"&strDesc&"' is created successfully", "PASSED"
			 	chdObj(i).Highlight
				 chdObj(i).click
				 Append_TestHTML StepCounter, "Create price rule", "Click on Price Rule description '"&strDesc&"'", "PASSED"
				 flag = True
				 Exit For
			End If
		Next								
		If flag = false Then
			Append_TestHTML StepCounter, "Create price rule", "New Price Rule with the Price Rule description '"&strDesc&"' is not created successfully", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "Create price rule", "User is not displayed with 'Search result table'", "FAILED"
	End  if
End Function

Function validate_DB(strDesc)
On error resume next
	bFlag = True
	query = "Select * from PriceRule where PriceRuleDescription = '"&strDesc & "';"
	set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2	
	db_priceRule = dictDbResultSet("PriceRuleDescription")
	Set dictDbResultSet = Nothing
	If strDesc = db_priceRule  Then
		Append_TestHTML StepCounter,"Database validation","Expected Value: " & strDesc & VBCRLF & "Actual Value: " & db_priceRule ,"PASSED"
	else
		Append_TestHTML StepCounter,"Database validation","Expected Value: " & strDesc & VBCRLF & "Actual Value: " & db_priceRule ,"FAILED"
		bFlag = False
	End If
End Function

Function verifyTblValueExist(tblID,tblName,cellData)
On error resume next
	bFlag = True
	flagValueprop  = False
	wait 2
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:="&tblID)
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Select table item", "User is displayed with '"&tblName & "' table", "PASSED"
		Append_TestHTML StepCounter, "verify table item", "Search for the table data '"&cellData&"' in '"&tblName&"' table", "PASSED"		
		rcount = tableObj.Rowcount
		For i = 1 To rcount Step 1
			ui_cellData = tableObj.GetCelldata(i,4)
			If Trim(cellData) = Trim(ui_cellData) Then				
				Append_TestHTML StepCounter, "verify table item", "The table data '"&cellData&"' is found in the '"&tblName&"' table, as expected", "PASSED"		
				 flagValueprop  = True
				 Exit For
			End If
		Next
	else
		Append_TestHTML StepCounter, "verify table item", "User is not displayed with '"&tblName&"' table",  "FAILED"
	end if
	wait 3
End Function

Function getEffectiveDate()
On error resume next
	bFlag = True
	sDay =  day(Date)
	sMonth = month(Date)
	sYear = Year(Date)
	sDay = cdbl(sDay) + 1
	If len(sDay) = 1 Then
		sDay = "0"&sDay
	End If
	If len(sMonth) = 1 Then
		sMonth = "0"&sMonth
	End If
	getEffectiveDate = sDay & "/" & sMonth & "/" & sYear
End Function

Function verifyTblElement(tblID,tblName,strDesc)
On error resume next
	bFlag = True
	Flag  = False
	wait 2
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_NewFeeRule")  Then
		Append_TestHTML StepCounter, "Verify table item", "User navigates to 'Search for Fee Rules' screen", "PASSED"
		Call enterTextbox_value("WedEdit_FeeRuleDescp","Fee Rule Description",strDesc)
		Call objectClick("WebLink_Search","Search link")
		wait 5
		Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:="&tblID)
		If tableObj.exist Then
			Append_TestHTML StepCounter, "Verify table item", "User is displayed with '"&tblName & "' table for the entered search criteria", "PASSED"				
			rcount = tableObj.Rowcount
			If rcount > 1 Then
				For i = 1 To rcount Step 1
					ui_strDesc =  tableObj.Getcelldata(i,3)
					If Trim(ui_strDesc) =  strDesc Then
						Append_TestHTML StepCounter, "Verify table item", "Expected New Rule Description '"&strDesc&"' is exist in the '"&tblName&"' table", "PASSED"					
						flag = True
						Exit For
					End  If
				Next
			else
				Append_TestHTML StepCounter, "Verify table item", "No data displayed for the entered search criteria", "FAILED"
			End  If
			If flag = false Then
				Append_TestHTML StepCounter, "Verify table item", "Expected New Rule Description '"&strDesc&"' is not exist in the '"&tblName&"' table", "FAILED"				
			End If
		else
			Append_TestHTML StepCounter, "Verify table item", "User is not displayed with '"&tblName&"'",  "FAILED"
		end if
	else
		Append_TestHTML StepCounter, "Verify table item", "User does not navigated to 'Search for Fee Rules' screen", "FAILED"
	End  If
End Function

Public Function navigateWithStartMenu(mainoptionobj,suboptionobj,verifyoptionobj,mainMenuName,subMenuName)
	On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "Link_Start", "Mover", ""
	Append_TestHTML StepCounter, "Navigate Menu","Mouse hover on 'Start' menu", "PASSED"
	If suboptionobj = "" Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", mainoptionobj, "Click", ""
		Append_TestHTML StepCounter, "Navigate Menu","Click on '"&mainMenuName&"'", "PASSED"
	Else
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "",mainoptionobj, "Mover", ""
		Append_TestHTML StepCounter, "Navigate Menu","Mouse hover on '"&mainMenuName&"' link", "PASSED"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", suboptionobj, "Click", ""
		Append_TestHTML StepCounter, "Navigate Menu","Click on '"&subMenuName&"' link", "PASSED"
	End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", verifyoptionobj)  Then
		Append_TestHTML StepCounter, "Navigate Menu","User navigates to '"&subMenuName&"' screen", "PASSED"
	Else
		Append_TestHTML StepCounter, "Navigate Menu","User does not navigates to '"&subMenuName&"' screen", "FAILED"
		bFlag = False
	End IF	

	navigateWithStartMenu = bFlag
	wait 3
End Function

Function validate_savePopup()
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Save_SuccessMsg")  Then
		Append_TestHTML StepCounter, "save item", "User is displayed with 'Your item has been saved successfully' message", "PASSED"
	else
		Append_TestHTML StepCounter, "save item", "User is not displayed with 'Your item has been saved successfully' message", "FAILED"
	end  If
End Function


	
Function RunJob(JobTypeID)
On error resume next
	bFlag = True
	query_job = "Select * from Job where JobTypeID = '"&JobTypeID&"' order by 1 desc"
	Append_TestHTML StepCounter, "Run Job",query_job, "PASSED"
	set dictDbResultSet = execute_db_query(query_job, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	jobID= dictDbResultSet("ID")
	
	query_JobUpt = "Update job set NextRunDate = null where ID = '"&jobID&"'"
	Append_TestHTML StepCounter, "Update Job table",query_JobUpt, "PASSED"
	Call update_db_query(query_JobUpt, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 10
	query_Runjob = "Select * from Job where ID = '"&jobID&"'"
	Append_TestHTML StepCounter, "Run Job",query_Runjob, "PASSED"
	set dictDbResultSet = execute_db_query(query_Runjob, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	db_statusID = dictDbResultSet("StatusID")	
	
	If db_statusID = "0" Then
		set dictDbResultSet = execute_db_query(query_Runjob, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
		wait 10
		db_statusID = dictDbResultSet("StatusID")
	End If
	If db_statusID = "4"  Then
		Append_TestHTML StepCounter, "Get JobID ","As expected, job Status ID changed to '4'", "PASSED"		
	End If		
End Function

Function validate_textbox_enabled(webEdit,field_name)
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", webEdit)  Then
		Append_TestHTML StepCounter, "verify object enable", field_name &" text box is exist in the screen", "PASSED"
		
		ui_disableProp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "GetROProperty", "disabled")
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "GetROProperty", "disabled" 
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "Highlight", "" 
		If ui_disableProp = 0 Then
			Append_TestHTML StepCounter, "verify object enable", field_name &" text box is enabled", "PASSED"	
		else
			Append_TestHTML StepCounter, "verify object enable", field_name &" text box is disabled",  "FAILED"	
		End If
	else
		Append_TestHTML StepCounter, "verify object enable", field_name &" text box is not exist in the screen", "FAILED"
	End  If
End  Function

Function validate_checkBox_enabled(webcheckBox,field_name)
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", webcheckBox)  Then
		Append_TestHTML StepCounter, "verify object enable", field_name &" check box is exist in the screen", "PASSED"		
		ui_disableProp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webcheckBox, "GetROProperty", "disabled")
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webcheckBox, "GetROProperty", "disabled" 

		If ui_disableProp = 0 Then
			Append_TestHTML StepCounter, "verify object enable", field_name &" check box is enabled", "PASSED"	
		else
			Append_TestHTML StepCounter, "verify object enable", field_name &" check box is disabled",  "FAILED"	
		End If
	else
		Append_TestHTML StepCounter, "verify object enable", field_name &" check box is not exist in the screen", "FAILED"
	End  If
End  Function

Function validate_checkBox_disabled(webcheckBox,field_name)
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", webcheckBox)  Then
		Append_TestHTML StepCounter, "verify object enable", field_name &" check box is exist in the screen", "PASSED"		
		ui_disableProp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webcheckBox, "GetROProperty", "disabled")
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webcheckBox, "GetROProperty", "disabled" 

		If ui_disableProp = 1 Then
			Append_TestHTML StepCounter, "verify object enable", field_name &" check box is disabled",  "PASSED"
		else
			Append_TestHTML StepCounter, "verify object enable", field_name &" check box is enabled",   "FAILED"	
		End If
	else
		Append_TestHTML StepCounter, "verify object enable", field_name &" check box is not exist in the screen", "FAILED"
	End  If
End  Function

Function validate_textbox_disabled(webEdit,field_name)
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", webEdit)  Then
		Append_TestHTML StepCounter, "verify object enable", field_name &" text box is exist in the screen", "PASSED"		
		ui_disableProp = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "GetROProperty", "disabled")
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "GetROProperty", "disabled" 
'		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", webEdit, "Highlight", "" 
		If ui_disableProp = 1 Then
			Append_TestHTML StepCounter, "verify object enable", field_name &" text box is disabled", "PASSED"	
		else
			Append_TestHTML StepCounter, "verify object enable", field_name &" text box is enabled",  "FAILED"	
		End If
	else
		Append_TestHTML StepCounter, "verify object enable", field_name &" text box is not exist in the screen", "FAILED"
	End  If
End  Function

Function validate_textbox_NotVisible(webEdit,field_name)
On error resume next
	bFlag = True
	If Not(VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", webEdit))  Then
		Append_TestHTML StepCounter, "verify object notVisible", field_name &" text box is not visible in the screen", "PASSED"		
	else
		Append_TestHTML StepCounter, "verify object enable", field_name &" text box is exist in the screen", "FAILED"
	End  If
End  Function

Function click_on_save_element()
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"	
	If VerifyWebObjectExist ("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Highlight", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Save' element in popup ", "PASSED"
	End  If	
End Function

Function enterTextbox_value_desc(htmlID,field_name,value)
On error resume next
	bFlag = True
	Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*"&htmlID).click
	Browser("creationTime:=1").Page("creationTime:=1").WebEdit("html id:=.*"&htmlID).set value
	Append_TestHTML StepCounter, "Enter value", "Enter the '"&field_name&"' value as '"&value&"'", "PASSED"
End  Function

Function click_on_saveElement()
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"	
	If VerifyWebObjectExist ("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebElement_Save") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Highlight", ""
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on 'Yes' in popup ", "PASSED"
	End  If	
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
Public Function GFN_CreateCard(cardGroup, embossName, cardType, embossType)
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
