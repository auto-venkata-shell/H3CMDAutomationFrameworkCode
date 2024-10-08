'Public customerERP_id, cardPAN_no, JobRundate, cardExpiry_date, dx026FinalFName, Trans_BatchID, settllReportDate, dx350_Summary_Docno, dx350_Balance, dx350_Paymentduedate,registrationnumber,customername,custBankNamedesc

Public  db_FeeTypeID, db_billingQuan, db_jobID , db_CustomerID,ui_feeRuleDesc
'db_FeeRuleID -- already exist

Function objectClick(objlink,objName)
On error resume next
	bFlag = True
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", objlink, "Click", ""
	Append_TestHTML StepCounter, "Click on object", "Click on '"&objName&"'", "PASSED"
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

Function validate_resultTable(strDesc)
On error resume next
	bFlag = True
	Call pageNavigation("Search Price Rule","WebLink_SearchPriceRule")
	Call objectClick("WebLink_Refresh","Refresh link")
	Call enterTextbox_value("WedEdit_PriceRuleDescription","Price Rule Description",strDesc)
	Call objectClick("WebLink_Search","Search link")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebTable_ResultTable") Then
		Append_TestHTML StepCounter, "Create price rule", "User is displayed with 'Search result table'", "PASSED"
		Set tableObj = Browser("name:=Pilipinas Shell Petroleum Corp").Page("title:=Pilipinas Shell Petroleum Corp").WebTable("html id:=.*mvPriceRules_grdResults")
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

Function fees_navigateMaintainFeeRules(feeType,strDesc,freq,feeproduct,productGroup,feeBasis,volInclRate,waiveOnPastDue,waiveIfCancelled,minValue,maxValue,availableFrom,strCurrency)
On error resume next
	bFlag = True
	Call navigateWithStartMenu("WebLink_Pricing_Fees","WebLink_SearchForFeeRules","WebLink_NewFeeRule","Pricing Fees","Search For Fee Rules")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_NewFeeRule")  Then
		Call objectClick("WebLink_NewFeeRule","New Fee Rule link")
		wait 2
		If feeType = "Card Transaction Fee" Then
			feeType = "Card Transaction Fee – Billing Delco"
		End If
		Browser("creationTime:=1").Page("creationTime:=1").WebList("html id:=.*ddlFeeTypeID").Select feeType
		Append_TestHTML StepCounter, "Select value", "Select the 'Fee Type' value as '"&feeType&"'", "PASSED"
		Call enterTextbox_value("WedEdit_FeeDescription","Description",strDesc)
		Call enterWebList_value("WebList_Frequency","Frequency",freq)
		If productGroup <> empty Then
			Call objectClick("WebLink_ProductSearchButton","Product Group Search link")
			Call enterTextbox_value("Fees_WebEdit_PGroup","Productgroup",productGroup)
			Call enterTextbox_value("Fees_WebEdit_Product","Product",feeproduct)
			Call objectClick("WebLink_Search","Product Group Search")
			Call operateOnTblElement(".*grdProduct","Product Group" ,feeproduct)	
		End If			
		Call enterWebList_value("WebList_FeeBasis","Fee Basis",feeBasis)
		If volInclRate <> "" Then
			Call enterWebList_value("WebList_VolumeInclusionRate","Volume Inclusion Rate",volInclRate)	
		End If
		
		Call operateOnCheckBox("WebCheckbox_WaiveonPastDue","Waive on Past Due",waiveOnPastDue)
		Call operateOnCheckBox("WebCheckbox_WaiveIfCancelled","Waive If Cancelled",waiveIfCancelled)
		Call enterTextbox_value("WedEdit_MinValue","Min Value",minValue)
		Call enterTextbox_value("WedEdit_MaxValue","Max Value",maxValue)
		Call enterTextbox_value("WedEdit_AvailableFrom","Available From",availableFrom)
		Call enterWebList_value("WebList_Currency","Currency",strCurrency)
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
		Call validate_savePopup()
		wait 2		
	else
		Append_TestHTML StepCounter,"Search for Customer","User does not navigates to 'Maintain Fee Rules' screen" ,"FAILED"
	End  If
End Function

Function fees_navigateFeeRuleLocation(cmyName,fuelNetwork,SiteGroup,SiteID)
On error resume next
	bFlag = True
	bFlag = navigateWithoutStartMenu("WebLink_MaintainFeeRules","WebLink_FeeRuleLocation","WebLink_NewFeeRuleLocation","Maintain Fee Rules","Fee Rule Location")
'	CAll pageNavigation("Fee Rule Location","WebLink_FeeRuleLocation")
	If  bFlag = True Then
		Call objectClick("WebLink_NewFeeRuleLocation","New Fee Rule Location link")	
		wait 2
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_LocationScopeSearchButton")  Then
			Append_TestHTML StepCounter, "Fee Rule Location", "User navigates to 'New Fee Rule Location' screen", "PASSED"
			ui_strDateEff = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_DateEffective", "GetROProperty", "value")		
			If Not(ui_strDateEff) = ""  Then
				Append_TestHTML StepCounter, "Date Effective", "Date Effective field is prepopulated with '"&ui_strDateEff&"' value", "PASSED"
			else
				Append_TestHTML StepCounter, "Date Effective", "Date Effective field is not prepopulated with '"&ui_strDateEff&"' value", "FAILED"
			End If
			If cmyName <> empty Then
				Call objectClick("Fees_Link_Delcosearch","Delco Location Scope Search Button")
				Call enterTextbox_value("Fees_WebEdit_CompanyName","Product",cmyName)
				Call objectClick("Fees_Link_Comapnysearch","Company Search")
				'Call operateOnTblElement(".*grdDelco","Company Name" ,cmyName)	
			End If 				
			If fuelNetwork <> empty Then
				Call objectClick("WebLink_FuelNetworkSearchButton","Fuel Network Search Button")
				Call operateOnTblElement(".*grdFuelNetwork","Fuel Network" ,fuelNetwork)
			End If
			'ui_feeLocScope = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeLocationScope", "GetROProperty", "value")				
			'Append_TestHTML StepCounter, "Location Scope", "'Location Scope' text box is populated with '"&ui_feeLocScope&"' Company Name", "PASSED"
			'cmyName = ui_feeLocScope			
			If SiteGroup <> empty Then
				Call objectClick("WebLink_SiteGroupSearchButton","Site Group Search Button")
				Call operateOnTblElement(".*grdSiteGroup","Site Group" ,SiteGroup)
			End If	
			If SiteID <> empty Then
				Call objectClick("WebLink_SiteIDSearchButton","Site ID Search Button")
				Call operateOnTblElement(".*grdSite","Site ID" ,SiteID)
			End If					
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			wait 2
			Call validate_savePopup()
		else
			Append_TestHTML StepCounter,"Fee Rule Location","User is not displayed with 'New Fee Rule Location' link" ,"FAILED"
		End  If
	End  If
End Function

Function fees_navigateFeeRuleProduct(prodGroup,product)
On error resume next
	bFlag = True
	flag = false
'	bFlag = navigateWithoutStartMenu("WebLink_FeeRuleLocation","WebLink_FeeRuleProduct","WebLink_NewFeeRuleProduct","Fee Rule Location","Fee Rule Product")
	Call pageNavigation("Fee Rule Product","WebLink_FeeRuleProduct")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_NewFeeRuleProduct")  Then
		Call objectClick("WebLink_NewFeeRuleProduct","New Fee Rule Product link")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_ProductGroupSearchButton")  Then		
			wait 2
			ui_existingProdtDetails = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_ExistingProductDetails", "GetROProperty", "innertext")		
			If Not(ui_existingProdtDetails) = ""  Then
				Append_TestHTML StepCounter, "Existing Product Details", "User is displayed with existing Product Details '"&ui_existingProdtDetails&"'", "PASSED"
			else
				Append_TestHTML StepCounter, "Existing Product Details", "User is not displayed with existing Product Details '"&ui_existingProdtDetails&"'", "FAILED"
			End If
			ui_strDateEff = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_DateEffective", "GetROProperty", "value")		
			If Not(ui_strDateEff) = ""  Then
				Append_TestHTML StepCounter, "Date Effective", "User is prepopulated with Date Effective '"&ui_strDateEff&"'", "PASSED"
			else
				Append_TestHTML StepCounter, "Date Effective", "User is not prepopulated with Date Effective '"&ui_strDateEff&"'", "FAILED"
			End If
			If prodGroup <> empty Then
				Call objectClick("WebLink_ProductGroupSearchButton","Product Group Search Button")
				Call operateOnTblElement(".*grdProductGroup","Product Group" ,prodGroup)
			End If	
			If product <> empty Then
				Call objectClick("WebLink_ProductSearchButton","Product Search Button")
				Call operateOnTblElement(".*grdProduct","Product" ,product)	
			End If	
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			wait 2
			Call validate_savePopup()
		else
			Append_TestHTML StepCounter, "Fee Rule Product", "User is not displayed with 'Fee Rule Product' link", "FAILED"
		End  IF
	End  If		
End Function

Function fees_navigateFeeRuleTier_maintainFeeRule()
On error resume next
	bFlag = True
	flag = false
'	bFlag = navigateWithoutStartMenu("WebLink_FeeRuleTier","WebLink_MaintainFeeRules","WedEdit_FeeRuleDescription","Fee Rule Tier","Maintain Fee Rules")
	Call pageNavigation("Maintain Fee Rules","WebLink_MaintainFeeRules")
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_FeeRuleDescription")  Then
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_FeeRuleDescription")  Then
		ui_feeRuleDesc = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeRuleDescription", "GetROProperty", "value")		
		If Not(ui_feeRuleDesc) = ""  Then
			Append_TestHTML StepCounter, "Fee Rule Details", "As Expected, Fee Rule Description field is populated with '"&ui_feeRuleDesc&"' description", "PASSED"
		else
			Append_TestHTML StepCounter, "Fee Rule Details", "Fee Rule Description field is not populated with '"&ui_feeRuleDesc&"' description", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "Fee Rule Details", "User navigates to 'Fee Rule Details' screen",  "FAILED"
	End  IF	
End Function

Function validate_feeRuleProductTableData(prodGroup,product,strDateEff)
On error resume next
	bFlag = True
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*gvFeeRuleProductList")
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Fee Rule Product", "Verify entered data is updated in 'Fee Rule Product' table", "PASSED"
		rcount = tableObj.Rowcount
		For i = 1 To rcount Step 1
			ui_prodGroup = tableObj.GetCellData(i,1)
			ui_product = tableObj.GetCellData(i,2)
			ui_dateEff = tableObj.GetCellData(i,3)
			ui_dateTer = tableObj.GetCellData(i,4)
			If (ui_prodGroup = prodGroup) and ui_product = product and ui_dateEff = strDateEff Then
				Append_TestHTML StepCounter, "Fee Rule Product", "Product Group : '"&prodGroup&"' is updated in 'Fee Rule Product' table", "PASSED"
				Append_TestHTML StepCounter, "Fee Rule Product", "Product : '"&product&"' is updated in 'Fee Rule Product' table", "PASSED"
				Append_TestHTML StepCounter, "Fee Rule Product", "Date Effective : '"&strDateEff&"' is updated in 'Fee Rule Product' table", "PASSED"
				Flag = True
				Exit For
			End If
		Next
		If Flag = False Then
			Append_TestHTML StepCounter, "Fee Rule Product", "Entered data is not updated in 'Fee Rule Product' table", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "Fee Rule Product", "'Fee Rule Product' table is not found in the screen", "FAILED"
	End If
End Function

Function validate_feeRuleLocationTableData(cmyName,fuelNetwork,SiteGroup,SiteID,strDateEff)
On error resume next
	bFlag = True
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*gvFeeRuleLocationList")
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Fee Rule Location", "Verify entered data is updated in 'Fee Rule Location' table", "PASSED"
		rcount = tableObj.Rowcount
		For i = 1 To rcount Step 1
			ui_delCompany = tableObj.GetCellData(i,1)
			ui_fuelNetwork = tableObj.GetCellData(i,2)
			ui_siteGroup = tableObj.GetCellData(i,3)
			ui_siteID = tableObj.GetCellData(i,4)
			ui_dateEff = tableObj.GetCellData(i,5)
			If  trim (ui_delCompany) =   trim(cmyName) and trim(ui_fuelNetwork) = trim(fuelNetwork) and trim(ui_siteGroup) = trim(siteGroup) and trim(ui_siteID) = trim(siteID) and trim(ui_dateEff) = trim(strDateEff) Then
				Append_TestHTML StepCounter, "Fee Rule Location", "Delivering Company  : '"&delCompany&"' is updated in 'Fee Rule Location' table", "PASSED"
				Append_TestHTML StepCounter, "Fee Rule Location", "Fuel Network : '"&fuelNetwork&"' is updated in 'Fee Rule Location' table", "PASSED"
				Append_TestHTML StepCounter, "Fee Rule Location", "Site Group : '"&siteGroup&"' is updated in 'Fee Rule Location' table", "PASSED"
				Append_TestHTML StepCounter, "Fee Rule Location", "Site ID : '"&siteID&"' is updated in 'Fee Rule Location' table", "PASSED"
				Append_TestHTML StepCounter, "Fee Rule Location", "Date Effective : '"&strDateEff&"' is updated in 'Fee Rule Location' table", "PASSED"
				Flag = True
				Exit For
			End If
		Next
		If Flag = False Then
			Append_TestHTML StepCounter, "Fee Rule Location", "Entered data is not updated in 'Fee Rule Location' table", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "Fee Rule Location", "'Fee Rule Location' table is not found in the screen", "FAILED"
	End If
End Function

Function validate_feeRuleTableData(strDesc)
On error resume next
	bFlag = True
	flag = false
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*ctl00_CPH_SearchControl_grdResults")	
	Call enterTextbox_value("WedEdit_FeeRuleDescp","Fee Rule Description",strDesc)
	Call objectClick("WebLink_Search","Search link")
	wait 2
	If tableObj.exist Then
		Append_TestHTML StepCounter, "Fee Rule table", "Verify entered data is updated in 'Fee Rule' table", "PASSED"
		rcount = tableObj.Rowcount
		For i = 1 To rcount Step 1
			ui_strDesc = tableObj.GetCellData(i,4)
			If trim(ui_strDesc) = strDesc Then
				Append_TestHTML StepCounter, "Fee Rule table", "Entered New Fee Rule data with Fee Rule description '"&strDesc&"' is updated in 'Fee Rule' table", "PASSED"
				'To click on table element
				Set desc = description.Create
				desc("micclass").value = "WebElement"
				Set childObj = tableObj.ChildObjects(desc)
				chdCount =  childObj.count
				For j = 1 To chdCount Step 1
					ui_strDesc = childObj(j).GetRoproperty("innertext")
					If trim(ui_strDesc) = strDesc Then
					 	childObj(j).click
					 	Append_TestHTML StepCounter, "Fee Rule table", "Click on Fee Rule description '"&strDesc&"' in 'Fee Rule' table", "PASSED"
					 	flag = True
					 	wait 2
					 	Exit For
					 End  If
				Next
				flag = True
				Exit For
			End If
		Next
		If Flag = False Then
			Append_TestHTML StepCounter, "Fee Rule table", "Entered New Fee Rule data with Fee Rule description '"&strDesc&"' is not updated in 'Fee Rule' table", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "Fee Rule table", "'Fee Rule' table is not found in the screen", "FAILED"
	End If
End Function

Function validate_feeRuleTierTableData(strDateEff,feeValue)	
On error resume next
	bFlag = True
	Append_TestHTML StepCounter, "Fee Rule Tier", "Verify data in 'Fee Rule Tier' table", "PASSED"
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*ctl00_CPH_gvFeeRuleTierList")
	rcount = tableObj.Rowcount
	If rcount > 1 Then
		ui_tierMin = tableObj.GetCellData(2,2)
		ui_tierMax = tableObj.GetCellData(2,3)
		ui_totalValue = tableObj.GetCellData(2,4)
		ui_dateEff = tableObj.GetCellData(2,5)
		ui_dateTer = tableObj.GetCellData(2,6)
		If ui_tierMin = "1" Then
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Tier Min : '"&ui_tierMin&"'", "PASSED"
		else
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Tier Min : '"&ui_tierMin&"'", "FAILED"
		End IF 
		If ui_tierMax = "999999999" Then
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Tier Max : '"&ui_tierMax&"'", "PASSED"
		else
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Tier Max : '"&ui_tierMax&"'", "FAILED"
		End IF 
		If Instr(ui_totalValue, feeValue) > 0 Then
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Total Value : '"&feeValue&"'", "PASSED"
		else
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Total Value : '"&feeValue&"'", "FAILED"
		End IF 
		If ui_dateEff = strDateEff Then
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Date Effective : '"&ui_dateEff&"'", "PASSED"
		else
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Date Effective : '"&ui_dateEff&"'", "FAILED"
		End IF 
		If ui_dateTer = "Forever" Then
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Date Terminated : '"&ui_dateTer&"'", "PASSED"
		else
			Append_TestHTML StepCounter, "Fee Rule Tier table", "Date Terminated : '"&ui_dateTer&"'", "FAILED"
		End IF
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
		wait 2
'		Call validate_savePopup()
	else
		Append_TestHTML StepCounter, "Fee Rule Tier table", "No data displayed for the user", "FAILED"
	End If		
End Function


Function navigate_feeRuleTierScreen(feeValue)
On error resume next
	bFlag = True
	flag = false
'	Call navigateWithoutStartMenu("WebLink_FeeRuleProduct","WebLink_FeeRuleTier","WebLink_NewFeeRuleTier","Fee Rule Product","Fee Rule Tier")
	Call pageNavigation("Fee Rule Tier","WebLink_FeeRuleTier")
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*ctl00_CPH_gvFeeRuleTierList")
	If tableObj.exist Then
		flag = True
		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Fee Rule Tier' screen", "PASSED"
		Append_TestHTML StepCounter, "Fee Rule Tier table", "User is displayed with 'Fee Rule Tier' table", "PASSED"
		Set desc = description.Create
		desc("micclass").value = "WebElement"
		Set chdObj = tableObj.ChildObjects(desc)
		For i = 1 To chdObj.count Step 1
			If chdObj(i).GetRoproperty("innertext") = "1" Then
				chdObj(i).click
				wait 2
				Append_TestHTML StepCounter, "Fee Rule Tier table", "Click on 'Fee Rule Tier' table cell data", "PASSED"
				Exit For				
			End If
		Next
'	else
'		Append_TestHTML StepCounter, "Fee Rule Tier", "'Fee Rule Tier' table is not found in the screen", "FAILED"
	End If
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_FeeRuleValue")  Then
		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Fee Rule Tier' screen", "PASSED"
		Call enterTextbox_value("WedEdit_FeeRuleValue","Value",feeValue)
		Call objectClick("WebLink_Save","Save link")
		flag = True
		wait 5
	End  If
	If flag = false Then
		Append_TestHTML StepCounter, "Fee Rule Tier", "User does not navigates to 'Fee Rule Tier' screen", "FAILED"
	End If
End  Function


'Function navigate_feeRuleTierScreen(feeValue)
''	Call navigateWithoutStartMenu("WebLink_FeeRuleProduct","WebLink_FeeRuleTier","WebLink_NewFeeRuleTier","Fee Rule Product","Fee Rule Tier")
'	Call pageNavigation("Fee Rule Tier","WebLink_FeeRuleTier")
'	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_FeeRuleTierDesc")  Then
'		Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Fee Rule Tier' screen", "PASSED"
'		Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*ctl00_CPH_gvFeeRuleTierList")
'		If tableObj.exist Then
'			Append_TestHTML StepCounter, "Fee Rule Tier", "User navigates to 'Fee Rule Tier' screen", "PASSED"
'			Append_TestHTML StepCounter, "Fee Rule Tier table", "User is displayed with 'Fee Rule Tier' table", "PASSED"
'			Set desc = description.Create
'			desc("micclass").value = "WebElement"
'			Set chdObj = tableObj.ChildObjects(desc)
'			For i = 1 To chdObj.count Step 1
'				If chdObj(i).GetRoproperty("innertext") = "1" Then
'					chdObj(i).click
'					wait 2
'					Append_TestHTML StepCounter, "Fee Rule Tier table", "Click on 'Fee Rule Tier' table cell data", "PASSED"
'					Exit For				
'				End If
'			Next
'			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_FeeRuleBasis")  Then
'				Append_TestHTML StepCounter, "Fee Rule Tier", "User is displayed with 'Fee Rule Tier' screen with enabled 'Fee Value' text box", "PASSED"
'				Call enterTextbox_value("WedEdit_FeeRuleBasis","Fee Value",feeValue)
'				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
'				Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
'				wait 2
'			End  If	
'		else
'			Append_TestHTML StepCounter, "Fee Rule Tier", "'Fee Rule Tier' table is not found in the screen", "FAILED"
'		End If
'	else
'		Append_TestHTML StepCounter, "Fee Rule Tier", "User does not navigates to 'Fee Rule Tier' screen", "FAILED"
'	End  If
'End  Function


Function addNewFeeRule(feeType,waiveOnPastDue,waiveIfCancelled)
On error resume next
	bFlag = True
	Call navigateWithoutStartMenu("Link_CustomerSummary","WebLink_SearchFeeRules","WebLink_NewFeeRule","Customer Summary","Search For Fee Rules")
	'Call pageNavigation("Search For Fee Rules","WebLink_SearchFeeRules")
	Call objectClick("WebLink_NewFeeRule","New Fee Rule link")
	wait 2
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_SearchNewFeeRule")  Then
		Append_TestHTML StepCounter, "Page Navigation", "User navigates to 'Customer Fee Rule Details' screen", "PASSED"
		Call objectClick("WebLink_SearchNewFeeRule","Fee Rule Search link")
		wait 2
'		Browser("creationTime:=1").Page("creationTime:=1").WebList("html id:=.*ddlFeeType").Select feeType
'		Append_TestHTML StepCounter, "Select value", "Select the 'Fee Type' value as '"&feeType&"'", "PASSED"
		Call enterTextbox_value("WebEdit_FeeRuleDesc","Fee Rule Description",strDesc)
		Append_TestHTML StepCounter, "Enter value", "Enter the 'Fee Rule Description '"&strDesc&"'", "PASSED"
		Call operateOnCheckBox("WebCheckBox_FeeWaivePastDue","Waive on Past Due",waiveOnPastDue)
		Call operateOnCheckBox("WebCheckBox_FeeWaivePastCancelled","Waive If Cancelled",waiveIfCancelled)
		Call objectClick("WebLink_Search","Search link")
		wait 5
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_CustFeeRuleList")  Then
'			Append_TestHTML StepCounter, "Page Navigation", "User displayed with Customer Fee Rule table' with the search criteria", "PASSED"		
'			strDesc = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*grdResults").Getcelldata(2,3)
		
			Call operateOnTblElement(".*ucFeeRuleSearchControl_grdResults","Customer Fee Rule" ,strDesc)
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_DateEffective")  Then	
				Append_TestHTML StepCounter, "Page Navigation", "User is displayed with selected Fee Rule Description '"&strDesc&"' in 'Customer Fee Rule Details' screen", "PASSED"
'				ui_feeLocationScope = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeLocationScope", "GetROProperty", "value")
'				Append_TestHTML StepCounter, "validate data", "'Fee Type' field is prepopulated with '"&ui_feeLocationScope&"' value", "PASSED"
				'ui_feeProduct = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeProduct", "GetROProperty", "value")
				'Append_TestHTML StepCounter, "validate data", "'Fee Product' field is prepopulated with '"&ui_feeProduct&"' value", "PASSED"
				ui_feeRuleBasis = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeRuleBasis", "GetROProperty", "value")
				Append_TestHTML StepCounter, "validate data", "'Fee Rule Basis' field is prepopulated with '"&ui_feeRuleBasis&"' value", "PASSED"
				ui_dateEff = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_DateEffective", "GetROProperty", "value")
				Append_TestHTML StepCounter, "validate data", "'Date Effective' field is prepopulated with '"&ui_dateEff&"' value", "PASSED"'
				If  Browser("creationTime:=1").Page("creationTime:=1").WebList("html id:=.*ddlBonusPayToID").exist Then
					Call enterWebList_value("WebList_BonusPayToID","BonusPayToID","Pay to Payer")
				End If								
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
				Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
				wait 2
'				Call validate_savePopup()
			else
				Append_TestHTML StepCounter, "Page Navigation", "User does not navigates to 'Customer Fee Rule Details' screen", "FAILED"
			End  If
		else
			Append_TestHTML StepCounter, "Page Navigation", "User is not displayed with 'Customer Fee Rules' table", "FAILED"
		End  If
	else
		Append_TestHTML StepCounter, "Click on save", "User does not navigates to 'Maintain Customer Fee Rules' screen", "FAILED"
	End  If
End Function

Function bonus_addNewFeeRule(feeType,waiveOnPastDue,waiveIfCancelled,bonusID)
On error resume next
	bFlag = True
	Call navigateWithoutStartMenu("Link_CustomerSummary","WebLink_SearchFeeRules","WebLink_NewFeeRule","Customer Summary","Search For Fee Rules")
	'Call pageNavigation("Search For Fee Rules","WebLink_SearchFeeRules")
	Call objectClick("WebLink_NewFeeRule","New Fee Rule link")
	wait 2
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_SearchNewFeeRule")  Then
		Append_TestHTML StepCounter, "Page Navigation", "User navigates to 'Customer Fee Rule Details' screen", "PASSED"
		Call objectClick("WebLink_SearchNewFeeRule","Fee Rule Search link")
		wait 2
'		Browser("creationTime:=1").Page("creationTime:=1").WebList("html id:=.*ddlFeeType").Select feeType
'		Append_TestHTML StepCounter, "Select value", "Select the 'Fee Type' value as '"&feeType&"'", "PASSED"
		Call enterTextbox_value("WebEdit_FeeRuleDesc","Fee Rule Description",strDesc)
		Append_TestHTML StepCounter, "Enter value", "Enter the 'Fee Rule Description '"&strDesc&"'", "PASSED"
		Call operateOnCheckBox("WebCheckBox_FeeWaivePastDue","Waive on Past Due",waiveOnPastDue)
		Call operateOnCheckBox("WebCheckBox_FeeWaivePastCancelled","Waive If Cancelled",waiveIfCancelled)
		Call objectClick("WebLink_Search","Search link")
		wait 5
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_CustFeeRuleList")  Then
'			Append_TestHTML StepCounter, "Page Navigation", "User displayed with Customer Fee Rule table' with the search criteria", "PASSED"		
			strDesc = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=.*grdResults").Getcelldata(2,3)
		
			Call operateOnTblElement(".*ucFeeRuleSearchControl_grdResults","Customer Fee Rule" ,strDesc)
			If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WedEdit_DateEffective")  Then	
				Append_TestHTML StepCounter, "Page Navigation", "User is displayed with selected Fee Rule Description '"&strDesc&"' in 'Customer Fee Rule Details' screen", "PASSED"
'				ui_feeLocationScope = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeLocationScope", "GetROProperty", "value")
'				Append_TestHTML StepCounter, "validate data", "'Fee Type' field is prepopulated with '"&ui_feeLocationScope&"' value", "PASSED"
				'ui_feeProduct = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeProduct", "GetROProperty", "value")
				'Append_TestHTML StepCounter, "validate data", "'Fee Product' field is prepopulated with '"&ui_feeProduct&"' value", "PASSED"
				ui_feeRuleBasis = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_FeeRuleBasis", "GetROProperty", "value")
				Append_TestHTML StepCounter, "validate data", "'Fee Rule Basis' field is prepopulated with '"&ui_feeRuleBasis&"' value", "PASSED"
				ui_dateEff = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_DateEffective", "GetROProperty", "value")
				Append_TestHTML StepCounter, "validate data", "'Date Effective' field is prepopulated with '"&ui_dateEff&"' value", "PASSED"
				If isempty(bonusID)=False Then
					Call enterWebList_value("WebList_BonusPayToID","BonusPayToID",bonusID)
				End If
				
								
				OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
				Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
				wait 5
				query = "select * from CustomerFeeRule where CustomerID = (select customerid from Customer where CustomerERP = '" & customerERP_id & "') order by 1 desc;"
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
				
				If instr(trim(feeRuleDesc),trim( strDesc))<>0 Then
					Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"PASSED"
				else
					Append_TestHTML StepCounter,"Fee Rule Validation","Expected Value: " & feeRuleDesc & VBCRLF & "Actual Value: " & feeRuleDescGFN ,"FAILED"
					bFlag = False
				End If
				set dbRecordSet2 = Nothing
				set dbRecordSet = Nothing
				
				
				query_FeeType = "Select * from FeeType where FeeTypeID = '"&db_FeeTypeID&"'"
				Append_TestHTML StepCounter, "FeeType table",query_FeeType, "PASSED"
				set dictDbResultSet = execute_db_query(query_FeeType, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
				wait 2
				db_FeeTypeID1 = dictDbResultSet("FeeTypeID")
				db_isBilling = dictDbResultSet("isBilling")
				
				If db_isBilling = 1 Then
					Append_TestHTML StepCounter, "DB_query","As expected, 'isBilling' value is 1", "PASSED"
				End If
'				Call validate_savePopup()
			else
				Append_TestHTML StepCounter, "Page Navigation", "User does not navigates to 'Customer Fee Rule Details' screen", "FAILED"
			End  If
		else
			Append_TestHTML StepCounter, "Page Navigation", "User is not displayed with 'Customer Fee Rules' table", "FAILED"
		End  If
	else
		Append_TestHTML StepCounter, "Click on save", "User does not navigates to 'Maintain Customer Fee Rules' screen", "FAILED"
	End  If
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

Function fees_navigateBillingAcceptance()
On error resume next
	bFlag = True
	Call navigateWithStartMenu("WebLink_Transactions","WebLink_BillAcceptance","WebButton_BillPreview","Transactions","Billing Acceptance")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_Preview")  Then
		Call objectClick("WebButton_Preview","Preview button")
		wait 5
		Do until VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebButton_Refresh") = False
			wait 15
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Refresh", "Click", ""
			wait 15
			Append_TestHTML StepCounter, "Refresh", "Clicked on refresh", "PASSED"
		Loop
'		Call objectClick("WebButton_Refresh","Refresh button")		
		wait 20
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Preview") Then
			previewMsg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Preview", "GetROProperty", "innertext")
			Append_TestHTML StepCounter, "verify Message", "User is displayed with '"&previewMsg&"' message", "PASSED"
		End If
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_Txns1") Then
			Append_TestHTML StepCounter, "verify table", "User is displayed with 'Billed Transactions Details' table with entered Fee Rule details", "PASSED"
		End If
		Call objectClick("WebElement_BilledFees","BilledFees tab")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_BilledFees") Then
			Append_TestHTML StepCounter, "verify table", "User is displayed with 'Billed Fees' table with entered Fee Rule details", "PASSED"
		End If
		
	else
		Append_TestHTML StepCounter,"Search for Customer","User does not navigates to 'Billing Acceptance' screen" ,"FAILED"
	End  If
End  Function
'		Call objectClick("WebButton_BillPreview","Billing Preview button")
'		wait 5
'		Call objectClick("WebButton_Refresh","Refresh button")	
'		wait 30
'		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Preview") Then
'			previewMsg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Preview", "GetROProperty", "innertext")
'			Append_TestHTML StepCounter, "verify Message", "User is displayed with '"&previewMsg&"' message", "PASSED"
'		End If

Function click_on_cutoff()
On error resume next
	bFlag = True
	Call objectClick("WebButton_Cutoff","Cut off")
	wait 5
	Call objectClick("WebButton_Refresh","Refresh button")	
	wait 30
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebElement_Cutoff") Then
		cutOffMsg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Cutoff", "GetROProperty", "innertext")
		Append_TestHTML StepCounter, "verify Message", "User is displayed with '"&cutOffMsg&"' message", "PASSED"
	End If
End Function

Function click_on_signoff()
On error resume next
	bFlag = True
	Call objectClick("WebButton_Signoff","Sign Off")
	wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebDialog_popup") Then
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebDialog_popup", "Highlight", ""
		signOffMsg = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebDialog_popup", "GetROProperty", "innertext")
		Append_TestHTML StepCounter, "verify Message", "User is displayed with '"&signOffMsg&"' message", "PASSED"
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_OK", "Click", ""
		Append_TestHTML StepCounter, "Click object", "Click on 'OK' button in the popup", "PASSED"
	End If
	Call objectClick("WebButton_Refresh","Refresh button")	
	wait 30
	
End Function

Function fees_DBqueries(strDesc,strDateEff)
	On error resume next
	bFlag = True
	query_feeRuleId = "Select * from FeeRule where FeeRuleDescription = '"&strDesc&"'"
	Append_TestHTML StepCounter, "Get 'Fee Rule ID' from 'Fee Rule' table",query_feeRuleId, "PASSED"
	set dictDbResultSet = execute_db_query(query_feeRuleId, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_FeeRuleID= dictDbResultSet("FeeRuleID")
	db_FeeTypeID = dictDbResultSet("FeeTypeID")
	Set dictDbResultSet = Nothing
	query_feeRule = "Update FeeRule Set IsCustomer = 1 , AvailableFrom = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"'" 
	Append_TestHTML StepCounter, "Update FeeRule table",query_feeRule, "PASSED"
	Call  update_db_query(query_feeRule, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2

	query_feeRuleLoc = "Update FeeRuleLocation Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
	Append_TestHTML StepCounter, "Update FeeRuleLocation table",query_feeRuleLoc, "PASSED"
	Call update_db_query(query_feeRuleLoc, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2

	query_feeRuleProdt = "Update FeeRuleProduct Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
	Append_TestHTML StepCounter, "Update FeeRuleProduct table",query_feeRuleProdt, "PASSED"
	Call update_db_query(query_feeRuleProdt, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	query_feeRuleTier = "Update FeeRuleTier Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
	Append_TestHTML StepCounter, "Update FeeRuleTier table",query_feeRuleTier, "PASSED"
	Call update_db_query(query_feeRuleTier, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
		
	query_salesItemunbilled = "Select * from SalesItemUnbilled where ColcoID = DelcoID order by BatchID desc"	
	Append_TestHTML StepCounter, "Get' SalesItemID' from 'SalesItemUnbilled' table",query_salesItemunbilled, "PASSED"
	set dictDbResultSet = execute_db_query(query_salesItemunbilled, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	db_salesItemID = dictDbResultSet("SalesItemID")	
	db_CustomerID = dictDbResultSet("CustomerID")
	wait 2
	
	query_customer = "Select * from Customer where CustomerID = '"&	db_CustomerID&"'"
	Append_TestHTML StepCounter, "GetCustomer ERP",query_customer, "PASSED"
	set dictDbResultSet = execute_db_query(query_customer, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")	
	db_CustomerERP = dictDbResultSet("CustomerERP")
	wait 2
	fees_DBqueries = db_CustomerERP
	
End  Function

Function afterCustomerUpdate(strDateEff,db_CustomerERP,FeeTypeID)
On error resume next
	bFlag = True
	query_custFeeRule = "Update CustomerFeeRule Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
	Append_TestHTML StepCounter, "Update CustomerFeeRule table",query_custFeeRule, "PASSED"
	Call update_db_query(query_custFeeRule, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	
	query_FeeType = "Select * from FeeType where FeeTypeID = '"&FeeTypeID&"'"
	Append_TestHTML StepCounter, "FeeType table",query_FeeType, "PASSED"
	set dictDbResultSet = execute_db_query(query_FeeType, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_FeeTypeID = dictDbResultSet("FeeTypeID")
	db_isBilling = dictDbResultSet("isBilling")
	
	If db_isBilling = 1 Then
		Append_TestHTML StepCounter, "DB_query","As expected, 'isBilling' value is 1", "PASSED"
	End If
	
	query_billingReportDate = "Select * from BillingAcceptance order by 1 desc"	
	Append_TestHTML StepCounter, "Get' Billing Report Date' from 'Billing Acceptance' table",query_billingReportDate, "PASSED"
	set dictDbResultSet = execute_db_query(query_billingReportDate, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	db_billReportDate = dictDbResultSet("BillingReportDate")
	wait 2
	
	next_date_dbval =  nextdateDBFormt(db_billReportDate)
	prev_month_date_dbval =  previousmonthDBFormt(db_billReportDate)
	crr_date_dbval =  nextdateDBFormt(date()-1)
'	query_salesItemUnbill = "Select * from SalesItemUnbilled where BillingReportDate = '"&next_date_dbval&"'"	
'	Append_TestHTML StepCounter, "Get' Billing Report Date' from 'SalesItemUnbilled' table",query_salesItemUnbill, "PASSED"
'	set dictDbResultSet = execute_db_query(query_salesItemUnbill, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
'	db_billReportDate = dictDbResultSet("BillingReportDate")
'	If db_billReportDate = Empty Then
'		Append_TestHTML StepCounter, "DB_query","As expected, there is no data found in 'SalesItemUnbilled' table for entered Billing Report date", "PASSED"
'	End If
	wait 2	
'	query_FeeItemUnbill = "Select * from FeeItemUnbilled where BillingReportDate = '"&next_date_dbval&"'"	
'	Append_TestHTML StepCounter, "Get' Billing Report Date' from 'FeeItemUnbilled' table",query_FeeItemUnbill, "PASSED"
'	set dictDbResultSet = execute_db_query(query_FeeItemUnbill, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
'	db_billReportDate = dictDbResultSet("BillingReportDate")
'	If db_billReportDate = Empty Then
'		Append_TestHTML StepCounter, "DB_query","As expected, there is no data found in 'FeeItemUnbilled' table for entered Billing Report date", "PASSED"
'	End If		

	query_updateSalesItemUnbilled = "Update SalesItemUnbilled Set BillingReportDate = '"&next_date_dbval&"' where SalesItemID = '"&db_salesItemID&"'"
	Append_TestHTML StepCounter, "Update SalesItemUnbilled table",query_updateSalesItemUnbilled, "PASSED"
	Call update_db_query(query_updateSalesItemUnbilled, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	dayAftertom_date_dbval  =  nextdateDBFormt(next_date_dbval)
	query_updateInfoPro = "Update InfoProviderNextRunDate Set NextReportDate = '"&next_date_dbval&"', NextRunDate = '"&dayAftertom_date_dbval&"' where InfoProviderID = 4"
	Append_TestHTML StepCounter, "Update InfoProviderNextRunDate table",query_updateInfoPro, "PASSED"
	Call update_db_query(query_updateInfoPro, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
'	query_FeeNextCreationDate = "Select * from FeeNextCreationDate where FeeRuleID = '"&db_FeeRuleID&"'"
'	Append_TestHTML StepCounter, "FeeNextCreationDate table",query_FeeNextCreationDate, "PASSED"
'	set dictDbResultSet = execute_db_query(query_FeeNextCreationDate, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
'	wait 2
'	Append_TestHTML StepCounter, "FeeNextCreationDate table","As expected, no data found in the 'FeeNextCreationDate' table ", "PASSED"
'			
	Call RunJob("4")	
	query_FeeNextcrtDate = "Select * from FeeNextCreationDate where FeeRuleID = '"&db_FeeRuleID&"'"	
	Append_TestHTML StepCounter, "Get' Billing Report Date' from 'SalesItemUnbilled' table",query_FeeNextcrtDate, "PASSED"
	set dictDbResultSet = execute_db_query(query_FeeNextcrtDate, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	db_FeeRuleID = dictDbResultSet("FeeRuleID")
	If db_FeeRuleID <> empty Then
		Append_TestHTML StepCounter, "FeeNextCreationDate table","As expected, FeeRuleID '"&db_FeeRuleID&"' is found in FeeNextCreationDate table" , "PASSED"
	End If
			
	query_update_FeeNextCreate = "Update FeeNextCreationDate Set NextFeeCreationDate = '"&next_date_dbval&"' ,  PreviousFeeCreationDate = '"&prev_month_date_dbval&"' where FeeRuleID  = '"&db_FeeRuleID&"'"
	Append_TestHTML StepCounter, "Update FeeNextCreationDate table",query_update_FeeNextCreate, "PASSED"
	Call update_db_query(query_update_FeeNextCreate, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	
	query_update_FeeBillRpt = "Update FeeNextBillingReportDate Set NextBillingReportDate = '"&next_date_dbval&"' ,  PreviousBillingReportDate = '"&prev_month_date_dbval&"' where FeeTypeID  = '"&db_FeeTypeID&"'"
	Append_TestHTML StepCounter, "Update FeeNextBillingReportDate table",query_update_FeeBillRpt, "PASSED"
	Call update_db_query(query_update_FeeBillRpt, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
'	query_FeeItemUnbilled = "Select * from FeeItemUnbilled where FeeTypeID = '"&db_FeeTypeID&"'"
'	Append_TestHTML StepCounter, "FeeItemUnbilled table",query_FeeItemUnbilled, "PASSED"
'	Call update_db_query(query_FeeItemUnbilled, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
'	wait 2
'	Append_TestHTML StepCounter, "FeeItemUnbilled table","As expected, no data found in the 'FeeItemUnbilled' table ", "PASSED"
	afterCustomerUpdate = db_FeeRuleID
End Function

Function afterBillingPreview(db_FeeRuleID)
On error resume next
	bFlag = True
	query_FeeItemUnbilled = "Select * from FeeItemUnbilled where FeeRuleID = '"&db_FeeRuleID&"'"
	Append_TestHTML StepCounter, "FeeItemUnbilled table",query_FeeItemUnbilled, "PASSED"
	Set dictDbResultSet = execute_db_query(query_FeeItemUnbilled, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_FeeRuleID = dictDbResultSet("FeeRuleID")
	Append_TestHTML StepCounter, "FeeItemUnbilled table","As expected, FeeRule ID '"&db_FeeRuleID&"' generated in 'FeeItemUnbilled' table ", "PASSED"
	afterBillingPreview = db_BillingDocumentID
End Function

Function afterCutoff()
On error resume next
	bFlag = True
	Call RunJob("19")
	Call RunJob("64")		
End Function

Function getBillingDocumentID(db_FeeRuleID)
On error resume next
	bFlag = True
	query_FeeItem = "Select * from FeeItem where FeeRuleID = '"&db_FeeRuleID&"'"
	Append_TestHTML StepCounter, "FeeItem table",query_FeeItem, "PASSED"
	Set dictDbResultSet = execute_db_query(query_FeeItem, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_FeeBillingDocumentID = dictDbResultSet("BillingDocumentID")
	query_salesItem = "Select  * from SalesItem where SalesItemID = '"&db_salesItemID&"'"
	Append_TestHTML StepCounter, "FeeItem table",query_salesItem, "PASSED"
	Set dictDbResultSet = execute_db_query(query_salesItem, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	db_salesBillingDocumentID = dictDbResultSet("BillingDocumentID")
	If db_FeeBillingDocumentID = db_salesBillingDocumentID Then
		Append_TestHTML StepCounter, "BillingDocumentID","Fee Item BillingDocumentID '"&db_FeeBillingDocumentID&"' matches with Sales Item BillingDocumentID '"&db_salesBillingDocumentID&"" , "PASSED"
	else
		Append_TestHTML StepCounter, "BillingDocumentID","Fee Item BillingDocumentID '"&db_FeeBillingDocumentID&"' does not matches with Sales Item BillingDocumentID '"&db_salesBillingDocumentID&"" , "FAILED"
	End If
End Function
			'Need to attach the fee rule with the customer
	
	'after Preview and Refresh FeeItemUnbilled table data is generate and Billed fees tab oda table also generate
	'then Click on Cut off and Sign off 
	'run job 19 and 64
	'Select * from FeeItem where FeeType = 42 - Get the Billing document ID
	'Select * from SalesItem where BillingReportDate = '' 	
	'Compare the BillingDocumentID of both table FeeItem and SalesItem																																																																																																										

Function CreateManualFee()
On error resume next
	bFlag = True
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
End Function




Function dobackdatesForFeeRuleid(strDesc,strDateEff)
	On error resume next
	bFlag = True
	
	query_feeRuleId = "Select * from FeeRule where FeeRuleDescription = '"&strDesc&"'"
	Append_TestHTML StepCounter, "Get 'Fee Rule ID' from 'Fee Rule' table",query_feeRuleId, "PASSED"
	set dictDbResultSet = execute_db_query(query_feeRuleId, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
	wait 2
	db_FeeRuleID= dictDbResultSet("FeeRuleID")
	db_FeeTypeID = dictDbResultSet("FeeTypeID")
	Set dictDbResultSet = Nothing
	If db_FeeRuleID <> empty Then
		
		query_feeRule = "Update FeeRule Set IsCustomer = 1 , AvailableFrom = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"'" 
		Append_TestHTML StepCounter, "Update FeeRule table",query_feeRule, "PASSED"
		Call  update_db_query(query_feeRule, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
	
		query_feeRuleLoc = "Update FeeRuleLocation Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
		Append_TestHTML StepCounter, "Update FeeRuleLocation table",query_feeRuleLoc, "PASSED"
		Call update_db_query(query_feeRuleLoc, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
	
		query_feeRuleProdt = "Update FeeRuleProduct Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
		Append_TestHTML StepCounter, "Update FeeRuleProduct table",query_feeRuleProdt, "PASSED"
		Call update_db_query(query_feeRuleProdt, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		query_feeRuleTier = "Update FeeRuleTier Set DateEffective = '"&strDateEff&"' where FeeRuleID = '"&db_FeeRuleID&"' "
		Append_TestHTML StepCounter, "Update FeeRuleTier table",query_feeRuleTier, "PASSED"
		Call update_db_query(query_feeRuleTier, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
		wait 2
		
	Else
		Append_TestHTML StepCounter,"Do back date of DateEffective FeeRule- "& strDesc ,"No Fee rule id exist in the FeeRule DB table" ,"FAILED"	
		bRunFlag = False
	End If
End Function


Function getFeeNextCreationDateValues()
	On error resume next
	bFlag = True
	
	query_feeRuleId = "Select * from FeeNextCreationDate where Feeruleid="&db_FeeRuleID&";"
	Append_TestHTML StepCounter, "Get 'Fee Rule ID' from 'Fee Rule' table",query_feeRuleId, "PASSED"
	set dictDbResultSet = execute_db_query(query_feeRuleId, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
	wait 2
	db_FeeNcreationDate= dictDbResultSet("NextFeeCreationDate")
	db_FeePcreationDate = dictDbResultSet("PreviousFeeCreationDate")
	Set dictDbResultSet = Nothing
	If db_FeeNcreationDate <> ""  and db_FeePcreationDate <> "" Then
		Append_TestHTML StepCounter,"Verify record entry for "& db_FeeRuleID ,"Entry updated after job-4" ,"PASSED"	
	
	Else
		Append_TestHTML StepCounter,"Verify record entry for "& db_FeeRuleID ,"Fail to add an entry into the table after job-4" ,"FAILED"		
		bRunFlag = False
	End If
	
End Function



Function getsalesItemBonusEntries()
	On error resume next
	bFlag = True
	query_feeRuleId = "Select * from Customer where CustomerERP='"&customerERP_id&"';"
	Append_TestHTML StepCounter, "Get CustomerID from Customer table",query_feeRuleId, "PASSED"
	set dictDbResultSet = execute_db_query(query_feeRuleId, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
	wait 2
	db_CustID= dictDbResultSet("CustomerID")
	db_CustomerID = db_CustID
	Set dictDbResultSet = Nothing
	query_feeRuleId = Nothing
	query = "Select Count(*) as 'NoofRows' from SalesItemBonus where CustomerID='"&db_CustID&"';"
	Append_TestHTML StepCounter, "Get row count from SalesItemBonus table using customerid",query, "PASSED"
	set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
	wait 2
	db_rcount= dictDbResultSet("NoofRows")
	Set dictDbResultSet = Nothing
	query = Nothing
	If (db_rcount <> Empty  or db_rcount > 0) and  db_FeeRuleID <> "" and db_CustID <> "" Then
		query = "Select Year(min(BillingReportDate)) as 'MinreportDateYear' ,month(min(BillingReportDate)) as 'MinreportDateMonth' ,min(BillingReportDate) as 'MinReportDate' from SalesItemBonus  group by CustomerID,ProductGroupID,DelcoID having CustomerID='"&db_CustID&"';"
		Append_TestHTML StepCounter, "Get earliest date from SalesItemBonus table",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_minYear= dictDbResultSet("MinreportDateYear")
		db_minMonth= dictDbResultSet("MinreportDateMonth")
		db_oldestDate= dictDbResultSet("MinReportDate")
		Set dictDbResultSet = Nothing
		query = Nothing
		query = "Select cast(getdate()-1 as date) as 'BeforeDate';"
		Append_TestHTML StepCounter, "Get Currentdate",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_nfcenddate= dictDbResultSet("BeforeDate")
		Set dictDbResultSet = Nothing
		query = Nothing
		query = "Select DATEFROMPARTS(Year('"& db_oldestDate &"'),month('"& db_oldestDate &"'),1) as 'Pfeestartdate';"
		Append_TestHTML StepCounter, "Get Currentdate",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_pcsdate= dictDbResultSet("Pfeestartdate")
		db_Feeitemfilterdate = db_pcsdate
		Set dictDbResultSet = Nothing
		query = Nothing
		query = "Select Sum(Quantity) as 'BillingQuantity',Sum(ColcoInvoiceValueTotalGross) as 'BillingGross', Sum(ColcoInvoiceValueTotalNet) as 'BillingNet',ProductGroupID,DelcoID from SalesItemBonus  group by CustomerID,ProductGroupID,DelcoID having CustomerID='"&db_CustID&"';"
		Append_TestHTML StepCounter, "Get earliest date from SalesItemBonus table",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_billingQuan= dictDbResultSet("BillingQuantity")
		db_billingGross= dictDbResultSet("BillingGross")
		db_billingNet= dictDbResultSet("BillingNet")
		Set dictDbResultSet = Nothing
		query = Nothing
		query_feeRuleTier = "Update FeeNextCreationDate Set NextFeeCreationDate='"& db_nfcenddate & "', PreviousFeeCreationDate='"& db_pcsdate & "' where FeeRuleID='" & db_FeeRuleID &"';"
			Append_TestHTML StepCounter, "Update FeeRuleTier table",query_feeRuleTier, "PASSED"
			Call update_db_query(query_feeRuleTier, 1, "GFN_SHELL_SPRINTQA_ID_OLTP")
			wait 2
		query_feeRuleTier = Nothing	
	getsalesItemBonusEntries = True
	Else
		Append_TestHTML StepCounter, "Validate entries in SalesItemBonu Table","No Entries are found is table using customer id"&db_CustID , "FAILED"
		bRunFlag = False
		getsalesItemBonusEntries = False
	End If
		
End Function

Function validateFeeUnbilledEntries()
	On error resume next
	bFlag = True
	query = "Select Count(*) as 'Numofentries' from JobLog where JobID ='"& db_jobID &"' order by 1 desc;"
		Append_TestHTML StepCounter, "Get Currentdate",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_rowentries= dictDbResultSet("Numofentries")
		Set dictDbResultSet = Nothing
		query = Nothing
		query = "Select * from JobLog where JobID ='"& db_jobID &"' order by 1 desc;"
		Append_TestHTML StepCounter, "Get Currentdate",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, cint(db_rowentries), "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_Message= dictDbResultSet("Message")
		date_messagess = split(db_Messages,"|")
		Set dictDbResultSet = Nothing
		query = Nothing
		If instr(db_Message,usp_Billing_BonusFee_Insert) > 0 Then
				Append_TestHTML StepCounter, "Validate Job log entries","Picked bonus fee successfully and valdiated" & db_Message  , "PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Job log entries","Fail to picked bonus fee " & db_Message  , "FAILED"
				bRunFlag = False
			
		End If
		query = "Select CustomerInvoiceAmountGross,CustomerInvoiceAmountNet,VATCategoryID,VATPercentage,ColcoFeeAmountVat,ColcoFeeAmountNet,ColcoFeeAmountGross,FeeDateTime,SystemEntryDateTime,BillingReportDate,FeeItemID from FeeItemUnbilled where FeeTypeID ="& db_FeeTypeID & " and CustomerID="& db_CustomerID & " and FeeDateTime= '"& db_Feeitemfilterdate &"';"
		Append_TestHTML StepCounter, "Get Currentdate",query, "PASSED"
		set dictDbResultSet = execute_db_query(query, 1, "GFN_SHELL_SPRINTQA_NL_BATCH")
		wait 2
		db_ciAN= dictDbResultSet("CustomerInvoiceAmountNet")
		db_ciAG= dictDbResultSet("CustomerInvoiceAmountGross")
		db_vatCI= dictDbResultSet("VATCategoryID")
		db_vatPer= dictDbResultSet("VATPercentage")
		db_colcoFAV= dictDbResultSet("ColcoFeeAmountVat")
		db_colcoFAN= dictDbResultSet("ColcoFeeAmountNet")
		db_feeitemID= dictDbResultSet("FeeItemID")
		Set dictDbResultSet = Nothing
		query = Nothing
		If isNull(db_feeitemID) = False Then
				db_feeciAN = cdbl(db_vatCI) * cdbl(db_vatPer)
			db_feeciAG = cdbl(db_colcoFAV) + cdbl(db_colcoFAN)
			If cdbl(db_billingQuan) = cdbl(db_feeubQuan) Then
				Append_TestHTML StepCounter, "Validate Quantites-","Expected Value: " & db_billingQuan & VBCRLF & "Actual Value: " & db_feeubQuan ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Quantites-","Expected Value: " & db_billingQuan & VBCRLF & "Actual Value: " & db_feeubQuan ,"FAILED"
			
			End If
			If cdbl(db_ciAG) = cdbl(db_feeciAG) Then
				Append_TestHTML StepCounter, "Validate Quantites-","Expected Value: " & db_ciAG & VBCRLF & "Actual Value: " & db_feeciAG ,"PASSED"
			Else
				Append_TestHTML StepCounter, "Validate Quantites-","Expected Value: " & db_ciAG & VBCRLF & "Actual Value: " & db_feeciAG ,"FAILED"
			
			End If
		Else
			Append_TestHTML StepCounter, "Verify Feeitemunbilled Entries","Fail to identify entries in unbilled table", "FAILED"
			bRunFlag = False		
		End If
	
End Function



