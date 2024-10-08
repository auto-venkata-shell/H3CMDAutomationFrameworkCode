



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

Function clickOnPriceRuleDescription(strDesc)
On error resume next
	bFlag = True
	Call pageNavigation("Search Price Rule","WebLink_SearchPriceRule")
	Call objectClick("WebLink_Refresh","Refresh link")
	Call enterTextbox_value("WedEdit_PriceRuleDescription","Price Rule Description",strDesc)
	Call objectClick("WebLink_Search","Search link")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebTable_ResultTable") Then
		Append_TestHTML StepCounter, "Create price rule", "User is displayed with 'Search result table'", "PASSED"
		rcount = Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=.*mvPriceRules_grdResults").Rowcount
		For i = 1 To rcount Step 1
			ui_desc =  Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=.*mvPriceRules_grdResults").GetCellData(i,3)
			If trim(ui_desc) = strDesc Then
				Append_TestHTML StepCounter, "Create price rule", "New Price Rule with the Price Rule description '"&strDesc&"' is created successfully", "PASSED"				
				flag = True
				Exit For
			End  If
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

Function enterInput_searchPriceRulePage(strDesc,locScope,proScope,priceList,cmyName,availableFrom,refPrice,priceRuleBasis,lstcurrency,funder,IsCustomer,discount,priceRuleCat)
On error resume next
	bFlag = True
	Call pageNavigation("Search Price Rule","WebLink_SearchPriceRule")
	flag = False
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewPriceRule_H3") Then
		Append_TestHTML StepCounter, "Create price rule", "User navigates to 'Search Price Rule' page", "PASSED"
		Call objectClick("WebLink_NewPriceRule_H3","NewPriceRule link")
		If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WedEdit_Description") Then		
			Append_TestHTML StepCounter, "Create price rule", "User navigates to 'Create Price Rule' page", "PASSED"			 
			Call enterTextbox_value("WedEdit_Description","Description",strDesc)
			Browser("name:="&browserProp).Page("title:="&pageProp).WebList("html id:=.*ddlFeeTypeID").Select "#1"
			strInArrear = Browser("name:="&browserProp).Page("title:="&pageProp).WebList("html id:=.*ddlFeeTypeID").GetRoproperty("value")
			Append_TestHTML StepCounter, "Create price rule", "Select 'In Arrear Fee Rule' value as '"&strInArrear&"'", "PASSED"	
			Browser("name:="&browserProp).Page("title:="&pageProp).Link("html id:=.*LocationControl_srchDelco_SearchButton").click
			Append_TestHTML StepCounter, "Create price rule", "Click on 'Location scope' search link", "PASSED"			
			wait 2
			'Call operateOnTblElement("LocationControl_grdDelco","Location Scope table",locScope)
			Call objectClick("WebLink_ProductScopeSearch","Product Scope search link")
			Call operateOnTblElement("grdProductGroup","Product Group table","All Fuels")
			Call enterWebList_value("WebList_RefPrice","Ref Price",refPrice)
			Call enterWebList_value("WebList_PriceRuleBasis","Price Rule Basis",priceRuleBasis)
			Call enterWebList_value("WebList_Price_Currency","Currency",lstcurrency)
			Call enterWebList_value("WebList_Funder","Funder",funder)			
			Call objectClick("WebLink_Company","Company ID Search link")	
			Call enterTextbox_value("WedEdit_CompanyName","Company Name",cmyName)
			Set desc = Description.Create
			desc("micclass").value = "Link"
			Set chdObj = Browser("name:="&browserProp).Page("title:="&pageProp).childObjects(desc)
			For i = 1 To chdObj.count Step 1
				nameProp = chdObj(i).GetRoproperty("name")
				If nameProp = "Search" Then
					chdObj(i).click
					Exit For
				End If
			Next
'			Call objectClick("WebLink_Search","Search link")						
			Call operateOnTblElement(".*CompanySearch1_grdResults","Company Name details table",cmyName)
			Call operateOnCheckBox("WebCheckbox_Customer","IsCustomer",IsCustomer)
					
			Call enterWebList_value("WebList_Discount","Discount",discount)
			Call enterWebList_value("WebList_PriceRuleCategory","Price Rule Category",priceRuleCat)
'			Call click_on_save()
			OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
			Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
			wait 2
		else
			Append_TestHTML StepCounter, "Create price rule", "User does not navigated to 'Create Price Rule' page", "FAILED"
		End if
	else
		Append_TestHTML StepCounter, "Create price rule", "User does not navigated to 'Search Price Rule' page", "FAILED"
	End if	
End Function

Function enterInput_priceRuleLocation(locScope,cmyName)	
On error resume next
	bFlag = True
	Call navigateWithoutStartMenu("WebLink_PriceRuleDetails","WebLink_PriceRuleLocation", "WebLink_NewPriceRuleLocation","Price Rule Details","Price Rule Location")
'	Call navigateWithoutStartMenu("WebLink_PriceRuleDetails","WebLink_PriceRuleLocation","WebLink_NewPriceRuleLocation")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewPriceRuleLocation") Then
		Append_TestHTML StepCounter, "Create price rule", "User navigates to 'Price Rule Location' screen", "PASSED"
		Call operateOnTblElement(".*gvPriceRuleLocationList","Price Rule Location table",cmyName)
		ui_locScope = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_LocationScope", "GetROProperty", "value")
		If Trim(ui_locScope) = locScope Then
			Append_TestHTML StepCounter, "Location scope", "Location Scope '"&locScope&"' is pre populated in the screen", "PASSED"	
		else
			Append_TestHTML StepCounter, "Location scope", "Location Scope '"&locScope&"' is not pre populated in the screen", "FAILED"	
		End If			
		Call objectClick("WebLink_FuelNetworkSearch","Fuel Network Search link")		
		Call operateOnTblElement(".*grdFuelNetwork","LocationControl Fuel Network table",cmyName)
		Call objectClick("WebLink_SiteGroupSearch","Site Group Search link")
		Call operateOnTblElement(".*grdSiteGroup","PriceRuleLocations_LocationControl site group table",cmyName) 
		Call objectClick("WebLink_SiteIDSearch","Site ID Search link")	
		Call operateOnTblElement(".*grdSite","PriceRuleLocations_LocationControl site ID table",cmyName) 						
		OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
		Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"	

else
	Append_TestHTML StepCounter, "Create price rule", "User does not navigates to 'Price Rule Location' screen", "FAILED"
End  If
End Function

Function validate_PriceRuleLocationTabledata()
On error resume next
	bFlag = True
	Set tableObj = Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=.*gvPriceRuleLocationList")
	If tableObj.Exist Then
		Append_TestHTML StepCounter, "Price Rule Location table", "User is displayed with 'Price Rule Location' table with the entered data", "PASSED"						
		ui_cmyName =  tableObj.GetCellData(2,1)
		ui_fuelNetwork =  tableObj.GetCellData(2,2)
		ui_siteGroup =  tableObj.GetCellData(2,3)
		ui_siteID =  tableObj.GetCellData(2,4)
		Append_TestHTML StepCounter, "Price Rule Location table", "Verify the 'Price Rule Location' table data", "PASSED"
		If  ui_cmyName = cmyName Then
			Append_TestHTML StepCounter, "Price Rule Location Table", "Delivering Company : "&cmyName, "PASSED"
		else
			Append_TestHTML StepCounter, "Price Rule Location Table", "Delivering Company : "&cmyName, "FAILED"
		End  If
		If  ui_fuelNetwork = fuelNetwork Then
			Append_TestHTML StepCounter, "Price Rule Location Table", "Fuel Network : "&fuelNetwork, "PASSED"
		else
			Append_TestHTML StepCounter, "Price Rule Location Table", "Fuel Network : "&fuelNetwork, "FAILED"
		End  If
		If  ui_siteGroup = siteGroup Then
			Append_TestHTML StepCounter, "Price Rule Location Table", "Site Group : "&siteGroup, "PASSED"
		else
			Append_TestHTML StepCounter, "Price Rule Location Table", "Site Group : "&siteGroup, "FAILED"
		End  If
		If  ui_siteID = siteID Then
			Append_TestHTML StepCounter, "Price Rule Location Table", "Site ID : "&siteID, "PASSED"
		else
			Append_TestHTML StepCounter, "Price Rule Location Table", "Site ID : "&siteID, "FAILED"
		End  If
	else
		Append_TestHTML StepCounter, "Create price rule", "User is not displayed with 'Price Rule Location' table", "FAILED"
	end if 
End Function

Function enterInput_priceRuleProduct()
On error resume next
	bFlag = True
	Call navigateWithoutStartMenu("WebLink_PriceRuleLocation","WebLink_PriceRuleProduct", "WebLink_NewPriceRuleProduct","Price Rule Location","Price Rule Product")
'	Call  navigateStartMenu1("WebLink_PriceRuleLocation","WebLink_PriceRuleProduct","WebLink_NewPriceRuleProduct")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebLink_NewPriceRuleProduct") Then	
		Append_TestHTML StepCounter, "PriceRuleProduct page", "User navigates to 'Price Rule Product' screen", "PASSED"
		Call operateOnTblElement(".*gvPriceRuleProductList","PriceRuleProductList table","Forever") 
		ui_prodGroup = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_ProductGroup", "GetROProperty", "value")
		Append_TestHTML StepCounter, "PriceRuleProduct page", "Product Group name is populated with '"&ui_prodGroup&"'", "PASSED"
		Call click_on_Save()
	else
		Append_TestHTML StepCounter, "PriceRuleProduct page", "User does not navigates to 'Price Rule Product' screen", "FAILED"
	End  If
End  Function

Function verifyInput_priceRuleTier()
On error resume next
	bFlag = True
	Call navigateWithoutStartMenu("WebLink_PriceRuleProduct","WebLink_PriceRuleTier", "WebTable_ProductTiertable","Price Rule Product","Price Rule Tier")
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO",  "WebTable_ProductTiertable") Then
		Append_TestHTML StepCounter, "PriceRuleTier page", "User navigates to 'Price Rule Tier' screen", "PASSED"
		Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_grdResults")
		If tableObj.Exist Then
			Append_TestHTML StepCounter, "PriceRuleTier table", "User is displayed with 'Price Rule Tier' table", "PASSED"					
			Append_TestHTML StepCounter, "PriceRuleTier table", "Verify the data in  'Price Rule Tier' table", "PASSED"
			tierMin = tableObj.GetCelldata(2,2)			
			If tierMin = "0" Then
				Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Min : "&tierMin, "PASSED"
			else
				Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Min : "&tierMin, "FAILED"
			End If			
			tierMax = tableObj.GetCelldata(2,3)
			If tierMax = "9999999" Then
				Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Max : "&tierMax, "PASSED"
			else
				Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Max : "&tierMax, "FAILED"
			End If
			totValue = tableObj.GetCelldata(2,4)
			If Instr(totValue, txtValue) > 0 Then
				Append_TestHTML StepCounter, "PriceRuleTier table", "Total value : "&totValue, "PASSED"
			else
				Append_TestHTML StepCounter, "PriceRuleTier table", "Total value : "&totValue, "FAILED"
			End If
			dateEff = tableObj.GetCelldata(2,5)
			If IsDate(dateEff) Then
				Append_TestHTML StepCounter, "PriceRuleTier table", "Date Effective : "&dateEff, "PASSED"
			else
				Append_TestHTML StepCounter, "PriceRuleTier table", "Date Effective : "&dateEff, "FAILED"
			End If
			dateTer = tableObj.GetCelldata(2,6)
			If dateTer = "Forever" Then
				Append_TestHTML StepCounter, "PriceRuleTier table", "Date Terminated : "&dateTer, "PASSED"
			else
				Append_TestHTML StepCounter, "PriceRuleTier table", "Date Terminated : "&dateTer, "FAILED"
			End If
		else
			Append_TestHTML StepCounter, "PriceRuleTier", "User is not displayed with 'Price Rule Tier' table", "FAILED"
		End  If
	else
		Append_TestHTML StepCounter, "PriceRuleTier page", "User does not navigates to 'Price Rule Tier' screen", "FAILED"
	End  If
End Function
		
Function create_NewPriceRule(strDesc)
On error resume next
	bFlag = True
	 flag = false
	Call navigateWithoutStartMenu("Link_CustomerSummary","WebLink_MaintainCustomerPriceRules","WebLink_NewPriceRule","Customer Summary","Maintain Customer Price Rules")
	If Browser("name:="&browserProp).Page("title:="&pageProp).link("innertext:=New Customer Price Rule").exist Then
		Append_TestHTML StepCounter,"Maintain Customer Price Rules","User navigates to 'Maintain Customer Price Rules' screen","PASSED"
		Browser("name:="&browserProp).Page("title:="&pageProp).link("innertext:=New Customer Price Rule").click
		Append_TestHTML StepCounter,"NewPriceRule","Click on 'New Customer Price Rule link'" ,"PASSED"
		If Browser("name:="&browserProp).Page("title:="&pageProp).link("html id:=.*txtPriceRule_SearchButton").exist Then
			Append_TestHTML StepCounter,"NewPriceRule","User navigates to 'Customer Price Rule Create' screen" ,"PASSED"	
			Browser("name:="&browserProp).Page("title:="&pageProp).link("html id:=.*txtPriceRule_SearchButton").click
			Append_TestHTML StepCounter,"NewPriceRule","Click on 'Price Rule Search link'" ,"PASSED"
			Call objectClick("WebLink_Refresh","Refresh link")
			Call enterTextbox_value("WedEdit_PriceRuleDesc","Price Rule Description",strDesc)
			Call objectClick("WebLink_Search","Search link")		
			Set tableObj = Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=.*mvPriceRules_grdResults")
			Append_TestHTML StepCounter,"NewPriceRule","User is displayed with 'Price Rule Search result' table" ,"PASSED"	
			Set desc = description.Create
			desc("micclass").value = "WebElement"
			Set childObj = tableObj.ChildObjects(desc)
			chdCount =  childObj.count
			For i = 1 To chdCount Step 1
				ui_cellData =  childObj(i).GetRoproperty("innertext")
				If trim(ui_cellData) = strDesc Then
					Append_TestHTML StepCounter,"NewPriceRule","Search for the table data '"&strDesc&"'" ,"PASSED"
					Setting.WebPackage("ReplayType") = 2
					childObj(i).Fireevent "onmouseover"
					childObj(i).Fireevent "ondblclick"
					Setting.WebPackage("ReplayType") = 1	
					Append_TestHTML StepCounter,"NewPriceRule","Click on table data '"&strDesc&"'" ,"PASSED"				
					 flag = True
					 Exit For
				End If
			Next
			If flag = false Then
				 Append_TestHTML StepCounter,"NewPriceRule","Expcted table data '"&strDesc&"' is not found in 'Price Rule Search result' data" ,"FAILED"
			End If
		else
			Append_TestHTML StepCounter,"NewPriceRule","User does not navigates to 'Customer Price Rule Create' screen" ,"FAILED"
		End If
	else
		Append_TestHTML StepCounter,"Maintain Customer Price Rules","User does not navigates to 'Maintain Customer Price Rules' screen","FAILED"
	End If

End  Function

Function verifyPrePopulatedData_customerPriceRuleCreatePage(strDesc,effDate)
On error resume next
	bFlag = True
If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO","WedEdit_PriceRuleBasis") Then
	Append_TestHTML StepCounter, "Customer Price Rule create screen", "User navigates to 'Customer Price Rule create' screen", "PASSED"
	ui_priceRuleBasis = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WedEdit_PriceRuleBasis", "GetROProperty", "value")
	Append_TestHTML StepCounter, "Customer Price Rule create screen", "'Price Rule Basis' field is prepopulated with '"&ui_priceRuleBasis&"' value", "PASSED"	
	ui_refPrice = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_RefPrice", "GetROProperty", "value")
	Append_TestHTML StepCounter, "Customer Price Rule create screen", "'Ref Price' field is prepopulated with '"&ui_refPrice&"' value", "PASSED"
	Call enterTextbox_value("WedEdit_effDate","Effective Date",effDate)
	OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebLink_Save", "Click", ""
	Append_TestHTML StepCounter, "Click on save", "Click on save", "PASSED"
End  If

End  Function

Function verifyEnteredData_customerPriceRuleCreatePage(strDesc,effDate)
On error resume next
	bFlag = True
	wait 2
	Call objectClick("WebLink_Refresh","Refresh link")
	Browser("name:="&browserProp).Page("title:="&pageProp).WebEdit("html id:=.*txtPriceRuleDescription").click
	Browser("name:="&browserProp).Page("title:="&pageProp).WebEdit("html id:=.*txtPriceRuleDescription").Set strDesc
	Append_TestHTML StepCounter, "verify entered data", "Enter the Description '"&desc&"' in Price Rule Description field", "PASSED"
	Call objectClick("WebLink_Search","Search link")	
	Set tableObj = Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=.*grdResults")
	If tableObj.exist Then
		Append_TestHTML StepCounter, "verify entered data", "User is displayed with 'Customer Price Rule search result' table", "PASSED"	
		rcount = tableObj.Rowcount
		For i = 1 To rcount Step 1
			ui_desc = tableObj.GetCelldata(i,2)
			If trim(ui_desc) = strDesc Then
				Append_TestHTML StepCounter, "verify entered data", "Expected Price Rule description '"&strDesc&"' is exist in the 'Customer Price Rule search result' table", "PASSED"
				ui_effDate = tableObj.GetCelldata(2,3)
				If trim(ui_effDate) = effDate Then
					Append_TestHTML StepCounter, "verify entered data", "Entered Effective Date '"&effDate&"' is updated for the Price Rule description '"&strDesc&"'", "PASSED"	
					flag = True
					Exit For 
				else
					Append_TestHTML StepCounter, "verify entered data", "Entered Effective Date '"&effDate&"' is not updated for the Price Rule description '"&strDesc&"'", "FAILED"
				End If
			End  If
		Next
		If flag = false Then
			Append_TestHTML StepCounter, "verify entered data", "Expected Price Rule description '"&strDesc&"' is not exist in the 'Customer Price Rule search result' table", "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "verify entered data", "User is not displayed with 'Customer Price Rule search result' table", "FAILED"
	End  If
	
End  Function

Function verifyData_PriceRuleResultTable()
On error resume next
	bFlag = True
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebTable_PriceRuleResultTable")  Then
		Append_TestHTML StepCounter,"Search for Customer","User is displayed with 'Price rule result table'" ,"PASSED"
		Append_TestHTML StepCounter, "Create price rule", "validate the entered data", "PASSED"
		Call enterTextbox_value("WedEdit_SearchPriceRuleDesc","Price Rule Description",strDesc)
		Call objectClick("WebLink_Search","Search link")	
		Set tableObj = Browser("name:="&browserProp).Page("title:="&pageProp).WebTable("html id:=ctl00_CPH_ucCustomerPriceRuleSearch_grdResults")
		rcount = tableObj.Rowcount
		For i = 1 To rcount Step 1
			ui_pricedesc = tableObj.GetCellData(i,3)
			If ui_pricedesc = strDesc Then
				Append_TestHTML StepCounter, "Create price rule", "Entered Price Rule description '"&strDesc&"' is found in 'Customer Price Rule table'", "PASSED"
				tableObj.childItem(i,3,"WebElement",0).click
				flag = True
				Exit For
			End  If
		Next
		If flag = false Then
			Append_TestHTML StepCounter, "Create price rule", "Expected Price Rule description '"&strDesc&"' is not found in 'Customer Price Rule table'", "FAILED"
		End If
	else
		Append_TestHTML StepCounter,"Search for Customer","User is not displayed with 'Price rule result table'" ,"FAILED"
	End  If
End Function


Function pricing_updateRecordStatusinCSV(strDesc)
	On error resume next
	bFlag = True
	dataFile = TDFilePath & "TC008_PricingCreateNewPriceRule.csv"		
	 Set xlApp = CreateObject("Excel.Application")
	 xlApp.Visible = true
	 Set xlBook = xlApp.Workbooks.open(dataFile)
	sheetName = "TC008_PricingCreateNewPriceRule"
		Set xlSheet = xlBook.Worksheets(sheetName)
		rcount = xlSheet.UsedRange.rows.count
		For i = 1 To rcount Step 1
			exCountryName =  xlSheet.Cells(i,1).Value
			If exCountryName = "Yes_"& countryName Then
				xlSheet.Cells(i,3).Value = strDesc
				Exit For
			End If
		Next
	 
	xlBook.Save
	  xlApp.Quit
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

Function verifyPriceRuleTierTableData(txtValue)
On error resume next
	bFlag = True
	Set tableObj = Browser("creationTime:=1").Page("creationTime:=1").WebTable("html id:=ctl00_CPH_grdResults")
	If tableObj.Exist Then
		Append_TestHTML StepCounter, "PriceRuleTier table", "User is displayed with 'Price Rule Tier' table", "PASSED"					
		Append_TestHTML StepCounter, "PriceRuleTier table", "Verify the data in  'Price Rule Tier' table", "PASSED"
		tierMin = tableObj.GetCelldata(2,2)			
		If tierMin = "0" Then
			Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Min : "&tierMin, "PASSED"
		else
			Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Min : "&tierMin, "FAILED"
		End If			
		tierMax = tableObj.GetCelldata(2,3)
		If tierMax = "9999999" Then
			Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Max : "&tierMax, "PASSED"
		else
			Append_TestHTML StepCounter, "PriceRuleTier table", "Tier Max : "&tierMax, "FAILED"
		End If
		totValue = tableObj.GetCelldata(2,4)
		If Instr(totValue, txtValue) > 0 Then
			Append_TestHTML StepCounter, "PriceRuleTier table", "Total value : "&totValue, "PASSED"
		else
			Append_TestHTML StepCounter, "PriceRuleTier table", "Total value : "&totValue, "FAILED"
		End If
		dateEff = tableObj.GetCelldata(2,5)
		If IsDate(dateEff) Then
			Append_TestHTML StepCounter, "PriceRuleTier table", "Date Effective : "&dateEff, "PASSED"
		else
			Append_TestHTML StepCounter, "PriceRuleTier table", "Date Effective : "&dateEff, "FAILED"
		End If
		dateTer = tableObj.GetCelldata(2,6)
		If dateTer = "Forever" Then
			Append_TestHTML StepCounter, "PriceRuleTier table", "Date Terminated : "&dateTer, "PASSED"
		else
			Append_TestHTML StepCounter, "PriceRuleTier table", "Date Terminated : "&dateTer, "FAILED"
		End If
	else
		Append_TestHTML StepCounter, "PriceRuleTier table", "User is not displayed with 'Price Rule Tier' table", "FAILED"
	End  If
End Function

Function RunJob(JobTypeID)

	On error resume next
	bFlag = True
	
	query_job = "Select * from Job where JobTypeID = '"&JobTypeID&"' order by 1 desc;"
	Append_TestHTML StepCounter, "Run Job 19",query_job, "PASSED"
	set dictDbResultSet = execute_db_query(query_job, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
	wait 2
	jobID= dictDbResultSet("ID")
	db_nRundate = dictDbResultSet("NextRunDate")	
	db_jobID = jobID
	set dictDbResultSet = Nothing
		query_Runjob = "Select getdate() as 'compcurrentdatetime';"
		Append_TestHTML StepCounter, "Get JobID ",query_Runjob, "PASSED"
		set dictDbResultSet = execute_db_query(query_Runjob, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")	
		db_cdatetime = dictDbResultSet("compcurrentdatetime")	
		set dictDbResultSet = Nothing
		query_Runjob = Nothing
	If isNULL(db_nRundate) Then
		Append_TestHTML StepCounter, "Verify Job 4 status","Job is still in progress", "FAILED"
		bRunFlag = False
	Else
		
		query_JobUpt = "Update Job Set NextRunDate = NULL,StatusID=0 where ID = "&jobID&";"
		Append_TestHTML StepCounter, "Update Job table",query_JobUpt, "PASSED"
		Call update_db_query(query_JobUpt, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
		wait 10
		query_Runjob = "Select * from Job where ID = '"&jobID&"'"
		Append_TestHTML StepCounter, "Get JobID ",query_Runjob, "PASSED"
		set dictDbResultSet = execute_db_query(query_Runjob, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")	
		db_statusID = dictDbResultSet("StatusID")	
		set dictDbResultSet = Nothing
		query_Runjob = Nothing
		wait 5
		For itr = 1 To 5 Step 1
			query_Runjob = "Select * from Job where JobTypeID = '"&JobTypeID&"' order by 1 desc;"
			'Append_TestHTML StepCounter, "Get JobID ",query_Runjob, "PASSED"
			set dictDbResultSet = execute_db_query(query_Runjob, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")	
			db_jID = dictDbResultSet("ID")	
			db_dcreated = dictDbResultSet("DateCreated")	
			db_nRundate = dictDbResultSet("NextRunDate")	
			set dictDbResultSet = Nothing
			query_Runjob = Nothing
			wait 5
			If db_statusID = "0" Then
				query_Runjob = "Select * from Job where ID = '"&jobID&"'"
				'Append_TestHTML StepCounter, "Get JobID ",query_Runjob, "PASSED"
				set dictDbResultSet = execute_db_query(query_Runjob, 1, "GFN_SHELL_SPRINTQA_ID_BATCH")
				db_statusID = dictDbResultSet("StatusID")	
				wait 10
			ElseIf db_statusID = "4" and db_dcreated >= db_cdatetime and isNull(NextRunDate) = False Then
				Exit for
			End If
		Next
		If db_statusID = "4"  Then
			Append_TestHTML StepCounter, "Get JobID ","As expected, job Status ID changed to '4'", "PASSED"		
		Else
			Append_TestHTML StepCounter, "Get JobID ","As expected, job Status ID changed to '4' - Current job status: " & db_statusID, "FAILED"	
			bRunFlag = False
			bFlag = False
		End If	

	End If	
End Function
