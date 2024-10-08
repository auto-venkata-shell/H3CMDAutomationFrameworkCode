Function maintainExemptionscheme(typeOfExe,companyID,descrip,countryID,VATCategoryID,productVal)
	Call pageNavigation("Maintain Exemption Scheme","WebLink_MaintainExemptionScheme")
	wait 5
	If VerifyWebObjectExist("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "WebLink_NewExemptionScheme")  Then
		Append_TestHTML StepCounter, "Maintain Exemption Scheme", "User navigates to 'Maintain Exemption Scheme' screen", "PASSED"
		Call objectClick("WebLink_NewExemptionScheme","New Exemption Scheme link")
		Call enterWebList_value("WebList_TypeOfExemption","TypeOfExemption",typeOfExe)		
		Call enterWebList_value("WebList_CompanyID","Card Group",companyID)
		Call enterTextbox_value("WebEdit_textBoxDescription","Description",descrip)		
		Call enterWebList_value("WebList_CountryID","Card Group",countryID)
		Call enterWebList_value("WebList_VATCategoryID","Card Group",VATCategoryID)
		Call objectClick("WebLink_SearchProduct","Search Product link")
		wait 2
		If Browser(strBrowser).Page(strPage).WebTable("html id:=.*grdProductGroup").exist Then
			Call operateOnTblElement(".*grdProductGroup","Product Group Exemption" ,productVal)			
		End If
		ui_DateEff = OperateOnObject("WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebEdit_Maintain_EfffectiveDate", "GetROProperty", "value")
		Append_TestHTML StepCounter, "Fee Rule Tier", "Date Effective field is populated with '"&ui_DateEff&"' value", "PASSED"
		Call objectClick("WebLink_Save","Save link")
		wait 4
		
		Append_TestHTML StepCounter, "Maintain Exemption Scheme", "verify, created NewExemptionScheme is found in 'ExemptionSchemeItems' table", "PASSED"
		If Browser(strBrowser).Page(strPage).WebTable("html id:=.*gridViewExemptionSchemeItems").exist Then
			Set objLink = Browser(strBrowser).Page(strPage).WebElement("html id:=.*multiViewMain_pager")
			Set desc = Description.Create
			desc("micclass").value = "Link"
			Set childObj =  objLink.ChildObjects(desc)
			chdCount = childObj.count
			For i = 1 To chdCount-1 Step 1
				Call verifyTblValueExist(".*gridViewExemptionSchemeItems","Exemption Scheme Items" ,descrip)
				If flagValueprop  = False Then			
					childObj(chdCount-2).highlight					
					childObj(chdCount-2).Click
					Append_TestHTML StepCounter, "Maintain Exemption Scheme", "Click on 'Next page' link", "PASSED"
				else
					Exit For
				End If
			Next
			If flagValueprop = False Then
				Append_TestHTML StepCounter, "Maintain Exemption Scheme", "Expected table data '"&descrip&"' is not found in the 'Exemption Scheme Items' table", "FAILED"
			End If
		else
			Append_TestHTML StepCounter, "Maintain Exemption Scheme", "The Entered data is not saved sccessfully", "FAILED"	
		end if 			
	else
		Append_TestHTML StepCounter, "Maintain Exemption Scheme", "User navigates to 'Maintain Exemption Scheme' screen", "FAILED"
	End  If
End  Function

