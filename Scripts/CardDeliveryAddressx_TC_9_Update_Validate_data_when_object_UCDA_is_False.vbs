Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		
		inputfilepathval = Get_Dictionary(ParamValDict,"inputfilepath" & "_" & iRowCount)
		outputfilepathval = Get_Dictionary(ParamValDict,"outputfilepath" & "_" & iRowCount)
		
		'********************** Scenario specific variables and business functions *********************
		reqType = "POST"
		apiurl = api_url & "/Card/CardDeliveryAddress_Update"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\jsonapi\" & inputfilepathval 
		headers = "Content-Type==>application/json,charset==>utf-8,ClientID==>" & client_ID & ",ClientSecretKey==>" & client_SecretKey
		newfile = sCurrentDirectory & "Test Data\jsonapi\" & outputfilepathval
		jsonFPath = strJson
		
		'*********** Generic Code for all scripts ****************
		'Getting error
		searchString = "Card_ContactName"
		replaceString = "Venkata3"
		strFlag= True
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)	
		searchString = "Card_RegionID"
		replaceString = "2"
		strFlag= True
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)
		searchString = "Card_CountryID"
		replaceString = "85"
		strFlag= True
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)
		searchString = "UserName"
		replaceString = "Venkata3"
		strFlag= True
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)	
		searchString = "ColCoID"
		replaceString = "85"
		strFlag= False
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)
		searchString = "CustomerID"
		replaceString = "2"
		strFlag= False
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)
		searchString = "CardID"
		replaceString = "1224"
		strFlag= False
		Call searchandReplaceString(jsonFPath, searchString,replaceString,strFlag)
		Call verifyPreCardActiveDetails("1224")
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
			Call validatejob2013()
			Call verifyPostCardActiveDetails("1224")
			Set fileSysObj = createObject("Scripting.FileSystemObject")
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
		End If

	
	End If
next