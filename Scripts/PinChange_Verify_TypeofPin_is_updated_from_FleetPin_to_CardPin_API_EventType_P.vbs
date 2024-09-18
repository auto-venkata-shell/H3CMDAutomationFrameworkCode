Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes"
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		inputfilepathval = Get_Dictionary(ParamValDict,"TempFileName" & "_" & iRowCount)
		etType = Get_Dictionary(ParamValDict,"EventType" & "_" & iRowCount)
		
		inputfilepathval = "pinchange.txt"
		outputfilepathval = "pinchangeoutput.txt"
		
		
		'*********** Initialize Global Data for JSON Request *****************************************
		customerERP_id = cardapi_data("custcardERP")
		
		reqType = "POST"
		apiurl = api_url & "/Card/PINChange"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CSSReg\PinChange\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CSSReg\PinChangeOutput\" & outputfilepathval
		newoutputxmlfile = sCurrentDirectory & "Test Data\CSSReg\PinChange\UpdateXML.xml" 
		newcreationxmlfile = sCurrentDirectory & "Test Data\CSSReg\PinChange\CreationInfoXML.xml" 
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
			
		newpindetails =  pinChagepremodifieddetailsinInputAPI(jsonFPath, etType)
		newpinver = Split(newpindetails,"|")(0)
		newpinPVV = Split(newpindetails,"|")(1)
		newpinPVK = Split(newpindetails,"|")(2)
		headers = getapiConfigInfoFromCSV(Apconfigpath)
		
		rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
		
		If rcheck = Empty Then
			Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
		Else
			
			Call validateAPIjob("2014",newoutputxmlfile,newcreationxmlfile)
			'XML_Cadrid = findValuefromString(newcreationxmlfile,"<Identity>","</Identity>")
			'XML_Cardpanid = findValuefromString(newcreationxmlfile,"<Identity2>","</Identity2>")
			Call validatepinChangePostAPIDetails(etType, newpinver, 2 ,newpinPVV, newpinPVK, 148, 3 )
						
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
			
		End If
	End If
next