Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	'strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase("Yes") Then
		
		set err_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		err_data.add "envval", Get_Dictionary(ParamValDict,"Environment" & "_" & iRowCount)
		err_data.add "userName", Get_Dictionary(ParamValDict,"Username" & "_" & iRowCount)
		err_data.add "page", Get_Dictionary(ParamValDict,"pageno" & "_" & iRowCount)
		err_data.add "dpageSize", Get_Dictionary(ParamValDict,"dpagesize" & "_" & iRowCount)
		err_data.add "startDate", Get_Dictionary(ParamValDict,"sdate" & "_" & iRowCount)
		err_data.add "endDate", Get_Dictionary(ParamValDict,"endate" & "_" & iRowCount)

		inputfilepathval = "ErrorAPI.txt"
		outputfilepathval = "ErrorAPIoutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		'Below flag for checking default price profile and fee values
		defaultcheckflag = True
		
		reqType = "POST"
		apiurl = api_url & "/Customer/MasterDataError"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************
	'	Append_TestHTML StepCounter, "Display Page ", "Page filed is: " & err_data("page") , "PASSED"
	'	Append_TestHTML StepCounter, "Display defalut PageSize ", "PageSize is: " & err_data("dpageSize") , "PASSED"
		
		Call getErrorapiprerequisits(err_data,jsonFPath)
		newPagcount = errorapi_Totalnumrecs/err_data("dpageSize")
		'msgbox newPagcount
		
		err_data.add "pageSize", newPagcount
		
		
		If errorapi_Totalnumrecs > 0 Then
				Append_TestHTML StepCounter, "Get Error Records", errorapi_Totalnumrecs & " entries are exist in the choosen start and end dates"  ,"PASSED"	
				If instr((newPagcount),".")>0 Then
					newPagcount = (newPagcount) + 0.5
	
				else
					newPagcount = newPagcount
				End If
				For itr = 1 To newPagcount Step 1
						
						searchreplaceString = "Page-" & itr & ";PageSize-" & err_data("dpageSize")
						Call searchandReplaceMultipleStrings_in_JSON(jsonFPath, searchreplaceString)
						headers = getapiConfigInfoFromCSV(Apconfigpath)
						rcheck = invokeAPI(reqType, apiurl, asynctype, strType, strJson, headers,stropJson)
						
						If rcheck = Empty Then
							Append_TestHTML StepCounter, "Verify Response JSON ", "Error:" & transErrDesc , "FAILED"
						Else
							'Call verifyPosterrorAPIDetails(err_data,rcheck)
							Call verifyPagecountValidations(rcheck,newPagcount)
							Call errorAPIResponseDetailsValidation(jsonFPath,rcheck)
							Set fileSysObj = createObject("Scripting.FileSystemObject")
							fileSysObj.CreateTextFile(newfile)
							Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
							DXwrite.Write rcheck
							DXwrite.Close
							Set fileSysObj = Nothing
						End If
				Next
		Else
				Append_TestHTML StepCounter, "Get Error Records", "Fail to get error entitypID entries from MessageQueue Error table"  ,"FAILED"	

		End If
		
	End If
next