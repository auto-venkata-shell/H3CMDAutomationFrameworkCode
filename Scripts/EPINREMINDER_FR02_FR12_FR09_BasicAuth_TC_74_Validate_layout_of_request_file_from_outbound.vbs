Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		set cardapi_data = CreateObject("Scripting.Dictionary")
		'cust_data.add "custERP", Get_Dictionary(ParamValDict,"Customer_ERP" & "_" & iRowCount)
		'msgbox customerERP_id
		'cardapi_data.add "custcardERP", customerERP_id
		
		cardapi_data.add "PinAdviceTypeidval", Get_Dictionary(ParamValDict,"PinAdviceTypeID" & "_" & iRowCount)
		cardapi_data.add "PinContactTypeIDval", Get_Dictionary(ParamValDict,"PinContactTypeID" & "_" & iRowCount)
	
		
		'*********** Generic Code for all scripts ****************'
		'Change Status from 0 to 1
		Call doSysVarColcoKeydataFlagAction("APILogging",0,1)  'Change Status from 0 to 1
		Call doAPIOutboundControlAction("APIAuthMethodID",2,1)
		Call doUseFleetPinAction("Resend",1)
		Call OpenApplication(url)
		Call navigateCustomerSummaryscreenforOtherScreens()
		Call setandresendFleetPinInCardParameterScreen("Resend")
		Call checkJobStatus("2015","SFN_SHELL_SPRINTQA_ID_BATCH",customerERP_id)
		
			JobRundate = Date()
			dtsnow = Now()
			'Individual date components
			dd = Right("00" & Day(dtsnow), 2)
			mm = Right("00" & Month(dtsnow), 2)
			yy = Year(dtsnow)
			hh = Right("00" & Hour(dtsnow), 2)
			nn = Right("00" & Minute(dtsnow), 2)
			ss = Right("00" & Second(dtsnow), 2)
			'Build the date string in the format yyyy-mm-dd
			customdatevalue = dd & "-"& mm & "-" & Mid(yy,3,2)
			
			filePath = appoutbPath & "\OutboundAPILog\" & customdatevalue
			filetype ="OutboundAPILog"

			fileoutboundlog = verifyFileExistance(filePath,filetype)
			cAddressDet = returnCustomerDetails()
			if fileoutboundlog <> "" and cAddressDet <> "" Then
				Call validateOutboundFileData(fileoutboundlog,cAddressDet,"txt")
			
			Else
					Append_TestHTML StepCounter, "OutboundAPI file data validation" , "file/Card not created" , "FAILED"
			End if
			
			fileSysObj.CreateTextFile(newfile)
			Set DXwrite = fileSysObj.OpenTextFile(newfile,2)
			DXwrite.Write rcheck
			DXwrite.Close
			Set fileSysObj = Nothing
	End If
next