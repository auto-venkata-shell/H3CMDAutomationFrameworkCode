Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	
	If strExecute = "Yes" Then
	
		strUI = Get_Dictionary(ParamValDict,"Username" & "_" & iRowCount)
		strPassword = Get_Dictionary(ParamValDict,"Password" & "_" & iRowCount)		
		strOrgID = Get_Dictionary(ParamValDict,"OrgID" & "_" & iRowCount)	
		strRole = Get_Dictionary(ParamValDict,"Role" & "_" & iRowCount)		
		
		strItemName = Get_Dictionary(ParamValDict,"ItemName" & "_" & iRowCount)		
		strBU = Get_Dictionary(ParamValDict,"TransferFromBU" & "_" & iRowCount)	
		strQty = Get_Dictionary(ParamValDict,"Quantity" & "_" & iRowCount)
		'strSecListRetail = Get_Dictionary(ParamValDict,"SecListRetail" & "_" & iRowCount)
					
		'strNoOfQualifiers = Get_Dictionary(ParamValDict,"NoOfQualifiers" & "_" & iRowCount)
		'strStartDate = Get_Dictionary(ParamValDict,"StartDate" & "_" & iRowCount)
		'strEndDate = Get_Dictionary(ParamValDict,"EndDate" & "_" & iRowCount)
		
		'strMinAmt = Get_Dictionary(ParamValDict,"MinAmt" & "_" & iRowCount)
		'strDiscountType = Get_Dictionary(ParamValDict,"DiscountType" & "_" & iRowCount)
		'strCurrencyValue = Get_Dictionary(ParamValDict,"CurrencyValue" & "_" & iRowCount)
		'strPercentValue = Get_Dictionary(ParamValDict,"PercentValue" & "_" & iRowCount)
		
		'*********** Generic Code for all scripts ****************
		'Select Environment from test data file
		strEnv = Environment.Value("ENV_Flag")
		ENV_Flag = strEnv
		Call SelectEnvirnoment(ENV_Flag)
		'*********** Generic Code for all scripts ****************
		
		'********************** Scenario specific variables and business functions *********************	
		
		Call BOS_Login_Rp(strUI,strPassword,strOrgID,strRole)	
		
		Call TransferInventoryIn(strItemName,strBU,strQty,strRole,strOrgID)
		
		Call BOS_Logout()
		 
		'Call CreateTestDataOutputFile(iRowCount)
		
		'********************** Scenario specific variables and business functions *********************
	
	End If
next