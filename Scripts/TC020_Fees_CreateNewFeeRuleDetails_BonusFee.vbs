Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 

	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	feeType = Get_Dictionary(ParamValDict,"feeType" & "_" & iRowCount)
	freq = Get_Dictionary(ParamValDict,"freq" & "_" & iRowCount)
	productGroup = Get_Dictionary(ParamValDict,"productGroup" & "_" & iRowCount)
	feeproduct = Get_Dictionary(ParamValDict,"product" & "_" & iRowCount)
	feeBasis = Get_Dictionary(ParamValDict,"feeBasis" & "_" & iRowCount)
	volInclRate = Get_Dictionary(ParamValDict,"volInclRate" & "_" & iRowCount)
	minValue = Get_Dictionary(ParamValDict,"minValue" & "_" & iRowCount)
	maxValue = Get_Dictionary(ParamValDict,"maxValue" & "_" & iRowCount)
	waiveOnPastDue = Get_Dictionary(ParamValDict,"waiveOnPastDue" & "_" & iRowCount)
	waiveIfCancelled = Get_Dictionary(ParamValDict,"waiveIfCancelled" & "_" & iRowCount)
	strCurrency = Get_Dictionary(ParamValDict,"strCurrency" & "_" & iRowCount)
	strDesc = "BonusFee" & randomNumber(100,999)
	availableFrom = getEffectiveDate()
	companyName = Get_Dictionary(ParamValDict,"cmyName" & "_" & iRowCount)
	fuelNetwork = Get_Dictionary(ParamValDict,"fuelNetwork" & "_" & iRowCount)
	SiteGroup = Get_Dictionary(ParamValDict,"SiteGroup" & "_" & iRowCount)
	SiteID = Get_Dictionary(ParamValDict,"SiteID" & "_" & iRowCount)
	strDateEff = getEffectiveDate()
	prodGroup = Get_Dictionary(ParamValDict,"prodGroup" & "_" & iRowCount)
	rProduct = Get_Dictionary(ParamValDict,"RProduct" & "_" & iRowCount)
	feeValue = Get_Dictionary(ParamValDict,"feeValue" & "_" & iRowCount)
	strDateEff = Get_Dictionary(ParamValDict,"strDateEff" & "_" & iRowCount)
	
	'db_FeeRuleID, db_FeeTypeID, db_billingQuan, db_jobID , db_CustomerID
' already initialized strDesc
	customerERP_id ="NL00000822"
	cust_erp = customerERP_id
	
	If strExecute = "Yes_"& countryName Then
		Call launchApplicationnewGFN()
		Call fees_navigateMaintainFeeRules(feeType,strDesc,freq,feeproduct,productGroup,feeBasis,volInclRate,waiveOnPastDue,waiveIfCancelled,minValue,maxValue,availableFrom,strCurrency)
		Call fees_navigateFeeRuleLocation(companyName,fuelNetwork,SiteGroup,SiteID)
		Call validate_feeRuleLocationTableData(companyName,fuelNetwork,SiteGroup,SiteID,strDateEff)
		Call fees_navigateFeeRuleProduct(prodGroup,rProduct)
		Call validate_feeRuleProductTableData(prodGroup,rProduct,strDateEff)
		Call navigate_feeRuleTierScreen(feeValue)
		Call validate_feeRuleTierTableData(strDateEff,feeValue)
		Call fees_navigateFeeRuleTier_maintainFeeRule()
		Call dobackdatesForFeeRuleid(ui_feeRuleDesc,strDateEff)
		Call RunJob("4")
			Call getFeeNextCreationDateValues()
			Call launchApplicationnewGFN()
			Call customerSearch1(customerERP_id)
			Call bonus_addNewFeeRule(feeType,waiveOnPastDue,waiveIfCancelled,"Pay to Payer")
			Call verifyTblElement(".*SearchControl_grdResults","Fee Rule Details",ui_feeRuleDesc)
			'set fee_data = CreateObject("Scripting.Dictionary")
			'	fee_data.add "custERP", customerERP_id
			'	fee_data.add "items", "1"
			'	fee_data.add "item1", ui_feeRuleDesc
			'	Call addCustFeeRule(fee_data)
				sbstatus = getsalesItemBonusEntries()
				if sbstatus Then
					Call RunJob("4")
					if bFlag Then
						Call validateFeeUnbilledEntries()
					End If
				Else
				
				
				End IF
	End If
next