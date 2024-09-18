Dim DictTbl	
iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	strDateEff = Get_Dictionary(ParamValDict,"strDateEff" & "_" & iRowCount)
	feeType = Get_Dictionary(ParamValDict,"feeType" & "_" & iRowCount)
	FeeTypeID = Get_Dictionary(ParamValDict,"FeeTypeID" & "_" & iRowCount)
	bonusID = "Pay to Payer"
	strDesc = "BonusFee88"
	If strExecute = "Yes_"& countryName Then
		db_CustomerERP = fees_DBqueries(strDesc,strDateEff)
		Call launchApplicationnewGFN()
		Call customerSearch1(db_CustomerERP)
		Call bonus_addNewFeeRule(feeType,waiveOnPastDue,waiveIfCancelled,bonusID)
		Call verifyTblElement(".*SearchControl_grdResults","Fee Rule Details",strDesc)
		db_FeeRuleID = afterCustomerUpdate(strDateEff,db_CustomerERP,FeeTypeID)
		Call fees_navigateBillingAcceptance()
		db_BillingDocumentID = afterBillingPreview(db_FeeRuleID)
		Call click_on_cutoff()
		Call click_on_signoff()
		Call RunJob("19")
		Call RunJob("64")
		Call getBillingDocumentID(db_FeeRuleID)
	End If
next