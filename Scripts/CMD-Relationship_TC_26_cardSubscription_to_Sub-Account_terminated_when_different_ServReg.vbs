Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 
	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)
	strExecutevalue = "Yes" & "_" & countryName
	If ucase(strExecute) = ucase(strExecutevalue) Then
		
		inputfilepathval = "Toplevelcustomer.txt"
		outputfilepathval = "Toplevelcustomeroutput.txt"
		'*********** Initialize Global Data for JSON Request *****************************************
		'Below flag for checking default price profile and fee values
		defaultcheckflag = True
		
		reqType = "POST"
		apiurl = api_url & "/Customer/MasterData"

		asynctype = false
		strType = "file"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		jsonFPath = strJson
		
		
		
		
		Apconfigpath=sCurrentDirectory & "API\API_Config.xlsx"
		
		'*********** Generic Code for all scripts ****************
		myrandNo = getRandomNumFromTextFile()
		Call basicPreConfigData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "GrandParent_"& lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"1",newERPCustNo,newERPCustNo,externalRefVal)
		topLevelERP = newERPCustNo
		service1toplevelERP = topLevelERP
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Address.txt"
		outputfilepathval = "Addressoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigAddressData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "address01"&lcase(countryCode) & Right(topLevelERP,4)
		Call preConfigJSONreq_and_Response1(strJson,"4",topLevelERP,"",externalRefVal)
		topnewAddressref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Contact.txt"
		outputfilepathval = "Contactoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigContactData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "contact01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"7",topLevelERP,topnewAddressref,externalRefVal)
		topnewContactref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Bank.txt"
		outputfilepathval = "Bankoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Bank01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"5",topLevelERP,"",externalRefVal)
		topnewBankref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "CustomerPaymentDetail.txt"
		outputfilepathval = "CustomerPaymentDetailoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Payment01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"13",topLevelERP,"",externalRefVal)
		topnewPaymentref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "CardTypesEmpty.txt"
		outputfilepathval = "CardTypesoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "CardTypes01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"12",topLevelERP,"",externalRefVal)
		topnewBankMandateref1 = externalRefVal
		'---------------------------------------------------------------------------------------------------
		Call navigateServicesPage(4)
		'------------------------------------------------------------------------------------------------
		inputfilepathval = "Toplevelcustomer.txt"
		outputfilepathval = "Toplevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval	
		
		myrandNo = getRandomNumFromTextFile()
		Call basicPreConfigData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "GrandParent_"& lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"1",newERPCustNo,newERPCustNo,externalRefVal)
		topLevelERP = newERPCustNo
		service2toplevelERP = topLevelERP
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Address.txt"
		outputfilepathval = "Addressoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigAddressData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "address01"&lcase(countryCode) & Right(topLevelERP,4)
		Call preConfigJSONreq_and_Response1(strJson,"4",topLevelERP,"",externalRefVal)
		topnewAddressref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Contact.txt"
		outputfilepathval = "Contactoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigContactData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "contact01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"7",topLevelERP,topnewAddressref,externalRefVal)
		topnewContactref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Bank.txt"
		outputfilepathval = "Bankoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Bank01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"5",topLevelERP,"",externalRefVal)
		topnewBankref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "CustomerPaymentDetail.txt"
		outputfilepathval = "CustomerPaymentDetailoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Payment01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"13",topLevelERP,"",externalRefVal)
		topnewPaymentref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "CardTypesEmpty.txt"
		outputfilepathval = "CardTypesoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "CardTypes01"&lcase(countryCode) & Right(topLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"12",topLevelERP,"",externalRefVal)
		topnewBankMandateref1 = externalRefVal
		'---------------------------------------------------------------------------------------------------
		Call navigateServicesPage(5)
		'--------------------------------------------------------------------------------------------------
		myrandNo = getRandomNumFromTextFile()
		inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,service1toplevelERP,externalRefVal)
		subPLevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Address.txt"
		outputfilepathval = "Addressoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigAddressData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "address01"&lcase(countryCode) & Right(subPLevelERP,4)
		Call preConfigJSONreq_and_Response1(strJson,"4",subPLevelERP,"",externalRefVal)
		topnewAddressref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Contact.txt"
		outputfilepathval = "Contactoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigContactData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "contact01"&lcase(countryCode) & Right(subPLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"7",subPLevelERP,topnewAddressref,externalRefVal)
		topnewContactref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Bank.txt"
		outputfilepathval = "Bankoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Bank01"&lcase(countryCode) & Right(subPLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"5",subPLevelERP,"",externalRefVal)
		topnewBankref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "CustomerPaymentDetail.txt"
		outputfilepathval = "CustomerPaymentDetailoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Payment01"&lcase(countryCode) & Right(subPLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"13",subPLevelERP,"",externalRefVal)
		topnewPaymentref = externalRefVal
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "CardTypesEmpty.txt"
		outputfilepathval = "CardTypesoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'Call basicPreConfigBankData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "CardTypes01"&lcase(countryCode) & Right(subPLevelERP,6)
		Call preConfigJSONreq_and_Response1(strJson,"12",subPLevelERP,"",externalRefVal)
		topnewBankMandateref1 = externalRefVal
		'---------------------------------------------------------------------------------------------------
		
		
		msgbox "Check Data"
		
		wait 50
		
		'---------------------------------------------------------------------------------------------------
		myrandNo = getRandomNumFromTextFile()	
		inputfilepathval = "Relationship.txt"
		outputfilepathval = "Relationshipoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "ThreelevelRel_"&lcase(countryCode) & "_" & Right(subPLevelERP,6)
		allcLErp = service1toplevelERP & "|" & service1toplevelERP & "|" & subPLevelERP & "|" & service2toplevelERP & "|" & service1toplevelERP & "|" & service1toplevelERP & "|" & service1toplevelERP & "|" & service1toplevelERP
		Call preConfigJSONreq_and_Response(strJson,"3",subPLevelERP,allcLErp,externalRefVal)
		RelationshipLevelERP = child1LevelERP
		Call validateCustomerRelationshipData(strJson,RelationshipLevelERP)
		Call navigateToRelationshipManagementScreen(RelationshipLevelERP)
		'---------------------------------------------------------------------------------------------------
				
	End If
next