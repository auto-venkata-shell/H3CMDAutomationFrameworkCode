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
		topGPLevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Parent_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,topGPLevelERP,externalRefVal)
		subPLevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child1_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		allParentERPs = topGPLevelERP & "|" & subPLevelERP
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,allParentERPs,externalRefVal)
		child1LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & ",Child2_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		allParentERPs = topGPLevelERP & "|" & subPLevelERP
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,allParentERPs,externalRefVal)
		child2LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child3_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		allParentERPs = topGPLevelERP & "|" & subPLevelERP
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,allParentERPs,externalRefVal)
		child3LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child4_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		allParentERPs = topGPLevelERP & "|" & subPLevelERP
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,allParentERPs,externalRefVal)
		child4LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child5_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		allParentERPs = topGPLevelERP & "|" & subPLevelERP
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,allParentERPs,externalRefVal)
		child5LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child6_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		allParentERPs = topGPLevelERP & "|" & subPLevelERP
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,allParentERPs,externalRefVal)
		child6LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Relationship.txt"
		outputfilepathval = "Relationshipoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "ThreelevelRel_"&lcase(countryCode) & "_" & Right(subPLevelERP,6)
		allcLErp = topGPLevelERP & "|" & subPLevelERP & "|" & child1LevelERP & "|" & child2LevelERP & "|" & child3LevelERP & "|" & child4LevelERP & "|" & child5LevelERP & "|" & child6LevelERP
		Call preConfigJSONreq_and_Response(strJson,"3",child1LevelERP,allcLErp,externalRefVal)
		RelationshipLevelERP = child1LevelERP
		Call validateCustomerRelationshipData(strJson,RelationshipLevelERP)
		Call navigateToRelationshipManagementScreen(RelationshipLevelERP)
		'---------------------------------------------------------------------------------------------------
				
	End If
next