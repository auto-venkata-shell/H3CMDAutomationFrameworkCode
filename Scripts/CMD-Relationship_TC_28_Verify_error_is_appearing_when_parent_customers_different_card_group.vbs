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
		externalRefVal = myrandNo &"_" & "Parent1_"& lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"1",newERPCustNo,newERPCustNo,externalRefVal)
		topP1LevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
		inputfilepathval = "Sublevelcustomer.txt"
		outputfilepathval = "Sublevelcustomeroutput.txt"
		
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		Call basicPreConfigData(strJson,"GFN_SHELL_SPRINTQA_PH_OLTP")
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Child_"&lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"2",newERPCustNo,topP1LevelERP,externalRefVal)
		subCLevelERP = newERPCustNo
		'---------------------------------------------------------------------------------------------------
			inputfilepathval = "Toplevelcustomer.txt"
		outputfilepathval = "Toplevelcustomeroutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Parent2_"& lcase(countryCode) & "_" & Right(newERPCustNo,6)
		Call preConfigJSONreq_and_Response(strJson,"1",newERPCustNo,newERPCustNo,externalRefVal)
		topP2LevelERP = newERPCustNo
		Call updateCustomerCardgroupTypeValue(topP1LevelERP,topP2LevelERP,"different")
		
		'---------------------------------------------------------------------------------------------------
		
		inputfilepathval = "Relationship.txt"
		outputfilepathval = "Relationshipoutput.txt"
		strJson = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & inputfilepathval 
		newfile = sCurrentDirectory & "Test Data\CMD\ToplevelCustomer\" & outputfilepathval
		'newERPCustNo =  preDatafetchingForEntityType1("GFN_SHELL_SPRINTQA_PH_OLTP")
		externalRefVal = myrandNo &"_" & "Rel_"&lcase(countryCode) & "_" & Right(subCLevelERP,6)
		allcLErp = topP1LevelERP & "|" & topP2LevelERP & "|" & topP1LevelERP & "|" & topP1LevelERP & "|" & topP1LevelERP & "|" & topP1LevelERP & "|" & topP1LevelERP & "|" & topP1LevelERP
		Call preConfigJSONreq_and_ErrorResponse(strJson,"3",subCLevelERP,allcLErp,externalRefVal)
		RelationshipLevelERP = subCLevelERP
		Call validateErrorCustomerRelationshipData(strJson,RelationshipLevelERP)
		Call navigateToRelationshipManagementScreen(RelationshipLevelERP)
		'---------------------------------------------------------------------------------------------------
				
	End If
next