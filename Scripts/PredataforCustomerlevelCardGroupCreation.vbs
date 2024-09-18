Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	'customerERP_id = Get_Dictionary(ParamValDict,"customerERP_id" & "_" & iRowCount)
	embossName = Get_Dictionary(ParamValDict,"embossName" & "_" & iRowCount)	
	cardType = Get_Dictionary(ParamValDict,"cardType" & "_" & iRowCount)
	embossType = Get_Dictionary(ParamValDict,"embossType" & "_" & iRowCount)
	cardDelPoint = Get_Dictionary(ParamValDict,"cardDelPoint" & "_" & iRowCount)	
	new_countryName = Get_Dictionary(ParamValDict,"countryName" & "_" & iRowCount)
	address = Get_Dictionary(ParamValDict,"address" & "_" & iRowCount)	 
	postalCode = Get_Dictionary(ParamValDict,"postalCode" & "_" & iRowCount)	
	email = Get_Dictionary(ParamValDict,"email" & "_" & iRowCount)
	DelOneTimeUse = Get_Dictionary(ParamValDict,"DelOneTimeUse" & "_" & iRowCount)
	PinOneTimeUse = Get_Dictionary(ParamValDict,"PinOneTimeUse" & "_" & iRowCount)
	phoneNo = Get_Dictionary(ParamValDict,"phoneNo" & "_" & iRowCount) 
	city = Get_Dictionary(ParamValDict,"city" & "_" & iRowCount)
	region = Get_Dictionary(ParamValDict,"region" & "_" & iRowCount) 
		

	If strExecute = "Yes_"& countryName Then			
		custERP = customerERP_id
		cardGrpName = "CARDGROUPNEW" & randomNumber(100,999)
		Call launchApplicationnewGFN()
		Call customerSearch1(customerERP_id)
		Call createNewCardGroup(cardGrpName,"#1",cardDelPoint)
		Call cardDelivery_overRideAddress(new_countryName,address,postalCode,email,phoneNo)
		wait 10
		Call DBvalidation_cardGroup(cardGrpName,address,postalCode,email,phoneNo)
		'Call GFN_CreateCard(cardGrpName, embossName,cardType, embossType)
		'Call verify_pinDeliveryAddress_enabled()
		'Call createCard_overRideAddress_DeliveryCard(DelOneTimeUse,email,phoneNo,new_countryName,postalCode)
		'Call createCard_overRideAddress_DeliveryPin(PinOneTimeUse,email,phoneNo,new_countryName,address,postalCode,city,region)
		'Call click_on_saveElement()
		'Call RunJob("196")
		'Call launchApplicationnewGFN()
		'Call ui_changes_navigate_searchForCards("Active",cardPANNum)
		'Call perform_ReprintPIN()
		'Call DBvalidation_ReprintPIN(cardPANNum)
		
	else
		Print "StrExecute value is false"	
	End If
	
next