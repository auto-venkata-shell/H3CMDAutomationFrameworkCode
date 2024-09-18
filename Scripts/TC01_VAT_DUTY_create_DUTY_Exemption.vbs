Dim DictTbl	

iTotalRows = ParamValDict.Item("DATAROWS")
For iRowCount=1 to iTotalRows 	
	strExecute = Get_Dictionary(ParamValDict,"Execute" & "_" & iRowCount)	
	typeOfExe = Get_Dictionary(ParamValDict,"typeOfExe" & "_" & iRowCount)
	companyID = Get_Dictionary(ParamValDict,"companyID" & "_" & iRowCount)	
	
	countryID = Get_Dictionary(ParamValDict,"countryID" & "_" & iRowCount)
	VATCategoryID = Get_Dictionary(ParamValDict,"VATCategoryID" & "_" & iRowCount)	
	productVal = Get_Dictionary(ParamValDict,"productVal" & "_" & iRowCount)
	descrip = "DUTY_EXEMPTION" & randomNumber(100,999)
	If strExecute = "Yes_"& countryName Then			

		Call launchApplicationnewGFN()
		Call maintainExemptionscheme(typeOfExe,companyID,descrip,countryID,VATCategoryID,productVal)
	else
		Print "StrExecute value is false"	
	End If
	
next