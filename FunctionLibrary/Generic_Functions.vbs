
'******************************************** HEADER ******************************************
' Name : Descriptive_Object_Parser
' Description : generic function to parser the descriptive object properties from the OR FILE 
'				and create the global dictionary object with KEY as OBJECT NAME and value as 
'				comma seperated property name and its value combinations.
' Ex: Wnd_N_POS=text:NAMOS compact POS,nativeclass:#32770
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function Descriptive_Object_Parser()
		
	Dim FileSysObj
	Dim objFile
	Dim strLine
	Dim arrDescObj_Props
	Dim strKeyName
	Dim strValueName

	Set FileSysObj= CreateObject("Scripting.FileSystemObject")
	Set objFile = FileSysObj.OpenTextFile(ORFilePath)

	Do until objFile.AtEndOfStream
		'''msgbox objFile.ReadLine
		
		strLine = objFile.ReadLine
		
		If strLine<>"" and instr(1,strLine,"<GROUP")=0 and instr(1,strLine,"</GROUP")=0 Then
		
				'Split by '='
				arrDescObj_Props = Split(strLine,"=")
				
				strKeyName = arrDescObj_Props(0)
				strValueName = arrDescObj_Props(1)
				
				if objORDict.Exists (strKeyName)=False then
					'objORDict.Add strKeyName,strValueName	
					Call Add_Dictionary(objORDict,strKeyName,strValueName)
				else
					'objORDict.Item(strKeyName) = strValueName		
					Call Update_Dictionary(objORDict,strKeyName,strValueName)
				End if
				
		
		End If
			
	Loop  
		
	'WriteLog "","","","","Descriptive_Object_Parser () : parsing the object descriptive properties and creation of global dictionary is completed successfully."
	
	Set FileSysObj= Nothing 
    Set objFile= Nothing 
End Function


'******************************************** HEADER ******************************************
' Name : Create_Dictionary
' Description : Create the dictionary object
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : newly created dictionary
'******************************************** HEADER ******************************************
Public Function Create_Dictionary()

	Dim objDict
	Set objDict = CreateObject("Scripting.Dictionary")
	
	Set Create_Dictionary = objDict
	
		Set objDict= Nothing 
    

End Function

'******************************************** HEADER ******************************************
' Name : Add_Dictionary
' Description : Add a new item to the dictionary object
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : dictObj, strKey, strValue
' Output Parameter : dictionary object with newly added item
'******************************************** HEADER ******************************************
Public Function Add_Dictionary(byref dictObj,strKey,strValue)

	if dictObj.Exists (strKey)=False then
		dictObj.Add strKey,strValue	
	else
		dictObj.Item(strKey) = strValue		
	End if
	
	Set Add_Dictionary = dictObj
	
End Function

'******************************************** HEADER ******************************************
' Name : Update_Dictionary
' Description : Update existing new item value to the dictionary object
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : dictObj, strKey, strValue
' Output Parameter : dictionary object with updated item value
'******************************************** HEADER ******************************************
Public Function Update_Dictionary(byref dictObj,strKey,strValue)
	if dictObj.Exists (strKey)=False then
		dictObj.Add strKey,strValue	
	else
		dictObj.Item(strKey) = strValue		
	End if
	
End Function

'******************************************** HEADER ******************************************
' Name : Get_Dictionary
' Description : Get existing item value from the dictionary object with Key provided
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : dictObj, strKey
' Output Parameter : item value for the specified key
'******************************************** HEADER ******************************************
Public Function Get_Dictionary(byref dictObj,strKey)

	strValue = dictObj.Item(strKey)
	
	Get_Dictionary = strValue
	
End Function

'******************************************** HEADER ******************************************
' Name : Delete_Dictionary
' Description : Delete existing item value from the dictionary object with Key provided
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : dictObj, strKey
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function Delete_Dictionary(byref dictObj,strKey)

	if dictObj.Exists (strKey)=False then
		dictObj.remove strKey
	End if
	
End Function

'******************************************** HEADER ******************************************
' Name : CreateDescriptionObject
' Description : Create descriptive object for the specified Object Name
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : strKeyName
' Output Parameter : Newly created descriptive object
'******************************************** HEADER ******************************************
Public Function CreateDescriptionObject(strKeyName)

If isobject(strKeyName) = False Then
		
	
	
	Dim DescObject
	Set DescObject = description.Create
	
	'Retrieve from the global object properties dictionary 'objORDict'
	strItem = objORDict.item(strKeyName)
		
	'Split the properties values ","
	arrProps = Split(strItem,",")
	
	'Loop through
	For iPropCnt = 0 to ubound(arrProps)
	
		'split by :
		arrPropKeyVal = Split(arrProps(iPropCnt),":")
		
		'in case your property value contains '&' replace it with '=' before updating the descriptive object
		If instr(1,arrPropKeyVal(1),"&&")<>0Then
			sTemp= arrPropKeyVal(1)
		 	sValue = replace(sTemp,"&&","=")
		 	PropName = arrPropKeyVal(0)
			PropValue = sValue
			
		Elseif instr(1,arrPropKeyVal(1),"++")<>0Then
			sTemp= arrPropKeyVal(1)
		 	sValue = replace(sTemp,"++",".*")
		 	PropName = arrPropKeyVal(0)
			PropValue = sValue
			
		Elseif instr(1,arrPropKeyVal(1),"^^")<>0Then
			sTemp= arrPropKeyVal(1)
		 	sValue = replace(sTemp,"^^",":")
		 	PropName = arrPropKeyVal(0)
			PropValue = sValue
		else
			'''msgbox "ubound arrPropKeyVal : " & ubound(arrPropKeyVal)
			PropName = arrPropKeyVal(0)
			PropValue = arrPropKeyVal(1)
		End If
		
		
		DescObject(PropName).value = cstr(PropValue)
	Next
		
	
	'Set DialObj = DescObject
	
	Set CreateDescriptionObject = DescObject

Else
	
	Set CreateDescriptionObject = strKeyName
    
End If	  

End Function

'******************************************** HEADER ******************************************
' Name : CreateDescriptionObject
' Description : Create descriptive object for the specified Object Name and with additional dynamic attributes
' Creator : Syed Shafi
' Date : 23th Sep,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : strKeyName,strAddtionalAttributes
' Output Parameter : Newly created descriptive object
'******************************************** HEADER ******************************************
Public Function CreateDescriptionObjectV2(strKeyName,strAddtionalAttributes)
	
	Dim DescObject
	Set DescObject = description.Create
	
	'Retrieve from the global object properties dictionary 'objORDict'
	strItem = objORDict.item(strKeyName)
		
	'Split the properties values ","
	arrProps = Split(strItem,",")
	
	'Loop through
	For iPropCnt = 0 to ubound(arrProps)
	
		'split by :
		arrPropKeyVal = Split(arrProps(iPropCnt),":")
			
		'in case your property value contains '&' replace it with '=' before updating the descriptive object
		If instr(1,arrPropKeyVal(1),"&&")<>0Then
			sTemp= arrPropKeyVal(1)
		 	sValue = replace(sTemp,"&&","=")
		 	PropName = arrPropKeyVal(0)
			PropValue = sValue
		else
			PropName = arrPropKeyVal(0)
			PropValue = arrPropKeyVal(1)
		End If
		
		DescObject(PropName).value = cstr(PropValue)
	Next
	
	
	'Additional attributes
	If instr(1,strAddtionalAttributes,",")<>0 Then
		'Split the properties values "," if more than one additional attributes
		arrProps1 = Split(strAddtionalAttributes,",")
		'Loop through
		For iPropCnt = 0 to ubound(arrProps1)
		
			'split by :
			arrPropKeyVal1 = Split(arrProps1(iPropCnt),":")
		
			'in case your property value contains '&' replace it with '=' before updating the descriptive object
			If instr(1,arrPropKeyVal1(1),"&&")<>0Then
				sTemp= arrPropKeyVal1(1)
			 	sValue = replace(sTemp,"&&","=")
			 	PropName = arrPropKeyVal1(0)
				PropValue = sValue
			else
				'''msgbox "ubound arrPropKeyVal : " & ubound(arrPropKeyVal)
				PropName = arrPropKeyVal1(0)
				PropValue = arrPropKeyVal1(1)
			End If
			DescObject(PropName).value = cstr(PropValue)
		Next
	else
		If instr(1,strAddtionalAttributes,":")<>0 Then
			'split by :
			arrPropKeyVal1 = Split(strAddtionalAttributes,":")

			'in case your property value contains '&' replace it with '=' before updating the descriptive object
			If instr(1,arrPropKeyVal1(1),"&&")<>0Then
				sTemp= arrPropKeyVal1(1)
			 	sValue = replace(sTemp,"&&","=")
			 	PropName = arrPropKeyVal1(0)
				PropValue = sValue
			else
				'''msgbox "ubound arrPropKeyVal : " & ubound(arrPropKeyVal)
				PropName = arrPropKeyVal1(0)
				PropValue = arrPropKeyVal1(1)
			End If
				
			DescObject(PropName).value = cstr(PropValue)
		End if
	End If
	
	'Set DialObj = DescObject
	''msgbox DescObject
	Set CreateDescriptionObjectV2 = DescObject
     Set DescObject = Nothing
End Function


'******************************************** HEADER ******************************************
' Name : ReadTestDataCSV
' Description : generic function to read the TestData csv file and create a dictionary object of
'				all the test data values which can be used for executing each test case script
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : sCSVFile
' Output Parameter : None
'******************************************** HEADER ******************************************
Function ReadTestDataCSV(sCSVFile)
	
	Dim FileSysObj
	Dim dataFile 
	Dim fields
	Dim ParamDict
	Dim bRowFlag
	Dim RowNum
	
	bRowFlag=False
	RowNum=0
	
	Set ParamDict = CreateObject("Scripting.Dictionary")
	Set FileSysObj = CreateObject("Scripting.FileSystemObject")
	
	Set dataFile = FileSysObj.OpenTextFile(TDFilePath & sCSVFile,1,False)
	
	Do
	
		If bRowFlag=False Then
			'only parameter names to retieved from the 1st row
			
			strLine = dataFile.ReadLine
			
			Row1 = strLine
			'Parameters list split and create a dictionary of parameters
			arrParams = split(Row1,",")
			For iCnt=0 to ubound(arrParams)
				sKey = iCnt+1
				Value = arrParams(iCnt)
				Add_Dictionary ParamDict,sKey,Value
			Next
			bRowFlag=True
		else
		
			strLine = dataFile.ReadLine
			
			fields = Split(strLine ,",")
			RowNum = RowNum + 1
			
			For fldcnt=0 to ubound(fields)
				
				ParamKey = fldcnt+1
				ParamName = Get_Dictionary(ParamDict,ParamKey)
				ParameterValue = fields(fldcnt)
				'Add Each row data by Parameter name 
				Add_Dictionary ParamValDict,ParamName & "_" & RowNum,ParameterValue
				
				reporter.ReportEvent micDone,"STEP " & ParamName & "_" & RowNum & ":" & ParameterValue,ParameterValue
				
				'WriteLog "","","","","ReadTestDataCSV () : reading the test data csv for a the test case is completed successfully."
			Next
		
		End if
	Loop While NOT dataFile.AtEndOfStream
	
	Add_Dictionary ParamValDict,"DATAROWS",RowNum
	
	dataFile.Close

	Set FileSysObj = nothing
	Set ParamDict = Nothing
	Set FileSysObj = Nothing
	
End Function

'******************************************** HEADER ******************************************
' Description : 
' Creator : 
' Date : 
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ClickBosDialog(sBrowser,sDialog,sButton)
	
	  Set bBrowser = CreateDescriptionObject(sBrowser)
      Set bDialog = CreateDescriptionObject(sDialog)
  
      Set bButton = CreateDescriptionObject(sButton)
	
	  Browser(bBrowser).Dialog(bDialog).WinButton(bButton).Click
	  
	
End Function

'******************************************** HEADER ******************************************
' Name : Index_initialize
' Description : Generic function to create an Index file for the execution summary report
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function Index_initialize()

	Dim sDestinationFolder
	Dim TextLog


	sFolderName = "Result_" & DateTimeRandomNumber

	sLogFolder = environment.Value("sCurrentDirectory") & "Logs"

	logOverrideFile = TestResults & "log.txt"

	If fsObject.FileExists(logOverrideFile) = False then
		Set fsoFile = fsObject.CreateTextFile (logOverrideFile)
	else
		Set fsoFile = fsObject.OpenTextFile(logOverrideFile)
	End if

	Dim sDate : sDate = Day(Now)
	Dim sMonth : sMonth = Month(Now)
	Dim sYear : sYear = year(Now)

	If sDate<10 Then
		sDate=0 & sDate
	End If

	If sMonth<10 Then
		sMonth=0 & sMonth
	End If

	sTodayDate= sDate&sMonth&sYear

	'strGetOnlyBOS = GetBOSInstruction("ONLYBOS")

	'If strGetOnlyBOS="YES" Then

		TextLog = "Log_" & sTodayDate& sFolderName & ".txt"

		strindexFile = TestResults & sFolderName & "\index.html"

		Update_Dictionary objDictBOSExecInstruction,"RESULTFILE",strindexFile

		ScreenshotPath = TestResults & sFolderName & "\"

		sDestinationFolder = TestResults & sFolderName

		ConfigDest = TestResults & sFolderName & "\" & Config

		If fsObject.FolderExists(sDestinationFolder) = False then
			fsObject.CreateFolder (sDestinationFolder)
		End if
		If fsObject.FolderExists(sLogFolder) = False then
			fsObject.CreateFolder (sLogFolder)
		End if


		TextLogDest = sLogFolder & "\" & TextLog



		If fsObject.FileExists(TextLogDest) = False then
			Set fsoFile = fsObject.CreateTextFile (TextLogDest)
		else
			Set fsoFile = fsObject.OpenTextFile(TextLogDest)
		End if

	'	set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
	'	objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
	'	objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"
	'	------------
		set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
		objFile_html.writeline "<head><style>body {    background-color: #E0FFFF;}</style></head>"
		objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
		objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"
	'	--------------
		'WriteLog "","","","","Index_initialize () : Initilization function executed successfully."
'	else
'		strParentIndexFile=GetBOSInstruction("RESULTFILE")
'
'		sRequireDate=Environment.Value("bAdhocStartDate")
'
'		If instr(sRequireDate,"/")<>0 Then
'			sRequireDate=replace(sRequireDate,"/","")
'		End If
'
'		If strParentIndexFile<>"" Then
'			strPOSResultsPath = mid(strParentIndexFile,1,instr(1,strParentIndexFile,"index.html")-1)
'		End If
'
'		sResults=split(strPOSResultsPath,"\")
'		sResultFolder = sResults(ubound(sResults)-1)
'
'		TextLog = "Log_" & sRequireDate& sResultFolder & ".txt"
'
'		strindexFile = strPOSResultsPath & "BOS.txt"
'
'		strSummaryFile = strPOSResultsPath & "Summary.html"
'		strSummaryFileAdhoc = strPOSResultsPath & "SummaryAdhoc.html"
'		strSummaryFileEOD = strPOSResultsPath & "SummaryEOD.html"
'		strSummaryFileAdhocdts = strPOSResultsPath & "SummaryAdhoc" & DateRandomNumber & ".html"
'		strSummaryFileEODdts = strPOSResultsPath & "SummaryEOD" & DateRandomNumber & ".html"
'
'		ScreenshotPath = strPOSResultsPath & "\"
'
'		sDestinationFolder = strPOSResultsPath
'		TextLogDestArr = split(strPOSResultsPath,"Results")
'		TextLogDestTemp = TextLogDestArr(0)
'		TextLogDest = TextLogDestTemp&"Logs\"&TextLog
'		
'		
'		'TextLogDest = strPOSResultsPath & "\" & TextLog
'		ConfigDest = strPOSResultsPath & "\" & TextLog
'
'
'		If fsObject.FolderExists(sDestinationFolder) = False and sDestinationFolder<>"" then
'			fsObject.CreateFolder (sDestinationFolder)
'		End if
'
'		If fsObject.FileExists(TextLogDest) = False then
'			Set fsoFile = fsObject.CreateTextFile (TextLogDest)
'		else
'			Set fsoFile =fsObject.OpenTextFile (TextLogDest,8,true)
'		End if
'
'
'		If strindexFile<>"" Then
'			set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
'		End If
'		If strParentIndexFile<>"" Then
'			set objFile_Parenthtml= fsObject.OpenTextFile(strParentIndexFile,1,True)
'		End If
'	End If


	'COMMENT ::: Not writing table headers as we can use the parent index file created in POS execution

'	objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
'	objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"

	'WriteLog "","","","","Index_initialize () : Initilization function executed successfully."

End Function

Public Function Index_initializeBkp()
	
	Dim sDestinationFolder
	Dim TextLog
	Dim TextLogDest
	
	sFolderName = "Result_" & DateTimeRandomNumber
	
	TextLog = "Log_" & DateTimeRandomNumber & ".txt"
	
	strGetOnlyBOS = GetBOSInstruction("ONLYBOS")
	
	If strGetOnlyBOS="YES" Then
	
		strindexFile = TestResults & sFolderName & "\index.html"
		strPDFFolder = TestResults & sFolderName & "\PDF"
		
		Update_Dictionary objDictBOSExecInstruction,"RESULTFILE",strindexFile
		
		ScreenshotPath = TestResults & sFolderName & "\"
		
		sDestinationFolder = TestResults & sFolderName
		TextLogDest = TestResults & sFolderName & "\" & TextLog
		
		If fsObject.FolderExists(sDestinationFolder) = False then
			fsObject.CreateFolder (sDestinationFolder)
		End if
		
		If fsObject.FileExists(TextLogDest) = False then
			Set fsoFile = fsObject.CreateTextFile (TextLogDest)
		End if
		
	'	set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
	'	objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
	'	objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"
	'	------------
		set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
		objFile_html.writeline "<head><style>body {    background-color: #E0FFFF;}</style></head>"
		objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
		objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"
	'	--------------
		'WriteLog "","","","","Index_initialize () : Initilization function executed successfully."
	else		
	strParentIndexFile=GetBOSInstruction("RESULTFILE")

	strPOSResultsPath = mid(strParentIndexFile,1,instr(1,strParentIndexFile,"index.html")-1)

	strindexFile = strPOSResultsPath & "BOS.txt"
	
	strPDFFolder = strPOSResultsPath &"PDF"
		
	If fsObject.FolderExists(strPDFFolder)=false Then
		fsObject.CreateFolder (strPDFFolder)
	End If
	
	strSummaryFile = strPOSResultsPath & "Summary.html"
	strSummaryFileAdhoc = strPOSResultsPath & "SummaryAdhoc.html"
	strSummaryFileEOD = strPOSResultsPath & "SummaryEOD.html"
	strSummaryFileAdhocdts = strPOSResultsPath & "SummaryAdhoc" & DateRandomNumber & ".html"
	strSummaryFileEODdts = strPOSResultsPath & "SummaryEOD" & DateRandomNumber & ".html"
	
	ScreenshotPath = strPOSResultsPath & "\"
	
	sDestinationFolder = strPOSResultsPath
	TextLogDest = strPOSResultsPath & "\" & TextLog
	
	If fsObject.FolderExists(sDestinationFolder) = False then
		fsObject.CreateFolder (sDestinationFolder)
	End if
	
	If fsObject.FileExists(TextLogDest) = False then
		Set fsoFile = fsObject.CreateTextFile (TextLogDest)
	End if
	
	set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
	set objFile_Parenthtml= fsObject.OpenTextFile(strParentIndexFile,1,True)
	
	'COMMENT ::: Not writing table headers as we can use the parent index file created in POS execution
	
'	objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
'	objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"
	
	'WriteLog "","","","","Index_initialize () : Initilization function executed successfully."
	End If		 
End Function

Public Function Index_initializeOLD()
	
	Dim sDestinationFolder
	Dim TextLog
	Dim TextLogDest
	
	sFolderName = "Result_" & DateTimeRandomNumber
	
	TextLog = "Log_" & DateTimeRandomNumber & ".txt"
	
	strParentIndexFile=GetBOSInstruction("RESULTFILE")

	strPOSResultsPath = mid(strParentIndexFile,1,instr(1,strParentIndexFile,"index.html")-2)

	strindexFile = strPOSResultsPath & "\BOS.txt"
	strSummaryFile = strPOSResultsPath & "\Summary.html"
	strSummaryFileAdhoc = strPOSResultsPath & "\SummaryAdhoc.html"
	strSummaryFileEOD = strPOSResultsPath & "\SummaryEOD.html"
	
	ScreenshotPath = strPOSResultsPath & "\"
	
	sDestinationFolder = strPOSResultsPath
	TextLogDest = strPOSResultsPath & "\" & TextLog
	
	If fsObject.FolderExists(sDestinationFolder) = False then
		fsObject.CreateFolder (sDestinationFolder)
	End if
	
	If fsObject.FileExists(TextLogDest) = False then
		Set fsoFile = fsObject.CreateTextFile (TextLogDest)
	End if
	
	set objFile_html= fsObject.OpenTextFile(strindexFile,2,True)
	set objFile_Parenthtml= fsObject.OpenTextFile(strParentIndexFile,1,True)
	
	'COMMENT ::: Not writing table headers as we can use the parent index file created in POS execution
	
'	objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
'	objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Sl.No</font></th><th BGCOLOR='879773'><font size=2 >Script Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Duration(Mins)</font></th><th BGCOLOR='879773'><font size=2>Status</font></th></tr>"
	
	'WriteLog "","","","","Index_initialize () : Initilization function executed successfully."
			 
End Function
'******************************************** HEADER ******************************************
' Name : WriteLog
' Description : Generic function to write log into text file
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : iStep,sStepName,sStepDescrption,sStatus,sOptDescription
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function WriteLog(iStep,sStepName,sStepDescrption,sStatus,sOptDescription)
		
	If sOptDescription <> "" Then
		fsoFile.writeline Now & " => " & sOptDescription 
	else
	
		fsoFile.writeline Now &  " => " & iStep & " : " & " : " & sStepName & " : " & sStepDescrption & " : " & sStatus	
	End If
	
	
End Function

Public Function WriteLog1(sLogLevel,sSource,sLogDescription)

	'sDate=DateOperation("TodayDate","","")
	'sDate=replace(sDate,"/","-")
    'Convert it into YYYY-MM-DD
    sDateT=DateOperation("TodayDate","","")
	sDateTT = split(sDateT,"/")
    sDate = sDateTT(2)&"-"&sDateTT(1)&"-"&sDateTT(0)


	If sLogDescription <> "" Then
		fsoFile.writeline sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000) & " => " & sLogLevel & ":" & sTCScriptName & " => " & sSource & " => " & sLogDescription
	else
        fsoFile.writeline sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000) & " => " & sLogLevel & ":" & sTCScriptName & " => " & sSource & " => " & sLogDescription

		'fsoFile.writeline Now &  " => " & iStep & " : " & " : " & sStepName & " : " & sStepDescrption & " : " & sStatus
	End If


End Function


Public Function WriteConfig1()

    'Convert it into YYYY-MM-DD
    sDateT=DateOperation("TodayDate","","")
    sDateTT = split(sDateT,"/")
    sDate = sDateTT(2)&"-"&sDateTT(1)&"-"&sDateTT(0)

    'config
    Config = "Results\Config.txt"
    ConfigDest = sCurrentDirectory & Config

    Set objFSO = createobject("Scripting.FileSystemObject")

      If objFSO.FileExists(ConfigDest) = False Then
           Set fsoConfig = fsObject.CreateTextFile (ConfigDest)
           sBatchNo = 1
      Else
       
       sBatchNo = 1
       Set fsoConfig = objFSO.OpenTextFile(ConfigDest,1)

       Do Until fsoConfig.AtEndOfStream
            strLine = fsoConfig.ReadLine
            If instr(strLine,"BatchNo")> 0 Then
                  sCurrentBatch = split(strLine,"==>")(1)
                  sBatchNo = sCurrentBatch+1
            End If
        Loop

      End If
       
        fsoConfig.Close
        Set fsoConfig = Nothing
        Set objFSO = Nothing
       
        wait 5
       
        Set objFSO1 = createobject("Scripting.FileSystemObject")
       
        Set fsoConfig1 = objFSO1.OpenTextFile(ConfigDest,2)

        'Timestamp==>
        fsoConfig1.writeline "Timestamp==>" & sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000)

        'Path==>
        fsoConfig1.writeline "Path==>" & Environment.Value("Solution") & "\" & Environment.Value("Market") & "\" & Environment.Value("ReleaseName") & "\" & sDate

        'BatchNo==>
        fsoConfig1.writeline "BatchNo==>"&sBatchNo

        'ReleaseName==>
        fsoConfig1.writeline "Release==>"&Environment.Value("ReleaseName")

        'ModuleName==>
        fsoConfig1.writeline "Module==>"&Environment.Value("ModuleName")
       
        fsoConfig1.Close
        Set fsoConfig1 = Nothing
        Set objFSO1 = Nothing
       
End Function

'******************************************** HEADER ******************************************
' Name : DeInitialize
' Description : de initialize the index file object instances
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function DeInitialize()
	
	objFile_html.close
	
	'strGetOnlyBOS = GetBOSInstruction("ONLYBOS")
	'If ucase(strGetOnlyBOS)="YES" Then
		
		systemutil.Run "msedge.exe",strindexFile
'	Else
'	
'		sPOSResults = objFile_Parenthtml.ReadAll
'	    objFile_Parenthtml.Close
'	    
'	    set objFile_BOSResult= fsObject.OpenTextFile(strindexFile,1)
'	    sBOSResults = objFile_BOSResult.ReadAll
'	    
'	    set objFile_Summaryhtml= fsObject.OpenTextFile(strSummaryFile,2,True)
'	    
'	    'objFile_Parenthtml.Close
'	    
'	    If instr(1,sPOSResults,"@BOSHTMLRESULTS@") <>0 Then
'	    	strNewSumary = replace(sPOSResults,"@BOSHTMLRESULTS@",sBOSResults)
'	    	objFile_Summaryhtml.writeline strNewSumary
'	    End If
'	    objFile_Summaryhtml.Close
'	    
'		wait 5
'		
'		If fsObject.FileExists(strSummaryFileAdhoc) = False Then
'			set objFileEod = fsObject.CreateTextFile(strSummaryFileAdhoc)
'		End If
'		
'		If fsObject.FileExists(strSummaryFileEOD) = False Then
'			set objFileEod = fsObject.CreateTextFile(strSummaryFileEOD)
'		End If
'	
'	
'
'    If Strcomp(ucase(Environment.Value("bAdhocEODFlag")), "ADHOC",1) = 0 Then
'    
'    	set fsObject_adhoc = fsObject.GetFile(strSummaryFileAdhoc)
'    		If fsObject_adhoc.Size > 0 Then
'    	 		set objFile_Summaryhtml_Adhocdts = fsObject.OpenTextFile(strSummaryFileAdhocdts,2,True)
'	    		objFile_Summaryhtml_Adhocdts.writeline strNewSumary
'	    		objFile_Summaryhtml_Adhocdts.Close
'	    	Else
'	    		set objFile_Summaryhtml_Adhoc = fsObject.OpenTextFile(strSummaryFileAdhoc,2,True)
'	    		objFile_Summaryhtml_Adhoc.writeline strNewSumary
'	    		objFile_Summaryhtml_Adhoc.Close
'	    	End If
'	    	
'    Elseif Strcomp(ucase(Environment.Value("bAdhocEODFlag")),"EOD",1) = 0 Then
'     
'    	set fsObject_eod = fsObject.GetFile(strSummaryFileEOD)
'    		
'    		If fsObject_eod.Size > 0 Then
'    			set objFile_Summaryhtml_EODdts = fsObject.OpenTextFile(strSummaryFileEODdts,2,True)
'		    		objFile_Summaryhtml_EODdts.writeline strNewSumary
'		    		objFile_Summaryhtml_EODdts.Close
'		    Else
'		    		set objFile_Summaryhtml_EOD = fsObject.OpenTextFile(strSummaryFileEOD,2,True)
'		    		objFile_Summaryhtml_EOD.writeline strNewSumary
'	    			objFile_Summaryhtml_EOD.Close
'    		End If
' 
'   End If
	
	Wait 5
	
	set objFile_Summaryhtml = nothing
	Set objFile_Summaryhtml_Adhoc = nothing
	Set objFile_Summaryhtml_EOD = nothing
	Set objFile_Summaryhtml_Adhocdts = nothing
	Set objFile_Summaryhtml_EODdts = nothing
	
	'open the execution summary page on i-explore
'	systemutil.Run "iexplore",strSummaryFile
'End if
End Function

Public Function DeInitializeOLD()
	
	objFile_html.close
	
	sPOSResults = objFile_Parenthtml.ReadAll
    objFile_Parenthtml.Close
    
    set objFile_BOSResult= fsObject.OpenTextFile(strindexFile,1)
    sBOSResults = objFile_BOSResult.ReadAll
    
    set objFile_Summaryhtml= fsObject.OpenTextFile(strSummaryFile,2,True)
    
    
    
    'objFile_Parenthtml.Close
    
    If instr(1,sPOSResults,"@BOSHTMLRESULTS@") <>0 Then
    	strNewSumary = replace(sPOSResults,"@BOSHTMLRESULTS@",sBOSResults)
    	objFile_Summaryhtml.writeline strNewSumary
    End If
    objFile_Summaryhtml.Close
    
    wait 5
	
    If Strcomp(ucase(Environment.Value("bAdhocEODFlag")), "ADHOC",1) = 0 Then 'ADHOC
    	set objFile_Summaryhtml_Adhoc = fsObject.OpenTextFile(strSummaryFileAdhoc,2,True)
        objFile_Summaryhtml_Adhoc.writeline strNewSumary
    	objFile_Summaryhtml_Adhoc.Close
    Elseif Strcomp(ucase(Environment.Value("bAdhocEODFlag")),"EOD",1) = 0 Then
    	set objFile_Summaryhtml_EOD = fsObject.OpenTextFile(strSummaryFileEOD,2,True)
       	objFile_Summaryhtml_EOD.writeline strNewSumary
    	objFile_Summaryhtml_EOD.Close
    End If
	
	Wait 5

	'open the execution summary page on i-explore
	systemutil.Run "iexplore",strSummaryFile
	
End Function
'******************************************** HEADER ******************************************
' Name : Initialize_TestHTML
' Description : Generic function to create an Index file for each test case and append the same to summary report
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : strTC
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function Initialize_TestHTML(strTC)
	
	'strGetOnlyBOS = GetBOSInstruction("ONLYBOS")
	'If strGetOnlyBOS="YES" Then
		StepCounter = 1
		strTestFile = TestResults & sFolderName & "\" & strTC & ".html"
		
		set objFile_Testhtml= fsObject.OpenTextFile(strTestFile,2,True)
		'objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
	'	objFile_Testhtml.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Step.No</font></th><th BGCOLOR='879773'><font size=2 >Step Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Message</font></th><th BGCOLOR='879773'><font size=2>Status</font></th><th BGCOLOR='879773'><font size=2>Screenshot</font></th></tr>"
		
		objFile_Testhtml.writeline "<head><style>body {    background-color: #E0FFFF;}</style></head>"
		'objFile_Testhtml.writeline "<h4 align='center'><u>Business Unit Test Report</u></h4>"
		objFile_Testhtml.writeline "<h4 align='center'><u>Automation Test Steps</u></h4>"
		
		objFile_Testhtml.writeline "<br><br>"
		objFile_Testhtml.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>No.</font></th><th BGCOLOR='879773'><font size=2 >Step Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Message</font></th><th BGCOLOR='879773'><font size=2>Status</font></th><th BGCOLOR='879773'><font size=2>Screenshot</font></th></tr>"
		
'	else
'	
'		StepCounter = 1
'		strTestFile = strPOSResultsPath & "\" & strTC & ".html"
'		
'		set objFile_Testhtml= fsObject.OpenTextFile(strTestFile,2,True)
'		'objFile_html.writeline "<h4 align='center'><u>Automation Report</u></h4>"
'		objFile_Testhtml.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>Step.No</font></th><th BGCOLOR='879773'><font size=2 >Step Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Message</font></th><th BGCOLOR='879773'><font size=2>Status</font></th><th BGCOLOR='879773'><font size=2>Screenshot</font></th></tr>"
'	End if 
End Function
'******************************************** HEADER ******************************************
' Name : Append_TestHTML
' Description : Generic function to append the step description and status to each test case specific 
'				html report
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : iStep,sStepName,sStepDescrption,sStatus
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function Append_TestHTML(iStep,sStepName,sStepDescrption,sStatus)

	Dim sScreenshot
	
	If sStatus = "PASSED" Then
	    sScreenshot = TakeScreenshot
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sStepName & "</a></td><td align=""center""><font color=""blue"">"&sStepDescrption&"</td></font><td BGCOLOR=""#3ADF00"">PASSED</td><td align=""center""><a href=""" & sScreenshot & """>Screenshot</a></td></tr>"
	else
		sScreenshot = TakeScreenshot
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sStepName & "</a></td><td align=""center""><font color=""blue"">"&sStepDescrption&"</td></font><td BGCOLOR=""#FF0000"">FAILED</td><td align=""center""><font color=""blue""><a href=""" & sScreenshot & """>Screenshot</a></td></tr>"
	End If	
	
	StepCounter = StepCounter + 1
	
End Function

Public Function Append_TestHTMLWithNoMultipleScreenshot(iStep,sStepName,sStepDescrption,sStatus)

	If sStatus = "PASSED" Then
	    
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sStepName & "</a></td><td align=""center"" width=500><font color=""blue"">"&sStepDescrption&"</td></font><td BGCOLOR=""#3ADF00"">PASSED</td><td></td></tr>"
	else
		
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sStepName & "</a></td><td align=""center"" width=500><font color=""blue"">"&sStepDescrption&"</td></font><td BGCOLOR=""#FF0000"">FAILED</td><td></td></tr>"
	End If	
	
	StepCounter = StepCounter + 1
	
End Function

Public Function Append_TestHTMLWithOutSS(iStep,sStepName,sStepDescrption,sStatus)

	
	If sStatus = "PASSED" Then
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sStepName & "</a></td><td align=""center""><font color=""blue"">"&sStepDescrption&"</td></font><td BGCOLOR=""#3ADF00"">PASSED</td><td align=""center""></td></tr>"
	else
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sStepName & "</a></td><td align=""center""><font color=""blue"">"&sStepDescrption&"</td></font><td BGCOLOR=""#FF0000"">FAILED</td><td align=""center""><font color=""blue""></td></tr>"
	End If	
	
	StepCounter = StepCounter + 1
	
End Function


Public Function Blank_HTML()
	objFile_Testhtml.writeline "<tr><td></td><td></td><td></td><td></td><td></td></tr>"
End Function
'******************************************** HEADER ******************************************
' Name : DeInitialize_TestHTML
' Description : Generic function to to deinitialize the test case specific HTML file
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function DeInitialize_TestHTML()
	
	objFile_Testhtml.writeline "</table>" 	
	objFile_Testhtml.close
	
	StepCounter = 1
	
	'set objFile_Testhtml= nothing

End Function
'******************************************** HEADER ******************************************
' Name : Total_Execution_Time
' Description : Generic function to calculate the total execution time
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : Start_Time,End_Time
' Output Parameter : TotalTime
'******************************************** HEADER ******************************************
Function Total_Execution_Time(Start_Time,End_Time)

    'total number of seconds
    TotalTime_Secs  = Datediff("s",Start_Time,End_Time)

    'convert  total  Seconds into "Seconds only/ Mins+Secs/ Hrs+Mins+Secs"
    If TotalTime_Secs < 60 Then
        TotalTime = "Total Time Taken For Complete Execution = " & TotalTime_Secs  & " Second(s) Approx."
    ElseIf    TotalTime_Secs >=60 and TotalTime_Secs < 3600 Then
        TotalTime = "Total Time Taken For Complete Execution = " & int(TotalTime_Secs/60) & " Minute(s) and "& TotalTime_Secs Mod 60 & " Second(s) Approx."
    ElseIf    TotalTime_Secs >= 3600 Then
        TotalTime = "Total Time Taken For Complete Execution = " & int(TotalTime_Secs/3600) & " Hour(s) , " & int((TotalTime_Secs Mod 3600)/60)  & " Minute(s) and "& ((TotalTime_Secs Mod 3600 ) Mod 60) & " Second(s)  Approx."
    End If

    'Return the Message 
    Total_Execution_Time = TotalTime
End Function
'******************************************** HEADER ******************************************
' Name : DateTimeRandomNumber
' Description : Generic function to generate the random number with the format ddmmyyyyhhmmss
' Creator : Syed Shafi
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : ddmmyyyyhhmmss
'******************************************** HEADER ******************************************
Public Function DateTimeRandomNumber()
	
	Dim sDate : sDate = Day(Now)
	Dim sMonth : sMonth = Month(Now)
	Dim sYear : sYear = year(Now)
	Dim sHour : sHour = hour(Now)
	Dim sMinute : sMinute = Minute(Now)
	Dim sSecond : sSecond = second(Now)
	
	DateTimeRandomNumber = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)
	
End Function

	
'******************************************** HEADER ******************************************
' Name : Add_ObjectRepository
' Description : generic function to Associate the object Repository.
' Creator : Madhusmitta Pal
' Date : 15th Sep,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : sObjRepName
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
 Function Add_ObjectRepository(sObjRepName)
	
	Dim oObjPath
	
	On error resume next
	
	oObjPath=ORFolderPath&sObjRepName&".tsr"
	
    aAction = Environment.Value("ActionName")
		'''msgbox aAction
	Set qtpApp = CreateObject("QuickTest.Application")
	Set qtpRes =qtpApp.Test.Actions(aAction).ObjectRepositories
	qtpRes.Add oObjPath
	
	Reporter.ReportEvent micDone," Object repository adding to resources","OR Path :"&oObjPath
	Reporter.ReportEvent micDone," Object repository adding to resources ERROR (if any) : ","ERROR (if any)  :"& err.description
	
	Set qtpApp = Nothing
	Set qtpRes = Nothing
	
End Function	
		
'******************************************** HEADER ******************************************
' Name : VerifyImageCheckPoint(sobjname)
' Description : generic function to Click on the Button available inside WinList
' Creator : Madhusmitta Pal
' Date : 8th Sep,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : btnname
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
 Function VerifyImageCheckPoint(dDescObject,sobjname,checkObjName)
	
	Set DialObj = CreateDescriptionObject ("Wnd_N_POS")
	Set WinObj = CreateDescriptionObjectV2(dDescObject,"text:"&sobjname)
	
	Dialog(DialObj).WinButton(WinObj).highlight
	Dialog(DialObj).WinButton(WinObj).Check CheckPoint(checkObjName)
	'''msgbox Dialog(DialObj).WinButton(WinObj).Exist(1)
	If Dialog(DialObj).WinButton(WinObj).Exist(1)= True Then
		Reporter.ReportEvent micPass," Object is visible as Highlighted","Object Name :"&dDescObject
		 bflag = True
		else
		Reporter.ReportEvent micFail," Object is visible as Highlighted","Object Name :"&dDescObject
		bflag = False
	End If
	VerifyImageCheckPoint=bflag
	
	Set DialObj = Nothing
	Set WinObj = Nothing
End Function	

'******************************************** HEADER ******************************************
' Name : getdate(sdate)
' Description : generic function to Click on the Button available inside WinList
' Creator : Madhusmitta Pal
' Date :21st Sep,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : 
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
  Function getdate(sdate)
    sy=year(sdate)
    sm=Month(sdate)
      If len(sm)=1 Then
      	sm="0"&sm
      End If
     
    sd=day(sdate)
      If len(sd)=1 Then
      	sd="0"&sd
      End If
    ddate= sy&sm&sd
   ' ''msgbox m
      getdate=ddate
 End Function


Public Function TakeScreenshot()
	On Error Resume Next

	ScreenShotName = "Error" &  DateTimeRandomNumber & ".png"
	sScreenShotpng = ScreenShotName
	sScreenshotName =ScreenshotPath & ScreenShotName
	Desktop.CaptureBitmap sScreenshotName,True

	'''msgbox "TakeScreenshot - " & sScreenshotName
'	TakeScreenshot = sScreenshotName
	TakeScreenshot = sScreenShotpng
	
 End Function


 
 
 Public Function TakeScreenshot1()

	On Error Resume Next

	ScreenShotName = "Error" &  DateTimeRandomNumber & ".png"
	sScreenshotName =ScreenshotPath & ScreenShotName
	Desktop.CaptureBitmap sScreenshotName,True
	
	'''msgbox "TakeScreenshot - " & sScreenshotName
	TakeScreenshot = sScreenshotName

  
 End Function
 

'******************************************** HEADER ******************************************
' Name : SetRecoveryVariable
' Description : Value to set the Recovery Value
' Creator : Madhusmitta Pal
' Date : 18th Nov,2016
' Last Modified On : 06th Sep,2016
' Last Modified By : Madhusmitta Pal
' Input Parameter : objectName,propertyname,expval
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
Public Function SetRecoveryVariable(sValue)
   
     sRecoveryData = sValue
     
     SetRecoveryVariable = sRecoveryData
End Function
 
 
 '******************************************** HEADER ******************************************
' Name : DateTimeRandomNumber
' Description : Generic function to generate the random number with the format ddmmyyyyhhmmss
' Creator : Madhusmitta Pal
' Date : 26th Aug,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : None
' Output Parameter : ddmmyyyy_Random Number
'******************************************** HEADER ******************************************
Public Function DateRandomNumber()
	
	Dim sDate : sDate = Day(Now)
	Dim sMonth : sMonth = Month(Now)
	Dim sYear : sYear = year(Now)
	Dim sHr : sHr = Hour(Now)
	Dim sMin : sMin = Minute(Now)
	Dim sSec : sSec = Second(Now)
	sRandNo = Int((10000 - 1 + 1) * Rnd + 1)
      
	DateRandomNumber = Int(sDate & sMonth & sYear & sHour  & sHr & sMin & sSec)
	
End Function
 '******************************************** HEADER ******************************************
' Name : To Write the field to validate in receipt
' Description : generic function to click on the numeric keys (0-9)
' Creator : Madhusmitta Pal
' Date : 2nd Dec,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : ObjectName
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function WriteReceiptValidation(sTCName,sContent)
   ' ''msgbox ReceiptValidationPath
	sReceiptFileName = ReceiptValidationPath&sTCName&".txt"
	'''msgbox sReceiptFileName
	Set FileSysObj= CreateObject("Scripting.FileSystemObject")
	
	  If FileSysObj.FileExists(sReceiptFileName) Then
	  	Set objFile = FileSysObj.OpenTextFile(sReceiptFileName,8,True)
	  	else
	  	 Set objFile = FileSysObj.CreateTextFile(sReceiptFileName,True)
	  End If
	
	sCancelFileName = sReceiptFileName
	
	objFile.Write sContent
	objFile.Close
	
	Set objFile= Nothing
    Set FileSysObj = Nothing	
End Function


 '******************************************** HEADER ******************************************
' Name : To Write the String is AlphaNumeric
' Description : generic function to click on the numeric keys (0-9)
' Creator : Madhusmitta Pal
' Date : 24th Jan,2016
' Last Modified On : 
' Last Modified By : 
' Input Parameter : ObjectName
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function isAlphaNum(strString)
	
	Dim IsAlpha 
	 If Regex.IsMatch(strString, "^[a-zA-Z0-9]+$") Then
	 	IsAlpha = True
	   Else
        IsAlpha = False	   
	 End If
	 isAlphaNum= IsAlpha
End Function


'******************************************** HEADER ******************************************
' Description : The function for Opening The Bos URL
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : strBosURL
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function OpenBosURL(strBosURL)

	SystemUtil.Run "iexplore",strBosURL
	
End Function


 '******************************************** HEADER ******************************************
' Description : 
' Creator :  
' Date : 
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function VerifyWebObjectExist(sBrowser,sPage,sObject)
 
  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bPage = CreateDescriptionObject(sPage)
  
   Set bObject = CreateDescriptionObject(sObject)
 
' 	Set sDesObject=Description.Create
' sDesObject("name").value = "Login"
 
' ''msgbox Browser(bBrowser).Exist()
 ' Browser(bBrowser).highlight
  
  ' sStatus= Browser(bBrowser).Page(bPage).Exist(1)
   '''msgbox sStatus
 'Browser(bBrowser).Page(bPage).highlight
   
 set sochild = Browser(bBrowser).Page(bPage).ChildObjects(bObject)
 '''msgbox sLinkochild.count
 
 If  sochild.count<>0 Then
 
      '''msgbox sochild(0).GetRoProperty("name")
     ' ''msgbox sochild(0).Exist()
   ' If sochild(0).Exist() = True Then
    	
       VerifyWebObjectExist = True
    'End If
 
 End If
 End Function
 
 '******************************************** HEADER ******************************************
' Description : The function to Set the data in Edit field within multiple frames
' Creator : 
' Date : 
' Last Modified On :
' Last Modified By : 
' Input Parameter : sURL
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function SetDatainBosEditFrame(sBrowser,sPage,sFrame,sEdit,sData)
  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bPage = CreateDescriptionObject(sPage)
  Set bFrame = CreateDescriptionObject(sFrame)
  Set bEdit = CreateDescriptionObject(sEdit)
	
 set sFramechild = Browser(bBrowser).Page(bPage).ChildObjects(bFrame)
 '''msgbox sEditchild.count
 
 For i  = 0 To sFramechild.count-1
 	'If  sEditchild(i).getRoProperty("name")="WebEdit" and sEditchild(i).getRoProperty("html id")="wluSupplier" Then
	   ' sFramechild(i).highlight
	 	'sEditchild(0).Set sData
	 	if sFramechild(i).getRoProperty("name")<>"" then
	 	     'Set bField = Description.Create
              '  bField("name").Value ="WebEdit"
               ' bField("html id").Value ="wluSupplier"
             'Select Case Operation
                 'sFramechild(i).highlight
                  Set objchild = sFramechild(i).ChildObjects(bEdit)
                  
              ' Case "TypeSet"
               
			 	        ' ''msgbox objchild.count
			 	          If  objchild.count<>0 and objchild.count=1 Then
		                      'objchild(0).Highlight
		 	                  objchild(0).Set sData
		 	                  
		 	                ElseIf objchild.count>1 Then 
                                 For j = 1 To objchild.count-1
                                   '   objchild(j).Highlight
		 	                          objchild(j).Set sData
                                 	
                               Next		 	                
		 	
		                   End If
		                   
		        'Case "TypeClick"  
                          
                 
             ' End Select      
        End If
 Next
End Function

'******************************************** HEADER ******************************************
' Description : Function to set data to webedit under dialog
' Creator : Syed Shafi, Hemaraja 
' Date : 12th September,2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function SetWebEditDialogPage(sBrowser,sDialog,sPage,sObject,sOperation,strData)

	On error resume next
	
	
	If sPage = "NoPage" Then 
		Set bBrowser = CreateDescriptionObject(sBrowser)
		Set bDialog = CreateDescriptionObject(sDialog)		
		Set bObject = CreateDescriptionObject(sObject) 	
		set ochild = Browser(bBrowser).Page(bDialog).ChildObjects(bObject)
	Else 
		Set bBrowser = CreateDescriptionObject(sBrowser)
	    Set bDialog = CreateDescriptionObject(sDialog)
	    Set bPage = CreateDescriptionObject(sPage) 
	    Set bObject = CreateDescriptionObject(sObject)	
		set ochild = Browser(bBrowser).Dialog(bDialog).Page(bPage).ChildObjects(bObject) 	 
	End If
	
	Set sochild= ochild(0)
	
	Select Case sOperation
	
	Case "Submit"
  		sochild.Click		
  		wait 1
  		Set WshShell = CreateObject("WScript.Shell")
  		wshshell.SendKeys strData
  		wait 1
  		wshshell.SendKeys ("{ENTER}")
  
  
  	Case "Click"
  	
	  	If  sochild.Count<>0 Then
			sochild.Click			
		End If
		
	Case "Set"	
		
		If sochild.Count<>0 Then
			
			sochild.Set strData
						
		End If  

 End Select 
 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"SetWebEditDialogPage:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage	
	 	writelog "","","","","SetWebEditDialogPage(): SetWebEditDialogPage is unsucessfull! Object Name : "&sObject & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","SetWebEditDialogPage(): SetWebEditDialogPage is Successful. Object Name : "&sObject
	 End If
 
	
End Function


'******************************************** HEADER ******************************************
' Description : Function to set data to webedit under dialog
' Creator : Hemaraja 
' Date : 12th September,2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function SetWebEditDialogPage1(sBrowser,sDialog,sPage,sObject,sOperation,strData)

	On error resume next
	
	
		Set bBrowser = CreateDescriptionObject(sBrowser)
	    Set bDialog = CreateDescriptionObject(sDialog)
	    Set bPage = CreateDescriptionObject(sPage) 
	    Set bObject = CreateDescriptionObject(sObject)
	    set ochild = Browser(bBrowser).Dialog(bDialog).Page(bPage).ChildObjects(bObject)
		Set sochild= ochild(0)
	Select Case sOperation
	
	Case "Submit"
  		sochild.Click		
  		wait 1
  		Set WshShell = CreateObject("WScript.Shell")
  		wshshell.SendKeys strData
  		wait 1
  		wshshell.SendKeys ("{ENTER}")
  
  
  	Case "Click"
  	
	  	If  sochild.Count<>0 Then
			sochild.Click			
		End If
		
	Case "Set"	
		
		If sochild.Count<>0 Then
			
			sochild.Set strData
						
		End If  

 End Select 
 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"SetWebEditDialogPage:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage	
	 	writelog "","","","","SetWebEditDialogPage(): SetWebEditDialogPage is unsucessfull! Object Name : "&sObject & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","SetWebEditDialogPage(): SetWebEditDialogPage is Successful. Object Name : "&sObject
	 End If
 
	
End Function

Public Function ClickWinDialog(sWindow,sDialog,sButton)

	  Set bWindow = CreateDescriptionObject(sWindow)
      Set bDialog = CreateDescriptionObject(sDialog)
  
      Set bButton = CreateDescriptionObject(sButton)
	
	  Window(bWindow).Dialog(bDialog).WinButton(bButton).Click	  
	
End Function


'******************************************** HEADER ******************************************
' Description : The function to operate on object
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function OperateOnObject(sBrowser,sPage,sFrame,sObject,sOperation,strData)
	
	On error resume next	
	
'	Set sochild=GetFrameObject(sBrowser,sPage,sFrame,sObject)	
	
	WaitForObject sBrowser,sPage,sFrame,sObject,5
  
	If sFrame = "NoFrame" Then 
		Set bBrowser = CreateDescriptionObject(sBrowser)
		Set bPage = CreateDescriptionObject(sPage)		
		Set bObject = CreateDescriptionObject(sObject) 	
		set ochild = Browser(bBrowser).Page(bPage).ChildObjects(bObject)
		Set sochild= ochild(0)
	Else  	 
		Set sochild=GetFrameObject(sBrowser,sPage,sFrame,sObject)  	 
	End If
	
	If not(isobject(sochild)) Then
		OperateOnObject = False
		Exit function
	End If
  
 Select Case sOperation
  
  	Case "Submit"
  		sochild.Click		
  		wait 1
  		Set WshShell = CreateObject("WScript.Shell")
  		wshshell.SendKeys strData
  		wait 1
  		wshshell.SendKeys ("{ENTER}")
  
  	Case "Click"
  	
	  	If  sochild.Count<>0 Then	  	
	  	
'			If strData<>"" Then
'			sochild(0).Set strData 
'			Else
			sochild.Click		
'	  		End If
				
		End If
		
	Case "Set"
	
		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count<>0 Then
			sochild.Click
			wait 1
			sochild.Set strData
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1
			sochild.Click			
		End If  
		
	Case "Set1"
	
'		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count<>0 Then
			sochild.Click
			wait 1
			sochild.Set strData
		End If  
		
	Case "Select"
  		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count<>0 Then
			sochild.Click
			wait 1
			sochild.Set strData
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1
			sochild.Click
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1
		End If 
	
	Case "Select1"
  		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count <> 0 Then
			sochild.Click
			wait 1
			WshShell.SendKeys strData
			wait 1
			WshShell.SendKeys "{ENTER}"
		End If 
	
	Case "Select2"
  		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count <> 0 Then
			sochild.Click
			wait 1
			WshShell.SendKeys strData
			wait 1
			WshShell.SendKeys "{ENTER}"
			wait 1
			WshShell.SendKeys "{ENTER}"
		End If 
	
	Case "SelectAny"
		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count<>0 Then
			sochild.Click
			wait 1
			WshShell.SendKeys "{DOWN}"
			wait 1
			WshShell.SendKeys "{ENTER}"
		End If

	Case "Highlight"
  	
	  	If  sochild.Count<>0 Then	  	
	  	
			sochild.Highlight		
				
		End If
		
	Case "Exist"
	
		If sochild.exist(5) and sochild.GetROProperty("height") <> 0  Then
			OperateOnObject = "True"
		End If
	
	Case "RadioSelect"
	
		If sochild.Count<>0 Then
			sochild.Select strData
		End If
		
    Case "GetROProperty"
  	
	  	If  sochild.Count<>0 Then	  	
	  	
            sochild.Highlight
			sValuepro= sochild.GetROProperty(strData)		
            OperateOnObject = sValuepro
				
		End If
	Case "Mover"
	
	If sochild.Count<>0 Then
		Setting.webPackage("ReplayType")=2
		wait 2
		sochild.FireEvent "onmouseover"
		wait 2
		Setting.webPackage("ReplayType")=1
		wait 2
	End If
 End Select 
 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"OperateOnObject:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage	
	 	writelog "","","","","OperateOnObject(): Operate on object is unsucessfull! Object Name : "&sObject & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","OperateOnObject(): Operate on object is Successful. Object Name : "&sObject
	 End If
 
'Browser("Site Manager | 9123 |").Page("Site Manager | 9123 |").Frame("Frame4").WebTable("!Auto_PO").GetCellData


End Function


'******************************************** HEADER ******************************************
' Description : The function to get object property value
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function GetObjPropertyValue(sBrowser,sPage,sFrame,sObject,PropertyName)
	
	On error resume next
	
  If sFrame = "NoFrame" Then 
		Set bBrowser = CreateDescriptionObject(sBrowser)
		Set bPage = CreateDescriptionObject(sPage)		
		Set bObject = CreateDescriptionObject(sObject) 	
		set ochild = Browser(bBrowser).Page(bPage).ChildObjects(bObject)
		Set sochild= ochild(0)
	Else  	 
		Set sochild=GetFrameObject(sBrowser,sPage,sFrame,sObject)  	 
	End If
  
  If sochild.count<>0 Then
  	 GetObjPropertyValue=sochild.GetROProperty(PropertyName)
  End If  
  
   ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"OperateOnObject:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	 End If
  	
End Function



'******************************************** HEADER ******************************************
' Description : The function to wait for object
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function WaitForObject(sBrowser,sPage,sFrame,sObject,strTimeoutValue)
	On error resume next

If strTimeoutValue > 20 Then
	strTimeoutValue = 20
End If

		bExistFlag = False 

	
	Dim bExistFlag

	Set bBrowser = CreateDescriptionObject(sBrowser)
		Set bPage = CreateDescriptionObject(sPage)		
		Set bObject = CreateDescriptionObject(sObject)

	For Iterator = 1 To strTimeoutValue Step 1
	If sFrame = "NoFrame" or sFrame = "" or sFrame = "No Frame" Then 
		 	

		set ochild = Browser(bBrowser).Page(bPage).ChildObjects(bObject)
		If ochild.count <> 0 Then
		If sObject = "webbutton_instructions" Then
	If 	ochild.getroproperty("visible") Then
			bExistFlag=true
Exit for
else
wait 1
End If
else
		bExistFlag=true
Exit for


End If
'		Set sochild= ochild(0)	
'		bExistFlag=true
'Exit for	
Else 
wait 1
		End If

			
	Else  	 
		Set ochild=GetFrameObject(sBrowser,sPage,sFrame,sObject)  	 
		If ochild.count <> 0 Then
		bExistFlag=true
Exit for	
else
wait 1
		End If

		
	End If
		next
	set ochild=nothing
	'sochild.highlight
'	Set objKeyObject = sochild
	
'	Do	
'		If objKeyObject.exist(2) = true Then
''''msgbox objKeyObject.getroproperty("visible")
''			objKeyObject.Highlight
'			'waiting until object exist
'			bExistFlag = True
'			
'		End If
'		
'		If TimeCounter>strTimeoutValue Then
'			Exit do
Iterator=empty
	Set bBrowser = nothing
		Set bPage = nothing
		Set bObject=nothing
'		TimeCounter =TimeCounter + 1
'		
'	Loop While bExistFlag = False 
	
	ErrNumber = Err.Number
	ErrMessage = Err.Description
	
	If ErrNumber<>0 Then
		WriteLog "","","","","WaitForObjec(): Error has occured! "," Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	Else
		WriteLog "","","","","WaitForObjec(): Successfull for object: "&sObject
	End If
'	End if
	WaitForObject=bExistFlag
	 
End Function

'******************************************** HEADER ******************************************
' Description : The function to wait for Dialog object
' Creator :  Hemaraja 
' Date : 6th September, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function WaitForDialogObject(sWindow,sDialog,sObject,strTimeoutValue)

	On error resume next
	
	Dim bExistFlag
	Dim TimeCounter 
	
	bExistFlag = False 
	TimeCounter=0
	
	Set bWindow = CreateDescriptionObject(sWindow)
    Set bDialog = CreateDescriptionObject(sDialog)  
    Set bButton = CreateDescriptionObject(sObject)
    If instr(1,sWindow,"Browser")>0 Then
    	set sochild = Browser(bWindow).Dialog(bDialog).WinButton(bButton)
    Else
    	set sochild = Window(bWindow).Dialog(bDialog).WinButton(bButton)    	
    End If	

	'sochild.highlight
	Set objKeyObject = sochild
	
	Do	
		If objKeyObject.Exist(2) = true Then
'			objKeyObject.Highlight
			'waiting until object exist
			bExistFlag = True
			
		End If
		
		If TimeCounter>strTimeoutValue Then
			Exit do
		End If
		TimeCounter =TimeCounter + 1
		
	Loop While bExistFlag = False 
	
	ErrNumber = Err.Number
	ErrMessage = Err.Description
	
	If ErrNumber<>0 Then
		WriteLog "","","","","WaitForDialogObject(): Error has occured! "," Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	Else
		WriteLog "","","","","WaitForDialogObject(): Successfull for object: "&sObject
	End If
	
	WaitForObject=bExistFlag
	 
End Function

'******************************************** HEADER ******************************************
' Description : The function to wait until object to disappear
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Function WaitUntilObjectDisappear(sBrowser,sPage,sFrame,sObject,strTimeoutValue)
	
	Do
	
		bFlag = WaitForObject(sBrowser,sPage,sFrame,sObject,strTimeoutValue)
	
	Loop while (bFlag=True)
	
	ErrNumber = Err.Number
	ErrMessage = Err.Description
	
	If ErrNumber<>0 Then
		WriteLog "","","","","WaitUntilObjectDisappear(): Error has occured! "," Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	Else
		WriteLog "","","","","WaitUntilObjectDisappear(): Successfull for object disappeared: "&sObject
	End If
	
	
End Function



'******************************************** HEADER ******************************************
' Description : The function for validating the webtable cell value
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateWebTableCellValueSMR(sBrowser,sPage,sFrame,sWebTable,strCellSearchValue)

  On error resume next
  
  Dim rCount
  Dim cCount
  Dim row
  Dim col

  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bPage = CreateDescriptionObject(sPage)
  Set bFrame = CreateDescriptionObject(sFrame)
  Set bWebTable = CreateDescriptionObject(sWebTable)
  
  Set objchild = Browser(bBrowser).Page(bPage).Frame(bFrame).ChildObjects(bWebTable)
  
  If objchild.count<>0 Then
  	 objchild(0).Highlight
  	 
  	 rCount=objchild(0).RowCount
  	   	 
  	 For row = 3 To rCount Step 1
  	 	
  	 	cCount=objchild(0).Columncount(row)  	 	
  	 	 	 	
  	 	For col = 1 To cCount Step 1 	  	 		
  	 			
  	 		If objchild(0).GetCellData(row,col)=strCellSearchValue Then	  	 		
  	 		  	 	 		  	
  	 			ValidateWebTableCellValue=True 
  	 			
  	 			Exit for 
  	 			
  	 		End If 	 		
  	 		
  	 	Next
  	 	
  	 	If ValidateWebTableCellValue=True Then
  	 	
  	 		Exit for
  	 		
  	 	End If
  	 	
  	 Next
  	 
  End If
  
   	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"OperateOnObject:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	 End If

	
	
End Function

'******************************************** HEADER ******************************************
' Description : The function for validating the data of the webtable columnn value
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateDtWtColValue(sBrowser,sPage,sFrame,sObject,byref DictTbl)
	
  On error resume next  

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  LastRow=2
	
'  set TableObject=FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count
'  If objchild.count<>0 AND totalDictNum<>0 Then
  	 
	 Do
'	 	If DictTbl.Count<>0 Then
	 	
		 	For Each elem in DictTbl
		 			counter = counter+1
				 	TblColName=elem
				 	TbleColValue=DictTbl(elem)
				 	
				 	rowNum=TableObject.GetRowWithCellText(TblColName)
				 	colNumb=GetWTColumnNumber(TableObject,TblColName) 		
				 	
				 	rCount=TableObject.RowCount		
			
					If DNLoopRowflag=false Then
						'If Trim(TableObject.GetCellData(rowNum,colNumb))=Trim(TblColName) Then
					
						For row = LastRow To rCount Step 1
							strCellValue = TableObject.GetCellData(row,colNumb)
							
							Select Case blgCompFlag
								
								Case "INSTRFLAG"						 
									
									If instr(trim(strCellValue),TbleColValue)>0 Then
									
										'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue) ,"PASSED")
										AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"	
										
									   'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue) ,"PASSED"
									   writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 		   									   						   
									   LastRow=row
									   TblrowNum=row
									   DNLoopRowflag=True
									   Exit For
									   
									End If
									
								Case "STRCMPFLAG"
								
									If trim(strCellValue) = TbleColValue Then
										
										'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue) ,"PASSED")
										AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"
									  ' Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue) ,"PASSED"
									   writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue)
									   LastRow=row
									   TblrowNum=row
									   DNLoopRowflag=True
									   Exit For
									   
									End If
								
								
							End Select
							
							
						Next
						
						
					'End If	
					else
						strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
						
						Select Case blgCompFlag
							
							Case "INSTRFLAG"				
							
							If instr(trim(strCellValue1),TbleColValue)>0 Then
								'Call HTMLTableMessage("Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
							 'Append_TestHTMLWithOutSS StepCounter,"Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED"
							 
								writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue1) 														
								bMatchedRow = True
								else
								'Call HTMLTableMessage("Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED"
								writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue1) 
								bMatchedRow = False
								DNLoopRowflag=False
								RefreshDicObj=True
								TblrowNum=0
								LastRow=LastRow+1
								counter=0
								Exit for
							End if
							
							Case "STRCMPFLAG"
							
							If trim(strCellValue1) = TbleColValue Then
								'Call HTMLTableMessage("Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED"
								writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue1) 
								bMatchedRow = True
								else
								'Call HTMLTableMessage("Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value"&elem,"Expected Value : " & TbleColValue & VBCRLF &  " Actual Value: " & trim(strCellValue1) ,"PASSED"
								writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue1) 
								bMatchedRow = False
								DNLoopRowflag=False
								RefreshDicObj=True
								TblrowNum=0
								LastRow=LastRow+1
								counter=0
								Exit for
							End if
							
						End Select
						
										
					
			 		End If
		 		
		 	Next
	 		
	 		If bMatchedRow =true Then
	 			RefreshDicObj=False
	 		End If
	 		
	 		
'	 	End If 
	 Loop While (RefreshDicObj=True and counter <=totalDictNum)
  
 	
'  End If 

	If ErrNumber<>0 Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
  
  ValidateDtWtColValue = bMatchedRow
	
End Function



'******************************************** HEADER ******************************************
' Description : The function for getting webtable column number by passing columnname
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
  Public Function GetWTColumnNumber(byref objTableObject,TblColName)
  	
  On error resume next  
  
  
'  set objTableObject=FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)
  
'  Set bBrowser = CreateDescriptionObject(sBrowser)
'  Set bPage = CreateDescriptionObject(sPage)
'  Set bFrame = CreateDescriptionObject(sFrame)
'  Set bWebTable = CreateDescriptionObject(sWebTable)
'  
'  Set objchild = Browser(bBrowser).Page(bPage).Frame(bFrame).ChildObjects(bWebTable)
'  
'  If objchild.count<>0 Then
  	 'objchild(0).Highlight
  	 
  	 cCount=objTableObject.ColumnCount(1)
  	 
  	 For i = 1 To cCount Step 1
  	 	TblGetCellData=objTableObject.GetCellData(1,i)
  	 	If Trim(TblColName)=Trim(TblGetCellData) Then
  	 	   GetWTColumnNumber=i
  	 	   Exit For  	 		
  	 	End If
  	 Next
  	 
'  End If 
  
  	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"GetWTColumnNumber:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	 End If 
  	 
  End Function
  


'******************************************** HEADER ******************************************
' Description : The function for verifying object existence in webtable
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : strBosURL
' Output Parameter : None
'******************************************** HEADER ******************************************  	
Public Function VerifyObjectExistInWebTbl(sBrowser,sPage,sFrame,sWebTable,sObject,PropertyName,PropertyValue)
 
  On error resume next  
  
  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bPage = CreateDescriptionObject(sPage)
  Set bFrame = CreateDescriptionObject(sFrame)
  Set bWebTable = CreateDescriptionObject(sWebTable)
  Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
  
  Set sochild = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable).ChildObjects(bObject)   
 
  If sochild.count<>0 Then 
  	 sochild(0).Highlight  	 
     VerifyObjectExistInWebTbl = True   
 
  End If
  
 ErrNumber = Err.Number
 ErrMessage = Err.Description
	 
 If ErrNumber<>0 Then	 	
 	writelog "","","","","VerifyObjectExistInWebTbl(): VerifyObjectExistInWebTbl is unsucessfull! Object Name : "&PropertyValue & " Error Number : "&err.number & " Error Description : "&Err.description
 else
 	writelog "","","","","VerifyObjectExistInWebTbl(): VerifyObjectExistInWebTbl is Successful. Object Name : "&PropertyValue
 End If
 
 End Function

'******************************************** HEADER ******************************************
' Description : The function for extracting only numbers from string
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ExractOnlyNumFrmStr(strValue)

   On error resume next
	
	NumValue=""
	
	For i = 1 To len(strValue) Step 1
		
		Value=mid(strValue,i,1)
	
		If isnumeric(Value)=True Then
		
		  NumValue=NumValue&Value
			
		End If
		
	Next
	
	ExractOnlyNumFrmStr=NumValue
	
	ErrNumber = Err.Number
    ErrMessage = Err.Description
	 
    If ErrNumber<>0 Then
	   Reporter.ReportEvent micPass,"OperateOnObject:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
    End If 
	
End Function

'******************************************** HEADER ******************************************
' Description : The function for selecting envirnoment
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Function SelectEnvirnoment(ENV_Flag)

	If ENV_Flag ="" Then
     	ENV_Flag=Environment.Value("ENV_Flag")
    End If
	
	Select Case strPublicRole
		
		Case "Site Manager"
			Comm_Browser3="Browser_CH_Test_Home"
			Comm_Page3="Page_CH_Test_Home"
			Comm_Frame="Frame_BOS"
			Comm_Browser4="Browser_Common"
			Comm_Page4="Page_Common"
			Comm_Page3=Comm_Page4
			Comm_Browser3=Comm_Browser4
		Case "L1 Support"
			Comm_Browser3="Browser_CH_TestL1_Home"
			Comm_Page3="Page_CH_TestL1_Home"
			Comm_Frame="Frame_BOS"			
		Case "Manage Retail Data"
			Comm_Browser3="Browser_CH_Test_Home"
			Comm_Page3="Page_CH_Test_Home"
			Comm_Frame="Frame_BOS"
			Comm_Browser4="Browser_Common"
			Comm_Page4="Page_Common"
			Comm_Page3=Comm_Page4
			Comm_Browser3=Comm_Browser4	
			
	End Select

		
	Select Case ENV_Flag
	
		Case "ZARP-TEST"	
		
			ENV_URL = "sww-x64test2.zarp.redprairie.shell.com"
			
			Comm_Browser1="Browser_CH_Test"
			Comm_Page1="Page_CH_Test"
			Comm_Browser2="Browser_ZArp_Test_login"
			Comm_Page2="Page_ZArp_Test_login"
			Comm_Browser4="Browser_Common"
			Comm_Page4="Page_Common"
			
		Case "CH-PREP"
		
			ENV_URL = "https://sww-prep.CHRP.redprairie.shell.com" 			
		
			Comm_Browser_Sec="Bro_SecCertificate"
			Comm_Page_Sec="Page_SecCertificate"
			Comm_WebDialog1="WebDialog_Bos_SecCrt"
			Comm_Browser2="Bro_BOS"
			Comm_Page2="Page_BOS"
		
		Case "CH-TEST"
		
			ENV_URL = "sww-x64test1.chrp.redprairie.shell.com"			
			
			Comm_Browser1="Browser_CH_Test"
			Comm_Page1="Page_CH_Test"
			Comm_Browser2="Browser_CH_Test_login"
			Comm_Page2="Page_CH_Test_login"
			
		Case "ZA-PREP"
		
			ENV_URL = "sww-prep.Zarp.redprairie.shell.com"
			
			Comm_Window1="Window_ZA_Prep"
			Comm_Browser_Sec="Bro_SecCertificate"
			Comm_Page_Sec="Page_SecCertificate"
			Comm_WebDialog1="WebDialog_Bos_SecCrt"
			Comm_Browser1="Browser_CH_Test"
			Comm_Page1="Page_CH_Test"
			Comm_Browser2="Browser_ZA_Prep_login"
			Comm_Page2="Page_ZA_Prep_login"	
		
		Case "UK-RBA-PREP"		
		
			ENV_URL = "sww-prep.UKrba.redprairie.shell.com"
		
		Case "UK-RBA-TEST"
		
			ENV_URL = "sww-x64test1.ukrba.redprairie.shell.com"
			
			Comm_Browser1="Browser_CH_Test"
			Comm_Page1="Page_CH_Test"
			Comm_Browser2="Browser_UKrba_Test_login"
			Comm_Page2="Page_UKrba_Test_login"

		Case "NL-RBA"
		
			ENV_URL = "http://sww-x64test3.nlrba.redprairie.shell.com"
			
			Comm_Browser2="Browser_NL_Login"
			Comm_Page2="Page_NL_Login"				
			
		Case "LUX-RBA"
		
			ENV_URL = "http://sww-x64test3.luxrba.redprairie.shell.com"
			
			Comm_Browser2="Browser_LUX_Login"
			Comm_Page2="Page_LUX_Login"		
		
	End Select
	
	
End Function




'******************************************** HEADER ******************************************
' Description : The function to validate table data with multiple key column values.
' Creator : Niharika
' Last Modified On : 6th january,2019
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function ValidateDtWtMultipleColValue(sBrowser,sPage,sFrame,sObject,byref DictTbl,comparisonCnt,compareColValues,validateKeyColVal)


	On error resume next  
	
	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
	TableObject.Highlight
	counter=0
	totalDictNum = DictTbl.Count
	
  	Set keyColValueDict = CreateObject("Scripting.Dictionary")
  	Set keyColNumDict = CreateObject("Scripting.Dictionary")
  	Set strCellValueDict = CreateObject("Scripting.Dictionary")
  	
  	keyColDict = split(compareColValues,"|")
  	
  	For i = 0 To comparisonCnt-1 Step 1
  		key = keyColDict(i)
  		val = DictTbl.item(key)
  		Call Add_Dictionary(keyColValueDict,key,val)
  	Next
	
	For each key in keyColValueDict
		colNumb=GetWTColumnNumber(TableObject,key) 			 	
		Call Add_Dictionary(keyColNumDict,key,colNumb)
	Next
	rCount=TableObject.RowCount	

	For row = 2 To rCount Step 1
		For each key in keyColValueDict
			colNo = keyColNumDict.item(key)
  			strCellValue = TableObject.GetCellData(row,colNo)			 	
  			Call Add_Dictionary(strCellValueDict,key,strCellValue)
  		Next
		For each key in keyColValueDict
			newStrCellValue = strCellValueDict(key)
			newKeyColValue = keyColValueDict(key)
			If instr(1,newStrCellValue,newKeyColValue) <> 0 Then
				c = c+1
			ElseIf instr(newKeyColValue,">")<>0 Then
					If newStrCellValue>0 Then
						c = c + 1
					End If
			End If
			
			If c = comparisonCnt Then
				reqRowNum = row
				Exit for
			End If
		Next
		If c = comparisonCnt Then
			Exit for
		End If
		c = 0
	Next
  
	For Each elem in DictTbl
	
		TblColName=elem
		TbleColValue=DictTbl(elem)
		actualReqRowNum = reqRowNum
		if instr(TblColName,"Tank")<>0 Then
			reqRowNum = reqRowNum + 1
		End if 
		if validateKeyColVal = "Yes" then
			diffFlag = True
		ElseIf validateKeyColVal = "No" then
			If keyColValueDict.exists(elem)=False Then
				diffFlag = True
			ElseIf keyColValueDict.exists(elem)=True Then
				diffFlag = False			
			End If
		End if 
		
		If diffFlag = True Then
			If instr(TbleColValue,">") <>0 Then
				numericComparisonFlag = True
			End If
				 	
			colNumb=GetWTColumnNumber(TableObject,TblColName) 		
			strCellValue = TableObject.GetCellData(reqRowNum,colNumb)
			If numericComparisonFlag = true Then
				If strCellValue > 0 Then
					bMatchedRow = True
				else
					bMatchedRow = False
				End if 
			ElseIf numericComparisonFlag = false Then
				'--------------------Added newly-----------------------------------------------
				If instr(strCellValue,"RM")<>0 or isNumeric(strCellValue) = True Then
					valType = "numType"
				ElseIf instr(strCellValue,"%")<>0 Then
					tempVal = replace(strCellValue,"%","")
					If isNumeric(tempVal) = True Then
						valType = "numType"
					else
						valType = "strType"
					End If
				else
					valType = "strType"
				End If
				
				If valType = "numType" Then
					strCellValue = ChangeNumberFormat(strCellValue)
					TbleColValue = ChangeNumberFormat(TbleColValue)
					If cdbl(strCellValue)=cdbl(TbleColValue) Then
						bMatchedRow = True
					else
						bMatchedRow = false
					End If
				ElseIf valType = "strType" Then
					If instr(strCellValue,TbleColValue)<>0 Then
						bMatchedRow = True
					else
						bMatchedRow = false
					End If
				End If
				'-----------------Added newly end---------------------------------------
			End if 
			
		

			If bMatchedRow = True Then
				'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue),"PASSED")
				AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue)&"</B>","PASSED"	
				'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue),"PASSED"
				writelog "","","","","Report(): Column Value matched. Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue)
				bMatchedRow = True
			else
				'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue),"FAILED")
				AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue)&"</B>","FAILED"	
				'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue),"FAILED"
				writelog "","","","","Report(): Column Value not matched!! Expected Value : " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue)
				bMatchedRow=false
				If ErrNumber<>0 Then	 	
					writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
				else
					writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
				End If
			End If
		End If
		diffFlag = ""
		reqRowNum = actualReqRowNum
	Next
	
	ValidateDtWtMultipleColValue = bMatchedRow
	
End Function





'******************************************** HEADER ******************************************
' Description : The function for validating the webtable column value based on pump number
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateDtWtColValueBasedOnPumpNumber(sBrowser,sPage,sFrame,sObject,PumpNumber,DictTbl)
	
  On error resume next  

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
	
  set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  counter=0
  totalDictNum = DictTbl.Count 

	TableObject.Highlight
  
  	FirstPumpRowNo=GetExactPumpRowNumber(TableObject,PumpNumber,"")
  	
  	strNextPump=cstr(cint(PumpNumber)+1)
  	LastPumpRowNo=GetExactPumpRowNumber(TableObject,strNextPump,FirstPumpRowNo) 
  	
  	LastRow=FirstPumpRowNo+1
  	rCount=LastPumpRowNo-2
  	 
	 Do
		 	For Each elem in DictTbl
		 			counter = counter+1
				 	TblColName=elem
				 	TbleColValue=DictTbl(elem)
				 	
				 	rowNum=TableObject.GetRowWithCellText(TblColName)
				 	colNumb=GetWTColumnNumber(TableObject,TblColName) 		
				 	
				 	'rCount=TableObject.RowCount		
			
					If DNLoopRowflag=false Then
											
						For row = LastRow To rCount Step 1
							strCellValue = TableObject.GetCellData(row,colNumb)							
							If instr(1,trim(strCellValue), TbleColValue ) <> 0 Then
								'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"
							  ' Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
							   writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
							   LastRow=row
							   TblrowNum=row
							   DNLoopRowflag=True
							   Exit For
							   
							End If
						Next
									
					
					else
						strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
						strCellValue1 = replace(strCellValue1,",","")						
						If instr(1,trim(strCellValue1), TbleColValue ) <> 0Then
							'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"PASSED")
							AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
							'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"PASSED"
							writelog "","","","","Report(): Column Value matched. Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)
							bMatchedRow = True
						else
							'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"FAILED")
							AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
							'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"FAILED"
							writelog "","","","","Report(): Column Value not matched!! Expected Value : " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)
							bMatchedRow = False
							DNLoopRowflag=False
							RefreshDicObj=True
							TblrowNum=0
							LastRow=LastRow+1
							counter=0
							Exit for
						End if
					
					
			 		End If
		 		
		 	Next
	 		
	 		If bMatchedRow =true Then
	 			RefreshDicObj=False
	 		End If
	 		
	 		
	 Loop While (RefreshDicObj=True and counter <=totalDictNum)
	 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then	 	
	 	writelog "","","","","ValidateDtWtColValueBasedOnPumpNumber(): ValidateDtWtColValueBasedOnPumpNumber is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","ValidateDtWtColValueBasedOnPumpNumber(): ValidateDtWtColValueBasedOnPumpNumber is Successful. Object Name : "&sWebTable
	 End If
	 
  ValidateDtWtColValueBasedOnPumpNumber = bMatchedRow
	
End Function



'******************************************** HEADER ******************************************
' Description : The function for getting exact pump number
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function GetExactPumpRowNumber(byref objTableObject,PumpNumber,startrow)
  
  On error resume next  
	
	If startrow="" Then
		startrow=1
	End If
	
	objTableObject.Highlight
	
	rCount=objTableObject.RowCount
	
	For row = cint(startrow) To rCount Step 1	  		
		strActual = objTableObject.GetCellData(row,1)
		strActual = trim(strActual)
		
		If strActual=PumpNumber Then
			PumpRow = row
			Exit for
		End If
		
	Next 
	
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then	 	
	 	writelog "","","","","GetExactPumpRowNumber(): GetExactPumpRowNumber is unsucessfull! Object Name : "&PumpNumber & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","GetExactPumpRowNumber(): GetExactPumpRowNumber is Successful. Object Name : "&PumpNumber
	 End If
  
  GetExactPumpRowNumber = PumpRow
	
	
End Function


'******************************************** HEADER ******************************************
' Description : The function for finding the webtable object in a page
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)
	
	Dim TableObject
	Dim strcolumncnt
	
	Set bBrowser = CreateDescriptionObject(sBrowser)
	Set bPage = CreateDescriptionObject(sPage)
	Set bWebTable = CreateDescriptionObject(sWebTable)
	
	If sFrame<>"" Then
		Set bFrame = CreateDescriptionObject(sFrame)
		Set objchild = Browser(bBrowser).Page(bPage).Frame(bFrame).ChildObjects(bWebTable)
	ELSE
		Set objchild = Browser(bBrowser).Page(bPage).ChildObjects(bWebTable)
	End If
		
	For i = 1 To objchild.count
			
			strcolumncnt= objchild(i).ColumnCount(1)
			If strcolumncnt>1 Then
				Set TableObject =objchild(i)
				Exit for
			End If
			
	Next
	
	ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then	 	
	 	writelog "","","","","FindWTObjectInPage(): FindWTObjectInPage is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","FindWTObjectInPage(): FindWTObjectInPage is Successful. Object Name : "&sWebTable
	 End If

	Set FindWTObjectInPage = TableObject

End function


'******************************************** HEADER ******************************************
' Description : The function for finding the frame object
' Creator :  Hemaraja 
' Date : 22nd Auguest, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function GetFrameObject(sBrowser,sPage,sFrame,sObject)
	
	Dim oReturnObject
	
	Set bBrowser = CreateDescriptionObject(sBrowser)
	Set bPage = CreateDescriptionObject(sPage)
	Set bFrame = CreateDescriptionObject(sFrame)	
	Set bObject = CreateDescriptionObject(sObject)

If sObject = "txt_BosRelease" Then
			Set bWindow = CreateDescriptionObject("Window_ItemSetUp1")
			set oFrames = Browser(bBrowser).Window(bWindow).Page(bPage).ChildObjects(bFrame)
	else
		set oFrames = Browser(bBrowser).Page(bPage).ChildObjects(bFrame)
End If




	For i  = 0 To oFrames.count-1
	
	 	if oFrames(i).getRoProperty("name")<>"" then
'			oFrames(i).highlight
			Set oChild = oFrames(i).ChildObjects(bObject)
			If  oChild.count<>0 and oChild.count=1 Then
				'oChild(0).Highlight
				Set oReturnObject = oChild(0)
'				iindex = oReturnObject.GetROProperty("index")
'				''msgbox iindex
				svisible = oReturnObject.GetROProperty("visible")
				'''msgbox svisible
				If svisible = "True" Then
					
					Exit for
				End If
				
			ElseIf oChild.count>1 Then 
				For j = 1 To oChild.count-1
					'oChild(j).Highlight
					Set oReturnObject = oChild(j)
'					iindex = oReturnObject.GetROProperty("index")
'					''msgbox iindex
					svisible = oReturnObject.GetROProperty("visible")
'					''msgbox svisible
					If svisible = "True" Then
						Exit for
					End If
					
				Next		 	                
			End If
		End If
	Next
	
	Set GetFrameObject = oReturnObject
	
End Function


'******************************************** HEADER ******************************************
' Description : The function to click the object in webtable
' Creator :  Hemaraja 
' Date : 19th September, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************  	
Public Function ClickObjectInWebTbl(sBrowser,sWindow,sPage,sFrame,sWebTable,sObject,PropertyName,PropertyValue)
 
  On error resume next  
 
  If   sWindow = "NoWindow" and sFrame= "NoFrame" Then 
		Set bBrowser = CreateDescriptionObject(sBrowser)
'  		Set bWindow = CreateDescriptionObject(sWindow)
  		Set bPage = CreateDescriptionObject(sPage)
  		Set bWebTable = CreateDescriptionObject(sWebTable)
  		Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
		Set sochild = Browser(bBrowser).Page(bPage).WebTable(bWebTable).ChildObjects(bObject)
		
	ElseIf sWindow = "NoWindow" Then
	
		Set bBrowser = CreateDescriptionObject(sBrowser)
		Set bPage = CreateDescriptionObject(sPage)
 		Set bFrame = CreateDescriptionObject(sFrame)
  		Set bWebTable = CreateDescriptionObject(sWebTable)
  		Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
		Set sochild = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable).ChildObjects(bObject)
	
	ElseIf sFrame= "NoFrame"  Then
	
		Set bBrowser = CreateDescriptionObject(sBrowser)
  		Set bWindow = CreateDescriptionObject(sWindow)
  		Set bPage = CreateDescriptionObject(sPage)
  		Set bWebTable = CreateDescriptionObject(sWebTable)
  		Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
		Set sochild = Browser(bBrowser).window(bWindow).Page(bPage).WebTable(bWebTable).ChildObjects(bObject)
	
	Else  	 
	
		Set bBrowser = CreateDescriptionObject(sBrowser)
  		Set bWindow = CreateDescriptionObject(sWindow)
  		Set bPage = CreateDescriptionObject(sPage)
  		Set bFrame = CreateDescriptionObject(sFrame)
  		Set bWebTable = CreateDescriptionObject(sWebTable)
  		Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
  		Set sochild = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable).ChildObjects(bObject) 	 
	End If


  If sochild.count<>0 Then 
  	 sochild(0).Highlight  	 
     sochild(0).click 
 
  End If
  
 ErrNumber = Err.Number
 ErrMessage = Err.Description
	 
 If ErrNumber<>0 Then	 	
 	writelog "","","","","ClickObjectInWebTbl(): ClickObjectInWebTbl is unsucessfull! Object Name : "&PropertyValue & " Error Number : "&err.number & " Error Description : "&Err.description
 else
 	writelog "","","","","ClickObjectInWebTbl(): ClickObjectInWebTbl is Successful. Object Name : "&PropertyValue
 End If
 
 End Function
 
 
 '******************************************** HEADER ******************************************
' Description : The function to fetch user details for BOS Login
' Creator :  Hemaraja 
' Date : 05th October, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************  
Public Function Read_UserDetails()

	Set objUserMas = CreateObject("Scripting.FileSystemObject")
	Set bUserDetail = objUserMas.OpenTextFile (UserDetailPath)

	Do While bUserDetail.AtEndofStream <> True
	    sUserdata = bUserDetail.ReadLine

	   'config file
		If instr(sUserdata,"Environment") Then
			arEnvironmentdata = Split(sUserdata,":")

			If arEnvironmentdata(1)<>"" Then
				aEnvironmentID = arEnvironmentdata(1)
				
				Environment.Value("Environment") = aEnvironmentID

				aEnvironmentID = replace(aEnvironmentID,"-TEST","")
				aEnvironmentID = replace(aEnvironmentID,"-PREP","")

				If instr(aEnvironmentID,"ZARP")<>0 Then
					aEnvironmentID = "ZA"
				ElseIf instr(aEnvironmentID,"MY")<>0 Then
					aEnvironmentID = "MY"
				ElseIf instr(aEnvironmentID,"LCASS")<>0 Then
					aEnvironmentID = "LCASS"
				ElseIf instr(aEnvironmentID,"LION")<>0 Then
					aEnvironmentID = "LION"
				End If

				Environment.Value("Market") = aEnvironmentID

			End If

		'config file
		ElseIf instr(sUserdata,"Solution") Then
			arSolutiondata = Split(sUserdata,":")

			If arSolutiondata(1)<>"" Then
				aSolutionID = arSolutiondata(1)
				Environment.Value("Solution")= aSolutionID
			End If

		'config file
		ElseIf instr(sUserdata,"ReleaseName") Then
			arReleasedata = Split(sUserdata,":")

			If arReleasedata(1)<>"" Then
				aReleaseID = arReleasedata(1)
				Environment.Value("ReleaseName")= aReleaseID
			End If

		'config file
		ElseIf instr(sUserdata,"ModuleName") Then
			arModuledata = Split(sUserdata,":")

			If arModuledata(1)<>"" Then
			   aModuleID = arModuledata(1)
			   If instr(lcase(aModuleID),"pay") Then
			       aModuleID = "pl-paylo"
			   ElseIf instr(lcase(aModuleID),"data") Then
			       aModuleID = "so-test-data-suite"
			   ElseIf instr(lcase(aModuleID),"user") Then
			       aModuleID = "so-user-role"
			   Else
			       aModuleID = "so-site-ops"
			   End If
				Environment.Value("ModuleName")= aModuleID
			End If
			
			

	    ElseIf instr(sUserdata,"Username") Then
			arLogindata = Split(sUserdata,":")

			If arLogindata(1)<>"" Then
				aLoginID = arLogindata(1)
				Environment.Value("LoginId")= aLoginID
			End If

	     ElseIf instr(sUserdata,"Password")  Then
	        arPassdata = Split(sUserdata,":")

			If arPassdata(1)<>"" Then
				aPassword = arPassdata(1)
				Environment.Value("bPassword") = aPassword
			End If


	    ElseIf instr(sUserdata,"OrgID")  Then
	        arOrgID = Split(sUserdata,":")

			If arOrgID(1)<>"" Then
				aOrgID = arOrgID(1)
				Environment.Value("bOrgID") = aOrgID
			End If

		ElseIf instr(sUserdata,"Role")  Then
	        arRole = Split(sUserdata,":")

			If arRole(1)<>"" Then
				aRole = arRole(1)
				Environment.Value("bRole") = aRole
			End If

'			If userRole<>"" Then
'				Environment.Value("bRole")=userRole
'			End If
'
		ElseIf instr(sUserdata,"AdhocStartDate")  Then
        arAdhocStartDate = Split(sUserdata,":")

			If arAdhocStartDate(1)<>"" Then
				aAdhocStartDate = arAdhocStartDate(1)
				Environment.Value("bAdhocStartDate") = aAdhocStartDate
			End If

		ElseIf instr(sUserdata,"AdhocEndDate")  Then
        arAdhocEndDate = Split(sUserdata,":")

			If arAdhocEndDate(1)<>"" Then
				aAdhocEndDate = arAdhocEndDate(1)
				Environment.Value("bAdhocEndDate") = aAdhocEndDate
			End If

		ElseIf instr(sUserdata,"EODStartDate")  Then
        arEODStartDate = Split(sUserdata,":")

			If arEODStartDate(1)<>"" Then
				aEODStartDate = arEODStartDate(1)
				Environment.Value("bEODStartDate") = aEODStartDate
			End If

		ElseIf instr(sUserdata,"EODEndDate")  Then
        arEODEndDate = Split(sUserdata,":")

			If arEODEndDate(1)<>"" Then
				aEODEndDate = arEODEndDate(1)
				Environment.Value("bEODEndDate") = aEODEndDate
			End If

		ElseIf instr(sUserdata,"AdhocEODFlag")  Then
        arAdhocEODFlag = Split(sUserdata,":")

			If arAdhocEODFlag(1)<>"" Then
				aAdhocEODFlag = arAdhocEODFlag(1)
				Environment.Value("bAdhocEODFlag") = aAdhocEODFlag
			Else
				Environment.Value("bAdhocEODFlag") = ""
			End If

	    End If
	Loop

 End Function
 
 
 Public Function Read_UserDetailsBkp()
 	
	Set objUserMas = CreateObject("Scripting.FileSystemObject")
	Set bUserDetail = objUserMas.OpenTextFile (UserDetailPath) 	
	
	Do While bUserDetail.AtEndofStream <> True
	    sUserdata = bUserDetail.ReadLine
	       
	    If instr(sUserdata,"Username") Then
			arLogindata = Split(sUserdata,":") 
			
			If arLogindata(1)<>"" Then
				aLoginID = arLogindata(1)
				Environment.Value("LoginId")= aLoginID
			End If
			
	     ElseIf instr(sUserdata,"Password")  Then	
	        arPassdata = Split(sUserdata,":") 
			
			If arPassdata(1)<>"" Then
				aPassword = arPassdata(1)
				Environment.Value("bPassword") = aPassword
			End If
	   
	    
	    ElseIf instr(sUserdata,"OrgID")  Then	
	        arOrgID = Split(sUserdata,":") 
			
			If arOrgID(1)<>"" Then
				aOrgID = arOrgID(1)
				Environment.Value("bOrgID") = aOrgID
			End If
		
		ElseIf instr(sUserdata,"Environment")  Then	
	        arOrgID = Split(sUserdata,":") 
			
			If arOrgID(1)<>"" Then
				aEnv= arOrgID(1)
				Environment.Value("ENV_Flag") = aEnv
			End If
		
		ElseIf instr(sUserdata,"Role")  Then	
	        arRole = Split(sUserdata,":") 
			
			If arRole(1)<>"" Then
				aRole = arRole(1)
				Environment.Value("bRole") = aRole
			End If
			
		ElseIf instr(sUserdata,"AdhocStartDate")  Then	
        arAdhocStartDate = Split(sUserdata,":") 
		
		If arAdhocStartDate(1)<>"" Then
			aAdhocStartDate = arAdhocStartDate(1)
			Environment.Value("bAdhocStartDate") = aAdhocStartDate
		End If
		
		ElseIf instr(sUserdata,"AdhocEndDate")  Then	
        arAdhocEndDate = Split(sUserdata,":") 
		
		If arAdhocEndDate(1)<>"" Then
			aAdhocEndDate = arAdhocEndDate(1)
			Environment.Value("bAdhocEndDate") = aAdhocEndDate
		End If
		
		ElseIf instr(sUserdata,"EODStartDate")  Then	
        arEODStartDate = Split(sUserdata,":") 
		
		If arEODStartDate(1)<>"" Then
			aEODStartDate = arEODStartDate(1)
			Environment.Value("bEODStartDate") = aEODStartDate
		End If
		
		ElseIf instr(sUserdata,"EODEndDate")  Then	
        arEODEndDate = Split(sUserdata,":") 
		
		If arEODEndDate(1)<>"" Then
			aEODEndDate = arEODEndDate(1)
			Environment.Value("bEODEndDate") = aEODEndDate
		End If

		ElseIf instr(sUserdata,"AdhocEODFlag")  Then	
        arAdhocEODFlag = Split(sUserdata,":") 
		
			If arAdhocEODFlag(1)<>"" Then
				aAdhocEODFlag = arAdhocEODFlag(1)
				Environment.Value("bAdhocEODFlag") = aAdhocEODFlag
			Else
				Environment.Value("bAdhocEODFlag") = ""
			End If
		
	   End If
	   
	Loop
 	
 End Function
 


'******************************************** HEADER ******************************************
' Name : CreateConsolidatedReceiptDictionary
' Description : Create the consolidated receipt dictionary object bz parsing the consolidated receipt xml of BOSINSTRUCTION file.
' Creator : Syed Shafi
' Date : 10th Nov,2017
' Last Modified On : 
' Last Modified By : 
' Input Parameter : 
' Output Parameter : 
'******************************************** HEADER ****************************************** 
'Public Function CreateConsolidatedReceiptDictionary()
'	
'	Dim strTranData
'	Dim bFlagNext
'	Dim xmlDoc
'	Dim nodes
'	Dim iMixedCount
'	Dim iFuelCount
'	Dim objDictQty
'	
'	Set objDictQty = CreateObject("Scripting.Dictionary")
'	
'	iMixedCount=0
'	iFuelCount=0
'	strTranData=""
'	bFlagNext = False
'	
'	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
'	xmlDoc.setProperty "SelectionLanguage", "XPath"
'	
'	strXMLFILE = GetBOSInstruction("XMLFILE")
'	xmlDoc.load(strXMLFILE)
'	
'	set nodes = xmlDoc.selectNodes("//*")    
'	
'	Add_Dictionary objDictReceiptData,"TOTALSALES",""
'	Add_Dictionary objDictReceiptData,"NETSALES",""
'	Add_Dictionary objDictReceiptData,"TOTALTAXES",""
'	Add_Dictionary objDictReceiptData,"FUEL_V-Power Nitro_SALES",""
'	
'	Add_Dictionary objDictReceiptData,strNewTranID,""
'	
'	
'	for i = 0 to nodes.length-1
'	'    ''msgbox nodes(i).nodeName & " - " & nodes(i).text 
'	'    ''msgbox "NODE NAME : " & nodes(i).nodeName 
'	'    ''msgbox "NODE TEXT : " & nodes(i).text 
'	    
'	    If bFlagNext = False Then  '1
'		    'bFuelFlag = False
'		    If instr(1,nodes(i).nodeName,"Transaction_") <> 0 Then
'		    	bFlagNext = True
'		    	strNewTranID = nodes(i).nodeName
'	'	    	objDictReceiptData.Add strNewTranID,""
'		    	Add_Dictionary objDictReceiptData,strNewTranID,""
'		    End If
'	    else
'	    	strTag = nodes(i).nodeName
'			strTagVal = nodes(i).text
'	    	If instr(1,strTag,"SHIFT") = 0 Then  '21
'	    		
'	    		If strTranData="" Then
'	    			strTranData = strTag & ":" & strTagVal
'	    		else
'	    			strTranData = strTranData & "," & strTag & ":" & strTagVal
'	    		End If
'	    			    	
'	    		If instr(1,strTag,"ItemName_")<>0 Then
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,"CURRENTITEM",strTagVal) 
'	    			
'	    			'Item Quantity Calculations
'	    			If instr(1,strTag,"_ItemName_")<>0 Then
'	    				bFuelFlag = True
'		    			if objDictQty.Exists (strTagVal)=False then
'		    				Call Update_Dictionary(objDictQty,strTagVal,strCurrentFuelQty) 
'		    			else
'		    				strCQty = get_dictionary(objDictQty,strTagVal)
'		    				strFinVol = cdbl(strCQty) + cdbl(strCurrentFuelQty)
'		    				Call Update_Dictionary(objDictQty,strTagVal,strFinVol) 
'		    			End if
'	    			else
'		    			if objDictQty.Exists (strTagVal)=False then
'		    				Call Update_Dictionary(objDictQty,strTagVal,1) 
'		    			else
'		    				strCQty = get_dictionary(objDictQty,strTagVal)
'		    				strCQty = strCQty + 1
'		    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
'		    			End if
'	    			End if
'	    			
'	    			strQTYKeyItem1 = strTagVal & "_" & "QTY"
'		    		strTOTALKeyItem2 = strTagVal & "_" & "TOTAL"
'		    		strVATKeyItem3 = strTagVal & "_" & "VAT"
'		    		sInitialValue=""
'		    		
'		    		if objDictReceiptData.Exists (strQTYKeyItem1)=False then
'		    			if bFuelFlag = False then
'		    				Call Add_Dictionary(objDictReceiptData,strQTYKeyItem1,1) 
'		    			else
'		    				strCQty = get_dictionary(objDictQty,strTagVal)
'		    				Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
'		    			End if
'		    		else
'		    			strCQty = get_dictionary(objDictQty,strTagVal)
'		    			Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
'		    		End if
'			    	
'		    		if objDictReceiptData.Exists (strTOTALKeyItem2)=False then
'		    			Call Add_Dictionary(objDictReceiptData,strTOTALKeyItem2,sInitialValue) 
'		    		End if
'		    		if objDictReceiptData.Exists (strVATKeyItem3)=False then
'		    			Call Add_Dictionary(objDictReceiptData,strVATKeyItem3,sInitialValue) 
'		    		End if
'		    	
'	    		End If
'	    		
'	    		
'	    		If instr(1,strTag,"TOTALNET") <> 0 Then
'	    			strCurrentNetSales = objDictReceiptData.Item ( "NETSALES")
'	    			If strCurrentNetSales="" Then
'	    				strNewNetSale = strTagVal
'	    			else
'	    				strNewNetSale = cdbl(strCurrentNetSales) + cdbl(strTagVal)
'	    			End If
'	    			objDictReceiptData.Item ( "NETSALES") = strNewNetSale
'				elseIf instr(1,strTag,"Litres") <> 0 Then
'		    			strCurrentFuelQty = trim(replace(strTagVal,"litre",""))
'		    			'bFuelFlag = True
'		    			
'		    	elseIf instr(1,strTag,"FUELTOTAL") <> 0 Then
'		    			strCurrentFuelTotal =strTagVal
'		    			bFuelFlag =True
'		    	elseIf instr(1,strTag,"CRCOST") <> 0 and bFuelFlag =True Then
'		    			iMixedCount=iMixedCount + 1
'						'MIXED FUEL DATA TOTAL
'	    				strCurrentTotalSales = objDictReceiptData.Item ("MIXEDFUEL_TOTAL")
'		    			If strCurrentTotalSales="" Then
'		    				strMixTotalSale = strCurrentFuelTotal
'		    			else
'		    				strMixTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
'		    			End If
'		    			objDictReceiptData.Item ("MIXEDFUEL_TOTAL") = strMixTotalSale
'		    			strMixTotalSale=""
'		    			strPreviousQty = objDictReceiptData.Item ("MIXEDFUEL_QTY")
'		    			If strPreviousQty="" Then
'		    				strNewMixSaleQty = strCurrentFuelQty
'		    			else
'		    				strNewMixSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
'		    			End If
'		    			objDictReceiptData.Item ("MIXEDFUEL_QTY") = strNewMixSaleQty
'		    			bFuelFlag =False
'		    	elseIf instr(1,strTag,"CURRENCY") <> 0 and bFuelFlag =True Then
'		    			iFuelCount=iFuelCount + 1
'			    		'FUEL DATA TOTAL
'		    			strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
'		    			If strCurrentTotalSales="" Then
'		    				strFuelTotalSale = strCurrentFuelTotal
'		    			else
'		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
'		    			End If
'		    			objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
'		    			strFuelTotalSale=""
'		    			strPreviousQty = objDictReceiptData.Item ("FUELQTY")
'		    			If strPreviousQty="" Then
'		    				strNewFuelSaleQty = strCurrentFuelQty
'		    			else
'		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
'		    			End If
'		    			objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
'	'	    		End if
'		    	elseIf strcomp(strTag,"TOTAL") = 0 Then
'		    		strCI=objDictReceiptData.item("CURRENTITEM")
''		    		If instr(1,strCI,"Diesel Extra")<>0 Then
''		    					''msgbox "CUrrent Total -"&strTagVal 			
''		    		End If
'		    	
'	    			strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
'	    			'Shift Total transaction calculation
'	    			If strCurrentTotalSales="" Then
'	    				strNewTotalSale = strTagVal
'	    			else
'	    				strNewTotalSale = cdbl(strCurrentTotalSales) + cdbl(strTagVal)
'	    			End If
'	    			objDictReceiptData.Item ( "TOTAL") = strNewTotalSale
'	    			
'	    			'Item wise Total calculation
'    				strLastTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
'    				If strLastTotal="" Then
'    					strNewTotal = cdbl(strTagVal)
'    				else
'    					strNewTotal = cdbl(strLastTotal)+cdbl(strTagVal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewTotal) 
'	    		elseIf instr(1,strTag,"TAXAMOUNT") <> 0 Then
'	    			'Item wise Total calculation
'    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem3)
'    				If strLastVAT="" Then
'    					strNewVATTotal = cdbl(strTagVal)
'    				else
'    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strVATKeyItem3,strNewVATTotal) 
'		    			
'		    	elseIf instr(1,strTag,"MOP") <> 0 Then
'		    		strCurrentTranTotal = objDictReceiptData.Item ("TOTAL")
'		    		Select Case strTagVal
'		    			Case "CASH"
'		    				strTotalCashSales = objDictReceiptData.Item ("CASH")
'		    				If strTotalCashSales="" Then
'		    					strCashTotal = strCurrentTranTotal
'			    			else
'			    				strCashTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCashSales)
'			    			End If
'			    			objDictReceiptData.Item ("CASH") = strCashTotal
'			    			strCashTotal=""
'		    			Case "DRIVEOFF"
'		    				strTotalDOSales = objDictReceiptData.Item ("DRIVEOFF")
'		    				If strTotalDOSales="" Then
'		    					strDOTotal = strCurrentTranTotal
'			    			else
'			    				strDOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalDOSales)
'			    			End If
'			    			objDictReceiptData.Item ("DRIVEOFF") = strDOTotal
'			    			strDOTotal=""
'		    			Case "LOI"
'		    				strTotalLOISales = objDictReceiptData.Item ("LOI")
'		    				If strTotalLOISales="" Then
'		    					strLOITotal = strCurrentTranTotal
'			    			else
'			    				strLOITotal = cdbl(strCurrentTranTotal) + cdbl(strTotalLOISales)
'			    			End If
'			    			objDictReceiptData.Item ("LOI") = strLOITotal
'			    			strLOITotal=""
'		    			Case "CROSSOVER"
'		    				strTotalCOSales = objDictReceiptData.Item ("CROSSOVER")
'		    				If strTotalCOSales="" Then
'		    					strCOTotal = strCurrentTranTotal
'			    			else
'			    				strCOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCOSales)
'			    			End If
'			    			objDictReceiptData.Item ("CROSSOVER") = strCOTotal
'			    			strCOTotal=""
'		    		End Select
'		    	End if
'    			
'	    	else
'	    		strTranData = strTranData & "," & strTag & ":" & strTagVal
'	    		bFlagNext = False
'	'    		objDictReceiptData.Item(strNewTranID) = strTranData
'	    		Update_Dictionary objDictReceiptData,strNewTranID,strTranData
'	    		strNewTranID=""
'	    		strTranData=""
'	    	End if  '21
'	    		    
'	    End If '1
'	    'bFuelFlag = False
'	Next
'	objDictReceiptData.Item ("MIXEDCOUNT") = iMixedCount
'	objDictReceiptData.Item ("FUELCOUNT") = iFuelCount
'	
'	'Calculate the total taxes paid from net and total sales
'	strCurrentNetSales = objDictReceiptData.Item ( "NETSALES")
'	strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
'	
'	If strCurrentNetSales<>"" Then
'		strTaxesPaid = cdbl(strCurrentTotalSales)-cdbl(strCurrentNetSales)
'		objDictReceiptData.Item ("TOTALTAXES") = round(strTaxesPaid,2)
'	End If
'	
'	
'End Function


'Public Function CreateConsolidatedReceiptDictionaryOLD()
'	
'	Dim strTranData
'	Dim bFlagNext
'	Dim xmlDoc
'	Dim nodes
'	Dim iMixedCount
'	Dim iFuelCount
'	Dim objDictQty
'	Dim objDictVATCAT
'	
'	Set objDictQty = CreateObject("Scripting.Dictionary")
'	Set objDictVATCAT = CreateObject("Scripting.Dictionary")
'	
'	iMixedCount=0
'	iFuelCount=0
'	strTranData=""
'	bFlagNext = False
'	
'	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
'	xmlDoc.setProperty "SelectionLanguage", "XPath"
'	
'	strXMLFILE = GetBOSInstruction("XMLFILE")
'	xmlDoc.load(strXMLFILE)
'	
'	set nodes = xmlDoc.selectNodes("//*")    
'	
'	Add_Dictionary objDictReceiptData,"TOTALSALES",""
'	Add_Dictionary objDictReceiptData,"NETSALES",""
'	Add_Dictionary objDictReceiptData,"TOTALTAXES",""
'	Add_Dictionary objDictReceiptData,"FUEL_V-Power Nitro_SALES",""
'	Call Add_Dictionary(objDictReceiptData,"CR_TOTAL","") 
'	
'	Add_Dictionary objDictReceiptData,strNewTranID,""
'	
'	
'	for i = 0 to nodes.length-1
'	'    ''msgbox nodes(i).nodeName & " - " & nodes(i).text 
'	'    ''msgbox "NODE NAME : " & nodes(i).nodeName 
'	'    ''msgbox "NODE TEXT : " & nodes(i).text 
'	    
'	    If bFlagNext = False Then  '1
'		    bFuelFlag = False
'		    NoTotalCost = False
'		    If instr(1,nodes(i).nodeName,"Transaction_") <> 0 Then
'		    	bFlagNext = True
'		    	strNewTranID = nodes(i).nodeName
'	'	    	objDictReceiptData.Add strNewTranID,""
'		    	Add_Dictionary objDictReceiptData,strNewTranID,""
'		    End If
'	    else
'	    	strTag = nodes(i).nodeName
'			strTagVal = nodes(i).text
'	    	If instr(1,strTag,"SHIFT") = 0 Then  '21
'	    		
'	    		If strTranData="" Then
'	    			strTranData = strTag & ":" & strTagVal
'	    		else
'	    			strTranData = strTranData & "," & strTag & ":" & strTagVal
'	    		End If
'	    			    	
'	    		If instr(1,strTag,"ItemName_")<>0 Then
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,"CURRENTITEM",strTagVal) 
'	    			
'	    			'Item Quantity Calculations
'	    			If instr(1,strTag,"_ItemName_")<>0 Then
'	    				bFuelFlag = True
'						
'		    			if objDictQty.Exists (strTagVal)=False then
'		    				Call Update_Dictionary(objDictQty,strTagVal,strCurrentFuelQty) 
'		    			else
'		    				strCQty = get_dictionary(objDictQty,strTagVal)
'		    				strFinVol = cdbl(strCQty) + cdbl(strCurrentFuelQty)
'		    				Call Update_Dictionary(objDictQty,strTagVal,strFinVol) 
'		    			End if
'	    			else
''	    				If bFuelFlag = False Then
'			    			if objDictQty.Exists (strTagVal)=False then
'			    				Call Update_Dictionary(objDictQty,strTagVal,1) 
'			    			else
'			    				strCQty = get_dictionary(objDictQty,strTagVal)
'			    				strCQty = strCQty + 1
'			    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
'			    			End if
''			    		else
''			    			if objDictQty.Exists (strTagVal)=False then
''			    				Call Update_Dictionary(objDictQty,strTagVal,1) 
''			    			else
''			    				strCQty = get_dictionary(objDictQty,strTagVal)
''			    				strCQty = strCQty + 1
''			    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
''			    			End if
''			    		End if
'	    			End if
'	    			
'	    			strQTYKeyItem1 = strTagVal & "_" & "QTY"
'		    		strTOTALKeyItem2 = strTagVal & "_" & "TOTAL"
'		    		strVATKeyItem3 = strTagVal & "_" & "VAT"
'		    		sInitialValue=""
'		    		
'		    		if objDictReceiptData.Exists (strQTYKeyItem1)=False then
'		    			if bFuelFlag = False then
'		    				Call Add_Dictionary(objDictReceiptData,strQTYKeyItem1,1) 
'		    			else
'		    				strCQty = get_dictionary(objDictQty,strTagVal)
'		    				Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
'		    			End if
'		    		else
'		    			strCQty = get_dictionary(objDictQty,strTagVal)
'		    			Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
'		    		End if
'			    	
'		    		if objDictReceiptData.Exists (strTOTALKeyItem2)=False then
'		    			Call Add_Dictionary(objDictReceiptData,strTOTALKeyItem2,sInitialValue) 
'		    		End if
'		    		if objDictReceiptData.Exists (strVATKeyItem3)=False then
'		    			Call Add_Dictionary(objDictReceiptData,strVATKeyItem3,sInitialValue) 
'		    		End if
'		    	
'	    		End If
'	    		
'	    		
'	    		If instr(1,strTag,"TOTALNET") <> 0 Then
'	    			strCurrentNetSales = objDictReceiptData.Item ( "NETSALES")
'	    			If strCurrentNetSales="" Then
'	    				strNewNetSale = strTagVal
'	    			else
'	    				strNewNetSale = cdbl(strCurrentNetSales) + cdbl(strTagVal)
'	    			End If
'	    			objDictReceiptData.Item ( "NETSALES") = strNewNetSale
'	    		elseIf instr(1,strTag,"_TAXCAT") <> 0 Then
'					strCategory = Mid(strTag,1,1)
'					strITEM_NAME = strTagVal		
'					Call Update_Dictionary(objDictVATCAT,strCategory,strITEM_NAME) 
'
'				elseIf instr(1,strTag,"Litres") <> 0 Then
'	    			strCurrentFuelQty = trim(replace(strTagVal,"litre",""))
'	    			bFuelFlag = True		    			
'		    	elseIf instr(1,strTag,"FUELTOTAL") <> 0 Then
'	    			strCurrentFuelTotal =strTagVal
'	    			bFuelFlag =True
'	    			'Item wise Total calculation
'    				strLastFuelTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
'    				If strLastFuelTotal="" Then
'    					strNewFuelTotal = cdbl(strCurrentFuelTotal)
'    				else
'    					strNewFuelTotal = cdbl(strLastFuelTotal)+cdbl(strCurrentFuelTotal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewFuelTotal) 
'    				NoTotalCost = True
'		    	elseIf instr(1,strTag,"CRCOST") <> 0 and bFuelFlag =True Then
'		    		
'		    		strExistingCRTotal = Get_Dictionary(objDictReceiptData,"CR_TOTAL")
'					If strExistingCRTotal="" Then
'    					strNewCR_Total = cdbl(strTagVal)
'    				else
'    					strNewCR_Total = cdbl(strExistingCRTotal)+cdbl(strTagVal)
'    				End If	 
'					Call Update_Dictionary(objDictReceiptData,"CR_TOTAL",strNewCR_Total) 
'				
'					'Item wise Total calculation
'    				strLastCRTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
'    				If strLastCRTotal="" Then
'    					strNewCRTotal = cdbl(strTagVal)
'    				else
'    					strNewCRTotal = cdbl(strLastCRTotal)+cdbl(strTagVal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewCRTotal)
'					strFinalCRTotalTemp = strNewCRTotal
'					strFinalCRTotal = cdbl(strFinalCRTotalTemp)+cdbl(strTagVal)
'					'Call Update_Dictionary(objDictReceiptData,"FINALCRTOTAL",strNewCRTotal)					
'		    	
'	    			iMixedCount=iMixedCount + 1
'					'MIXED FUEL DATA TOTAL
'    				strCurrentTotalSales = objDictReceiptData.Item ("MIXEDFUEL_TOTAL")
'	    			If strCurrentTotalSales="" Then
'	    				strMixTotalSale = strCurrentFuelTotal
'	    			else
'	    				strMixTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
'	    			End If
'	    			objDictReceiptData.Item ("MIXEDFUEL_TOTAL") = strMixTotalSale
'	    			strFinalFuelTotalTemp = strMixTotalSale
''	    			objDictReceiptData.Item ("FINALFUELTOTAL") = strFinalFuelTotalTemp
'	    			strMixTotalSale=""
'	    			strPreviousQty = objDictReceiptData.Item ("MIXEDFUEL_QTY")
'	    			If strPreviousQty="" Then
'	    				strNewMixSaleQty = strCurrentFuelQty
'	    			else
'	    				strNewMixSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
'	    			End If
'	    			objDictReceiptData.Item ("MIXEDFUEL_QTY") = strNewMixSaleQty
'	    			bFuelFlag =False
'		    	elseIf instr(1,strTag,"CURRENCY") <> 0 and bFuelFlag =True Then
'		    			iFuelCount=iFuelCount + 1
'			    		'FUEL DATA TOTAL
'		    			strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
'		    			If strCurrentTotalSales="" Then
'		    				strFuelTotalSale = strCurrentFuelTotal
'		    			else
'		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
'		    			End If
'		    			objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
'		    			strFinalFuelTotal = cdbl(strFinalFuelTotalTemp) + cdbl(strFuelTotalSale)
'		    			'strFuelFinalTotalSale = cdbl(strFuelTotalSale) + cdbl(strMixTotalSale)
'		    			'objDictReceiptData.Item ("FUELFINALTOTAL") = strFuelFinalTotalSale
'		    			'Call Update_Dictionary(objDictReceiptData,"FINALFUELTOTAL",strFinalFuelTotal) 
'		    			objDictReceiptData.Item ("FINALFUELTOTAL") = strFinalFuelTotal
'		    			strFuelTotalSale=empty
'		    			strFinalFuelTotal=empty
''		    			strMixTotalSale=""
'		    			strPreviousQty = objDictReceiptData.Item ("FUELQTY")
'		    			If strPreviousQty="" Then
'		    				strNewFuelSaleQty = strCurrentFuelQty
'		    			else
'		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
'		    			End If
'		    			objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
'	'	    		End if
'		    	elseIf strcomp(strTag,"TOTAL") = 0 Then
'		    		If NoTotalCost = False and bFuelFlag =False Then
'		    			strExistingCRTotal = Get_Dictionary(objDictReceiptData,"CR_TOTAL")
'						If strExistingCRTotal="" Then
'	    					strNewCR_Total = cdbl(strTagVal)
'	    				else
'	    					strNewCR_Total = cdbl(strExistingCRTotal)+cdbl(strTagVal)
'	    				End If	 
'						Call Update_Dictionary(objDictReceiptData,"CR_TOTAL",strNewCR_Total) 
'		    		End If
'		    	
'		    		strCI=objDictReceiptData.item("CURRENTITEM")
''		    		If instr(1,strCI,"Diesel Extra")<>0 Then
''		    					''msgbox "CUrrent Total -"&strTagVal 			
''		    		End If
'		    	
'	    			strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
'	    			'Shift Total transaction calculation
'	    			If strCurrentTotalSales="" Then
'	    				strNewTotalSale = strTagVal
'	    			else
'	    				strNewTotalSale = cdbl(strCurrentTotalSales) + cdbl(strTagVal)
'	    			End If
'	    			objDictReceiptData.Item ( "TOTAL") = strNewTotalSale
'	    			
'	    			If NoTotalCost = False Then
'		    			'Item wise Total calculation
'	    				strLastTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
'	    				If strLastTotal="" Then
'	    					strNewTotal = cdbl(strTagVal)
'	    				else
'	    					strNewTotal = cdbl(strLastTotal)+cdbl(strTagVal)
'	    				End If	    				
'	    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewTotal) 
'	    			End if
'	    		elseIf instr(1,strTag,"TAXAMOUNT") <> 0 Then
'	    			strCategory1 = Mid(strTag,len(strTag))
'					strITEM_NAME1 = Get_Dictionary(objDictVATCAT,strCategory1)
'					
'    			If strITEM_NAME1="V-Power Nitro+ DSL" Then
'	    			strITEM_NAME1="Shell V-Power Diesel"
'	    		else
'	    			strITEM_NAME1=strITEM_NAME1
'	    		End If
'	    			strVATKeyItem3 = strITEM_NAME1 & "_VAT"
'	    			
'	    			'Item wise Total calculation
'    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem3)
'    				If strLastVAT="" Then
'    					strNewVATTotal = cdbl(strTagVal)
'    				else
'    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strVATKeyItem3,strNewVATTotal) 
'		    			
'		    	elseIf instr(1,strTag,"MOP") <> 0 Then
'		    		strCurrentTranTotal = objDictReceiptData.Item ("TOTAL")
'		    		Select Case strTagVal
'		    			Case "CASH"
'		    				strTotalCashSales = objDictReceiptData.Item ("CASH")
'		    				If strTotalCashSales="" Then
'		    					strCashTotal = strCurrentTranTotal
'			    			else
'			    				strCashTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCashSales)
'			    			End If
'			    			objDictReceiptData.Item ("CASH") = strCashTotal
'			    			strCashTotal=""
'		    			Case "DRIVEOFF"
'		    				strTotalDOSales = objDictReceiptData.Item ("DRIVEOFF")
'		    				If strTotalDOSales="" Then
'		    					strDOTotal = strCurrentTranTotal
'			    			else
'			    				strDOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalDOSales)
'			    			End If
'			    			objDictReceiptData.Item ("DRIVEOFF") = strDOTotal
'			    			strDOTotal=""
'		    			Case "LOI"
'		    				strTotalLOISales = objDictReceiptData.Item ("LOI")
'		    				If strTotalLOISales="" Then
'		    					strLOITotal = strCurrentTranTotal
'			    			else
'			    				strLOITotal = cdbl(strCurrentTranTotal) + cdbl(strTotalLOISales)
'			    			End If
'			    			objDictReceiptData.Item ("LOI") = strLOITotal
'			    			strLOITotal=""
'		    			Case "CROSSOVER"
'		    				strTotalCOSales = objDictReceiptData.Item ("CROSSOVER")
'		    				If strTotalCOSales="" Then
'		    					strCOTotal = strCurrentTranTotal
'			    			else
'			    				strCOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCOSales)
'			    			End If
'			    			objDictReceiptData.Item ("CROSSOVER") = strCOTotal
'			    			strCOTotal=""
'		    		End Select
'		    	End if
'    			
'	    	else
'	    		strTranData = strTranData & "," & strTag & ":" & strTagVal
'	    		bFlagNext = False
'	'    		objDictReceiptData.Item(strNewTranID) = strTranData
'	    		Update_Dictionary objDictReceiptData,strNewTranID,strTranData
'	    		strNewTranID=""
'	    		strTranData=""
'	    		'Set dictionary tax category to null
'	    		Set objDictVATCAT = Nothing
'	    		Set objDictVATCAT = CreateObject("Scripting.Dictionary")
'	    	End if  '21
'	    		    
'	    End If '1
'	    'bFuelFlag = False
'	Next
'	objDictReceiptData.Item ("MIXEDCOUNT") = iMixedCount
'	objDictReceiptData.Item ("FUELCOUNT") = iFuelCount
'	
'	'Calculate the total taxes paid from net and total sales
'	strCurrentNetSales = objDictReceiptData.Item ( "NETSALES")
'	strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
'	
'	If strCurrentNetSales<>"" Then
'		strTaxesPaid = cdbl(strCurrentTotalSales)-cdbl(strCurrentNetSales)
'		objDictReceiptData.Item ("TOTALTAXES") = round(strTaxesPaid,2)
'	End If
'	
'	
'End Function
Public Function CreateConsolidatedReceiptDictionaryOld()
	
	Dim strTranData
	Dim bFlagNext
	Dim xmlDoc
	Dim nodes
	Dim iMixedCount
	Dim iFuelCount
	Dim objDictQty
	Dim objDictVATCAT
	
	Set objDictQty = CreateObject("Scripting.Dictionary")
	Set objDictVATCAT = CreateObject("Scripting.Dictionary")
	
	iMixedCount=0
	iFuelCount=0
	strTranData=""
	bFlagNext = False
	VoucherFlag = False

	
	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
	xmlDoc.setProperty "SelectionLanguage", "XPath"
	
	strXMLFILE = GetBOSInstruction("XMLFILE")
	xmlDoc.load(strXMLFILE)
	
	set nodes = xmlDoc.selectNodes("//*")    
	
	Add_Dictionary objDictReceiptData,"TOTALSALES",""
	Add_Dictionary objDictReceiptData,"NETSALES",""
	Add_Dictionary objDictReceiptData,"TOTALTAXES",""
	Add_Dictionary objDictReceiptData,"FUEL_V-Power Nitro_SALES",""
	
	Add_Dictionary objDictReceiptData,strNewTranID,""
	
	
	for i = 0 to nodes.length-1
	'    ''msgbox nodes(i).nodeName & " - " & nodes(i).text 
	'    ''msgbox "NODE NAME : " & nodes(i).nodeName 
	'    ''msgbox "NODE TEXT : " & nodes(i).text 
	    
	    If bFlagNext = False Then  '1
		    bFuelFlag = False
		    NoTotalCost = False
		    ''''''ADDED
		    bPumpFlagFlag = False
			ShellVoucherItem = False
			bMixedSalesFlag = False
			PriceOverrideFlag = False
			bPayInOutFlag = False
			BottleItemFlag = False
			strAlreadyChecked = empty
			strCurrentFuelQty = 0
			''''''ADDED
		    If instr(1,nodes(i).nodeName,"Transaction_") <> 0 Then
		    	bFlagNext = True
		    	strNewTranID = nodes(i).nodeName
	'	    	objDictReceiptData.Add strNewTranID,""
		    	Add_Dictionary objDictReceiptData,strNewTranID,""
		    End If
	    else
	    	strTag = nodes(i).nodeName
			strTagVal = nodes(i).text
	    	If instr(1,strTag,"SHIFT") = 0 Then  '21
	    		
	    		If strTranData="" Then
	    			strTranData = strTag & ":" & strTagVal
	    		else
	    			strTranData = strTranData & "," & strTag & ":" & strTagVal
	    		End If
	    			    	
	    		If instr(1,strTag,"ItemName_")<>0 AND instr(1,strTag,"CostItemName_")=0 Then
	    			
	    			Call Update_Dictionary(objDictReceiptData,"CURRENTITEM",strTagVal) 
	    			
	    			'Item Quantity Calculations
	    			If instr(1,strTag,"_ItemName_")<>0 Then
	    				bFuelFlag = True
						If strTagVal="V-Power Nitro+ DSL" Then
	    					strTagVal="Shell V-Power Diesel"
	    				else
	    					strTagVal=strTagVal
	    				End If
		    			if objDictQty.Exists (strTagVal)=False then
		    				Call Update_Dictionary(objDictQty,strTagVal,strCurrentFuelQty) 
		    			else
		    				strCQty = get_dictionary(objDictQty,strTagVal)
		    				strFinVol = cdbl(strCQty) + cdbl(strCurrentFuelQty)
		    				Call Update_Dictionary(objDictQty,strTagVal,strFinVol) 
		    			End if
	    			else
'	    				If bFuelFlag = False Then
			    			if objDictQty.Exists (strTagVal)=False then
			    				Call Update_Dictionary(objDictQty,strTagVal,1) 
			    			else
			    				strCQty = get_dictionary(objDictQty,strTagVal)
			    				strCQty = strCQty + 1
			    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
			    			End if
'			    		else
'			    			if objDictQty.Exists (strTagVal)=False then
'			    				Call Update_Dictionary(objDictQty,strTagVal,1) 
'			    			else
'			    				strCQty = get_dictionary(objDictQty,strTagVal)
'			    				strCQty = strCQty + 1
'			    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
'			    			End if
'			    		End if
	    			End if
	    			
	    			strQTYKeyItem1 = strTagVal & "_" & "QTY"
		    		strTOTALKeyItem2 = strTagVal & "_" & "TOTAL"
		    		strVATKeyItem3 = strTagVal & "_" & "VAT"
		    		strPerItemKeyItem4 = strTagVal & "_" & "PerItemCost" ''''Delcy added
		    		strTAXKeyItem5 = strNewTranID & "_" & "TAXPERCENT"
		    		
		    		'SHAFI ADDED TESTING----------------
		    		strMOP = "@MOP@"
		    		strQTYKeyItemMOP1 = strTagVal & "_" & strMOP & "_" & "QTY"
		    		strTOTALKeyItemMOP2 = strTagVal & "_" & strMOP & "_" & "TOTAL"
		    		strVATKeyItemMOP3 = strTagVal & "_" & strMOP & "_" & "VAT"	    		
		    		'SHAFI ADDED TESTING----------------
		    		
		    		sInitialValue=""
		    		
		    		if objDictReceiptData.Exists (strQTYKeyItem1)=False then
		    			if bFuelFlag = False then
		    				Call Add_Dictionary(objDictReceiptData,strQTYKeyItem1,1) 
		    				'SHAFI ADDED TESTING----------------		
		    				Call Add_Dictionary(objDictReceiptData,strQTYKeyItemMOP1,1) 
		    				'SHAFI ADDED TESTING----------------
		    			else
		    				strCQty = get_dictionary(objDictQty,strTagVal)
		    				Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
		    				'SHAFI ADDED TESTING----------------
		    				Call Update_Dictionary(objDictReceiptData,strQTYKeyItemMOP1,strCQty)
		    				'SHAFI ADDED TESTING----------------
		    			End if
		    		else
		    			strCQty = get_dictionary(objDictQty,strTagVal)
		    			Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
		    			'SHAFI ADDED TESTING----------------
		    			Call Update_Dictionary(objDictReceiptData,strQTYKeyItemMOP1,strCQty)
		    			'SHAFI ADDED TESTING----------------
		    		End if
			    	
		    		if objDictReceiptData.Exists (strTOTALKeyItem2)=False then
		    			Call Add_Dictionary(objDictReceiptData,strTOTALKeyItem2,sInitialValue) 
		    			'SHAFI ADDED TESTING----------------
		    			Call Add_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,sInitialValue) 
		    			'SHAFI ADDED TESTING----------------
		    		End if
		    		if objDictReceiptData.Exists (strVATKeyItem3)=False then
		    			Call Add_Dictionary(objDictReceiptData,strVATKeyItem3,sInitialValue)
						'SHAFI ADDED TESTING----------------
		    			Call Add_Dictionary(objDictReceiptData,strVATKeyItemMOP3,sInitialValue) 
		    			'SHAFI ADDED TESTING----------------		    			
		    		End if
		    	ElseIf instr(1,strTag,"CostItemName_")<>0 Then
	    			PriceOverrideFlag = True
	    		End If
	    		
	    		If instr(1,strTag,"PumpTestItem")<> 0 Then
	    			bPumpFlagFlag = True
	    		End If
				
	    		If instr(1,strTag,"TYPE")<> 0 Then
	    			If instr(1,strTagVal,"Paid-in")<> 0 or instr(1,strTagVal,"Paid-out")<> 0 Then
	    				bPayInOutFlag = True
	    			End If
	    		End If
	    		
	    		If instr(1,strTag,"TOTALNET") <> 0 Then
	    			strCurrentNetSale = strTagVal
	    			
	    			'''''' PERITEMCOST LOGIC
	    		elseIf instr(1,strTag,"PerItemCost_") <> 0 Then
	    			If objDictReceiptData.Exists (strPerItemKeyItem4)=False Then
	    				'strLastIndItemTotal = Get_Dictionary(objDictReceiptData,strPerItemKeyItem4)
	    				Call Update_Dictionary(objDictReceiptData,strPerItemKeyItem4,strTagVal)
	    				'Saves all the items
	    				
	    				'ItemsPresent = split(strPerItemKeyItem4,"_PerItemCost")
	    				
	    				If strPerItemKeyItem4 = "" Then
								strAlreadyChecked  = strVATKeyItem3
						elseif instr(1,strAlreadyChecked,strVATKeyItem3) =0 then
								strAlreadyChecked = strAlreadyChecked & strVATKeyItem3
						End If

	    			End If
				elseIf instr(1,strTag,"TAX_") <> 0 Then
	    			'If objDictReceiptData.Exists (strTAXKeyItem5)=False Then
	    				'strLastIndItemTotal = Get_Dictionary(objDictReceiptData,strPerItemKeyItem4)
	    				strTagValPercent = cdbl(replace(strTagVal,"%",""))
	    				Call Update_Dictionary(objDictReceiptData,strTAXKeyItem5,strTagValPercent)
    			
	    		elseIf instr(1,strTag,"_TAXCAT") <> 0 Then
					strCategory = Mid(strTag,1,1)
					strITEM_NAME = strTagVal		
					Call Update_Dictionary(objDictVATCAT,strCategory,strITEM_NAME) 

				elseIf instr(1,strTag,"Litres") <> 0 Then
	    			strCurrentFuelQty = trim(replace(strTagVal,"litre",""))
	    			bFuelFlag = True		    			
		    	elseIf instr(1,strTag,"FUELTOTAL") <> 0 Then
	    			strCurrentFuelTotal =strTagVal
	    			bFuelFlag =True
	    			'Item wise Total calculation
    				strLastFuelTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
    				If strLastFuelTotal="" Then
    					strNewFuelTotal = cdbl(strCurrentFuelTotal)
    				else
'    					If strCurrentFuelTotal = objDictReceiptData.Item ("VOUCHER") Then
'    						strNewFuelTotal = cdbl(strLastFuelTotal)+cdbl(strCurrentFuelTotal)+cdbl(strCurrentFuelTotal)
'    					Else
    					strNewFuelTotal = cdbl(strLastFuelTotal)+cdbl(strCurrentFuelTotal)
'    					End If
    				End If	    				
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewFuelTotal) 
    				'SHAFI ADDED TESTING----------------
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,strNewFuelTotal) 
    				'SHAFI ADDED TESTING----------------
    				'All fuel total calculation-shabana
    				
    			If bmixedsalesflag = False and bPumpFlagFlag = False and  bPayInOutFlag = False  Then
    				iFuelCount=iFuelCount + 1	
						
						strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
		    			If strCurrentTotalSales="" Then
		    				
		    				strFuelTotalSale = strCurrentFuelTotal
						Else
		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
		    				End if
						objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
		    			strFuelTotalSale=""
		    			
		    			
						strPreviousQty = objDictReceiptData.Item ("FUELQTY")
		    			If strPreviousQty="" Then
		    				strNewFuelSaleQty = strCurrentFuelQty
		    			else
							strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				End If
		    			objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
		    			strNewFuelSaleQty = ""
		    			
				End If
    				NoTotalCost = True
		    	elseIf instr(1,strTag,"CRCOST") <> 0 and bFuelFlag =True Then
		    	
		    		'Item wise Total calculation
		    		bMixedSalesFlag = True

					If bMixedSalesFlag = True Then
		    			strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
		    			strPreviousQty = objDictReceiptData.Item ("FUELQTY")
		    			'iFuelCount=iFuelCount - 1	
		    			strFuelTotalSale = cdbl(strCurrentTotalSales) - cdbl(strCurrentFuelTotal)
		    			strNewFuelSaleQty = cdbl(strPreviousQty) - cdbl(strCurrentFuelQty)
		    		End If
		    		objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
		    		strFuelTotalSale=""
		    		objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
		    		strNewFuelSaleQty = ""

    				strLastCRTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
    				If strLastCRTotal="" Then
    					strNewCRTotal = cdbl(strTagVal)
    				else
    					strNewCRTotal = cdbl(strLastCRTotal)+cdbl(strTagVal)
    				End If	    				
					objDictReceiptData.Item ( "CRCOST") = strNewCRTotal
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewCRTotal) 
    				
    				'SHAFI ADDED TESTING----------------
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,strNewFuelTotal) 
    				'SHAFI ADDED TESTING----------------
    				
    				'MIXED FUEL DATA TOTAL
    				strCurrentTotalSales = objDictReceiptData.Item ("MIXEDFUEL_TOTAL")
    				
	    			If strCurrentTotalSales="" Then
	    				strMixTotalSale = strCurrentFuelTotal
	    				iMixedCount=iMixedCount + 1
	    			else
	    				strMixTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
	    				iMixedCount=iMixedCount + 1
	    			End If
	    			objDictReceiptData.Item ("MIXEDFUEL_TOTAL") = strMixTotalSale
	    			strMixTotalSale=""
	    			strPreviousQty = objDictReceiptData.Item ("MIXEDFUEL_QTY")
	    			If strPreviousQty="" Then
	    				strNewMixSaleQty = strCurrentFuelQty
	    			else
	    				strNewMixSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
	    			End If
	    			objDictReceiptData.Item ("MIXEDFUEL_QTY") = strNewMixSaleQty
	    			bFuelFlag =False
    				
		    	elseIf instr(1,strTag,"CURRENCY") <> 0 Then
					
					'Calculating Rounding					
					strRoundingTagVal = split(strTagVal," ")
		    		
		    		strRoundingVal = strRoundingTagVal(ubound(strRoundingTagVal))
		    		
		    		If instr(strRoundingVal,"Rand") Then
		    			strRoundingT =empty
		    		ElseIf instr(strRoundingVal,"R") and instr(strRoundingVal,".") and instr(strRoundingVal,"-")=0 Then
		    			
		    			strRoundingT = replace(strRoundingVal,"R","")
		    			strRoundingTT = split(strRoundingT,".")
		    			strRoundingTTT = strRoundingTT(1)
		    			
		    			strCurrentTransactionTotalT = split(strCurrentTransactionTotal,".")
		    			strCurrentTransactionTotalTT = strCurrentTransactionTotalT(1)
		    					    			
		    			strTotalRoundingVal = objDictReceiptData.Item ("ROUNDING")
		    		
			    		If strTotalRoundingVal="" Then
			    			strTotalRoundingVal = (strCurrentTransactionTotalTT - strRoundingTTT)/100
						Else		    		
			    			strTotalRoundingVal = cdbl(strTotalRoundingVal) + (strCurrentTransactionTotalTT - strRoundingTTT)/100		
			    		End If
			    		
			    		objDictReceiptData.Item ("ROUNDING")=strTotalRoundingVal
			    			
		    		End If
			    		
					If bFuelFlag =True Then
						'FUEL DATA TOTAL
			    		
		    			strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
		    			If strCurrentTotalSales="" Then
		    				iFuelCount=iFuelCount + 1
		    				strFuelTotalSale = strCurrentFuelTotal
		    			
		    			Else 
						    If strTagVal="Voucher" Then
'						    iFuelCount=iFuelCount + 3
'		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal) +  cdbl(strCurrentFuelTotal)
'		    				objDictReceiptData.Item ("VOUCHER") = cdbl(strCurrentFuelTotal)
							 iFuelCount=iFuelCount + 1'3
						    
		    			
						strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
		    				strTOTALKeyVoucher = strTOTALKeyItem2&"Voucher"
		    				Call Add_Dictionary(objDictVoucher,strTOTALKeyVoucher,strCurrentFuelTotal)


		    				objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
		    				objDictReceiptData.Item ("VOUCHER") = cdbl(strCurrentFuelTotal)
			    			strFuelTotalSale=""
			    			VoucherFlag = True
			    			'-----------------------------Shabana Testing---------------------------Commented
'		    				Else
'		    				iFuelCount=iFuelCount + 1

'		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
							'-----------------------------Shabana Testing---------------------------Commented									
		    				End if
		    				
		    			End If
'		    			objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
'		    			strFuelTotalSale=""
		    			strPreviousQty = objDictReceiptData.Item ("FUELQTY")
		    			If strPreviousQty="" Then
		    				strNewFuelSaleQty = strCurrentFuelQty
		    			else
		    				If strTagVal="Voucher" Then
		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				strTOTALKeyVoucherQty = strTOTALKeyItem2&"VoucherQty"
		    				Call Add_Dictionary(objDictVoucher,strTOTALKeyVoucherQty,strCurrentFuelQty)



		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
'		    				Else
'		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				End If
		    			End If
		    			'objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
	'	    		End if

		    	End If
		    	
		    	elseIf strcomp(strTag,"TOTAL") = 0 Then
		    		strCI=objDictReceiptData.item("CURRENTITEM")
'		    		If instr(1,strCI,"Diesel Extra")<>0 Then
'		    					''msgbox "CUrrent Total -"&strTagVal 			
'		    		End If
		    		strCurrentTransactionTotal = strTagVal
	    			strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
	    			'Shift Total transaction calculation
	    			if bPumpFlagFlag = False and bPayInOutFlag = False  Then
	    			If strCurrentTotalSales="" Then
	    				strNewTotalSale = strTagVal
	    			else
	    				strNewTotalSale = cdbl(strCurrentTotalSales) + cdbl(strTagVal)
	    			End If
	    			
	    			objDictReceiptData.Item ( "TOTAL") = strNewTotalSale
	    			
	    			If NoTotalCost = False Then
		    			'Item wise Total calculation
	    				strLastTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
	    				If strLastTotal="" Then
	    					strNewTotal = cdbl(strTagVal)
	    				else
	    					strNewTotal = cdbl(strLastTotal)+cdbl(strTagVal)
	    				End If	    				
	    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewTotal) 
	    				strRefundCashtotal = strTagVal	    				
	    				'SHAFI ADDED TESTING----------------
    					Call Update_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,strRefundCashtotal) 
    				'SHAFI ADDED TESTING----------------
	    			End if
	    			End If
'	    		elseIf instr(1,strTag,"TAXAMOUNT") <> 0 Then
'	    			strCategory1 = Mid(strTag,len(strTag))
'					strITEM_NAME1 = Get_Dictionary(objDictVATCAT,strCategory1)
'    				strTemp = strVATKeyItem3
'	    			'Item wise Total calculation
'    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem3)
'    				If strLastVAT="" Then
'    					strNewVATTotal = cdbl(strTagVal)
'    				else
'    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strVATKeyItem3,strNewVATTotal) 
'		    		strVATKeyItem3 = strTemp

				elseIf instr(1,strTag,"TAX_") <> 0 Then
					
					strCategory = Mid(strTag,len(strTag))
					strITEM_NAME = Get_Dictionary(objDictVATCAT,strCategory)
    				
    				strAlreadyCheckedTemp = strAlreadyChecked
    				If instr(strAlreadyCheckedTemp,"VAT")<> 0 Then
    				strAlreadyCheckedTemp = split(strAlreadyCheckedTemp,"_VAT")
					Else
					strAlreadyCheckedTemp = split(strAlreadyCheckedTemp," ")					
    				End If
    				
    				
    				strTemp = strVATKeyItem
   				
   				For ihead = 0 To ubound(strAlreadyCheckedTemp)
   				strITEM_NAME = strAlreadyCheckedTemp(ihead)	
   					strTagVal = cdbl(replace(strTagVal,"%",""))
    				If strITEM_NAME="V-Power Nitro+ DSL" Then
	    				strITEM_NAME="Shell V-Power Diesel"
	    			else
	    				strITEM_NAME=strITEM_NAME
	    			End If
	    			strVATKeyItem = strITEM_NAME & "_TAXPERCENT"
	    			
	    			'Item wise Total calculation
    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem)
    				If strLastVAT="" Then
    					strNewVATTotal = cdbl(strTagVal)
    				else
    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
    				End If	    				
					Call Update_Dictionary(objDictReceiptData,strVATKeyItem,strNewVATTotal)

					'SHAFI ADDED TESTING----------------
		    		Call Update_Dictionary(objDictReceiptData,strVATKeyItemMOP,strNewVATTotal) 
		    		'SHAFI ADDED TESTING----------------
		    		Next	

		    		strVATKeyItem = strTemp
					elseIf instr(1,strTag,"TAXAMOUNT") <> 0 Then
	    			'Item wise Total calculation
    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem3)
    				If strLastVAT="" Then
    					strNewVATTotal = cdbl(strTagVal)
    				else
    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
    				End If	    				
    				Call Update_Dictionary(objDictReceiptData,strVATKeyItem3,strNewVATTotal) 
					'SHAFI ADDED TESTING----------------
		    		Call Update_Dictionary(objDictReceiptData,strVATKeyItemMOP3,strNewVATTotal) 
		    		'SHAFI ADDED TESTING----------------
					strVATKeyItem3 = strTemp
		    	elseIf instr(1,strTag,"MOP") <> 0 Then
		    		'strCurrentTranTotal = objDictReceiptData.Item ("TOTAL")
		    		strCurrentTranTotal = strCurrentTransactionTotal
		    		
		    		'SHAFI ADDED LOGIC TESING-------------
		    		strCI=objDictReceiptData.item("CURRENTITEM")
		    		MOPSearch = strCI & "_@MOP@_" 
		    		
		    		For each MopKey in objDictReceiptData.keys
		    			strMopTemp1=Get_Dictionary(objDictReceiptData,MopKey)
		    			
		    			If instr(1,MopKey,MOPSearch)<>0 Then
		    				
		    				if instr(1,MopKey,"_QTY") then
			    				NewMOPKey = strCI & "_" & strTagVal & "_QTY" 
			    			ElseIf instr(1,MopKey,"_TOTAL") then 
			    				NewMOPKey = strCI & "_" & strTagVal & "_TOTAL" 
			    			ElseIf instr(1,MopKey,"_VAT") then 
			    				NewMOPKey = strCI & "_" & strTagVal & "_VAT" 
			    			End if
			    			Call Add_Dictionary(objDictReceiptData,NewMOPKey,strMopTemp1) 
		    			End If
		    			
		    		Next
		    	
		    		'SHAFI ADDED LOGIC TESING-------------
		    		Select Case strTagVal
		    			Case "CASH"
		    				strTotalCashSales = objDictReceiptData.Item ("CASH")
		    				
		    				If strRoundingT<>0 Then
			    				strCurrentTranTotal = strRoundingT
								strRoundingT = empty    				
			    			End If		    				
		    				
		    				If VoucherFlag =False and bPumpFlagFlag = False Then
		    					If strTotalCashSales="" Then
			    					strCashTotal = strCurrentTranTotal
				    			else
				    				strCashTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCashSales)
				    			End If	
		    				objDictReceiptData.Item ("CASH") = strCashTotal
			    			strCashTotal=""	    				
		    				Else
			    				VoucherFlag =False 
		    				End If
		    				
							If strCurrentNetSale>0 and bPumpFlagFlag = False Then
								stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    				'stCashVATTotal = strNewNetSale
			    				If stTotalOfNetSales = "" Then		    					
			    					stTotalOfNetSales = strCurrentNetSale
			    				Else
			    					stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(strCurrentNetSale)		    						
			    				End If
				    			'stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stCashTotalNetSales)
				    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales	
							End If
		    				
			    				
		    			Case "DRIVEOFF"
		    				
		    				strTotalDOSales = objDictReceiptData.Item ("DRIVEOFF")
		    				If strTotalDOSales="" Then
		    					strDOTotal = strCurrentTranTotal
			    			else
			    				strDOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalDOSales)
			    			End If
			    			objDictReceiptData.Item ("DRIVEOFF") = strDOTotal
							
							'Total Net Sales
			    			stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    			stDOVATTotal = (strDOTotal/1.15)*(0)
			    			stDOTotalNetSales = cdbl(strDOTotal) - cdbl(stDOVATTotal)
			    			stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stDOTotalNetSales)
			    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales
			    			strDOTotal=""
		    			
		    			Case "LOI"
		    				strTotalLOISales = objDictReceiptData.Item ("LOI")
		    				If strTotalLOISales="" Then
		    					strLOITotal = strCurrentTranTotal
			    			else
			    				strLOITotal = cdbl(strCurrentTranTotal) + cdbl(strTotalLOISales)
			    			End If
			    			objDictReceiptData.Item ("LOI") = strLOITotal
							
							'Total Net Sales
			    			stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    			stLOIVATTotal = (strLOITotal/1.15)*(0.15)
			    			stLOITotalNetSales = cdbl(strLOITotal) - cdbl(stLOIVATTotal)
			    			stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stLOITotalNetSales)
			    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales			    			
			    			
			    			strLOITotal=""
		    			Case "CROSSOVER"
		    				strTotalCOSales = objDictReceiptData.Item ("CROSSOVER")
		    				If strTotalCOSales="" Then
		    					strCOTotal = strCurrentTranTotal
			    			else
			    				strCOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCOSales)
			    			End If
							objDictReceiptData.Item ("CROSSOVER") = strCOTotal		    			
			    			
			    			'Total Net Sales
			    			stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    			stCOVATTotal = (strCOTotal/1.15)*(0)
			    			stCOTotalNetSales = cdbl(strCOTotal) - cdbl(stCOVATTotal)
			    			stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stCOTotalNetSales)
			    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales	
			    			
			    			strCOTotal=""
			    			
			    		Case "REFUND","REFUND AS REVERSAL"
		    				If PriceOverrideFlag = True  Then
								
								'Subtracting the price override reversal items from the CASHTOTAL TOTAL
			    				TotalCashVal = objDictReceiptData.Item ("CASH")
			    				If TotalCashVal = "" Then
			    					TotalCashValueT = cdbl(0) - cdbl(sCurrentMOPVal)
			    				Else
			    					TotalCashValueT = cdbl(TotalCashVal)-cdbl(sCurrentMOPVal)
			    				End If			    				
								objDictReceiptData.Item ("CASH") = TotalCashValueT
																
								'Subtracting the price override reversal items from the NETSALES TOTAL
								TotalNetSalesValue = objDictReceiptData.Item ("NETSALES")
								If TotalNetSalesValue = "" Then
									TotalNetSalesValueT = cdbl(0) - cdbl(strCurrentItemNetSales)								
								Else
									TotalNetSalesValueT = cdbl(TotalNetSalesValue)-cdbl(strCurrentItemNetSales)
								End If																
                                objDictReceiptData.Item ("NETSALES") = TotalNetSalesValueT
								
			    			End If
			    			
		    				If bFuelFlag = True and bPumpFlagFlag = False Then 
		    					strCurrentTotalSales = objDictReceiptData.Item ("MIXEDFUEL_TOTAL")
		    					
		    					If strCurrentTotalSales = "" Then
		    						strFRefundTotal = cdbl(strCurrentFuelTotal)
		    					Else
		    						strFRefundTotal = cdbl(strCurrentFuelTotal) +  cdbl(strCurrentTotalSales)
		    					End If	    					
		    					'objDictReceiptData.Item ("MIXEDFUEL_TOTAL")=strFRefundTotal
		    					
		    					strCurrentTotalQty = objDictReceiptData.Item ("MIXEDFUEL_QTY")
		    					'strFRefundQty = cdbl(strCurrentFuelQty) +  cdbl(strCurrentTotalQty)
		    					If strCurrentTotalQty = "" Then
		    						strFRefundQty = cdbl(strCurrentFuelQty)
		    					Else
		    						strFRefundQty = cdbl(strCurrentFuelQty) +  cdbl(strCurrentTotalQty)
		    					End If
		    					'objDictReceiptData.Item ("MIXEDFUEL_QTY")=strFRefundQty
		    				    					
		    					
		    					strTOTALKeyRefundFuel = strTOTALKeyItem2&"RefundFuel"

		    					Call Add_Dictionary(objDictVoucher,strTOTALKeyRefundFuel,strCurrentFuelTotal)
		    					
		    					strTOTALKeyRefundFuelQty = strTOTALKeyItem2&"RefundFuelQty"
		    					Call Add_Dictionary(objDictVoucher,strTOTALKeyRefundFuelQty,strCurrentFuelQty)
		    	
		    					
		    				End If
		    		End Select
		    	End if
    			
	    	else
	    		strTranData = strTranData & "," & strTag & ":" & strTagVal
	    		bFlagNext = False
	'    		objDictReceiptData.Item(strNewTranID) = strTranData
	    		Update_Dictionary objDictReceiptData,strNewTranID,strTranData
	    		strNewTranID=""
	    		strTranData=""
	  			bPumpFlagFlag = False
	  			VoucherFlag =False
	    		'Set dictionary tax category to null
	    		Set objDictVATCAT = Nothing
	    		Set objDictVATCAT = CreateObject("Scripting.Dictionary")
	    	End if  '21
	    		    
	    End If '1
	    'bFuelFlag = False
	Next
	objDictReceiptData.Item ("MIXEDCOUNT") = iMixedCount
	objDictReceiptData.Item ("FUELCOUNT") = iFuelCount
	
	'Calculate the total taxes paid from net and total sales
	strCurrentNetSales = objDictReceiptData.Item ( "NETSALES")
	strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
	
	If strCurrentNetSales<>"" Then
		strTaxesPaid = cdbl(strCurrentTotalSales)-cdbl(strCurrentNetSales)
		objDictReceiptData.Item ("TOTALTAXES") = round(strTaxesPaid,2)
	End If
	
	
End Function

Public Function CreateConsDict()
	
On error resume next
set globalTable=CreateObject("Scripting.Dictionary")
set objDictReceiptData=CreateObject("Scripting.Dictionary")
Set objFuelCategory = CreateObject("Scripting.Dictionary")
Set objTransac = CreateObject("Scripting.Dictionary")
Set objRefund = CreateObject("Scripting.Dictionary")

set subTable=CreateObject("Scripting.Dictionary")

set objXML=CreateObject("Microsoft.XMLDOM")
strXMLFILE = GetBOSInstruction("XMLFILE")
objXML.Load(strXMLFILE)
set sourceElement=objXML.documentElement.childNodes

for inti=0 to sourceElement.Length-1
	set subTable=CreateObject("Scripting.Dictionary")
	set childSourceElement=objXML.documentElement.childNodes.item(inti).childNodes
	for intj=0 to childSourceElement.Length-1
		subTable.Add childSourceElement.Item(intj).nodeName,childSourceElement.Item(intj).text
	Next
	globalTable.Add sourceElement.Item(inti).nodeName,subTable
	set subTable=Nothing
Next

set objXML=Nothing
fuelItemSetupFlag= False

For each data in globalTable

	flagQtyRefund = False
	
	If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"REFUND") <> 0 Then
		refundItemDet=Split(globalTable.Item(data).Item("CRITEMDETAILS0"),"||")
		If cdbl(refundItemDet(1)*refundItemDet(2))<>cdbl(globalTable.Item(data).Item("TOTAL")) and instr(refundItemDet(0),"Bottel")=0 Then
			globalTable.Item(data).Item("CRITEMDETAILS0")=refundItemDet(0)&"||"&cdbl(refundItemDet(1)-1)&"||"&refundItemDet(2)&"||"&refundItemDet(3)&"||"&refundItemDet(4)
			poAmount=formatnumber(abs(cdbl(globalTable.Item(data).Item("TOTAL"))-cdbl(refundItemDet(2)*cdbl(refundItemDet(1)-1))))
			globalTable.Item(data).Item("CRITEMDETAILS1")=refundItemDet(0)&"||"&cint(1)&"||"&poAmount&"-||"&refundItemDet(3)&"||"&refundItemDet(4)
			globalTable.Item(data).Item("TAXDETAILS1")=globalTable.Item(data).Item("TAXDETAILS0")
		End If
	
	End  If
	
	
	If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL FUEL") <> 0 or instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL MIX") <> 0 Then
		globalTable.Item(data).Item("TOTAL")=globalTable.Item(data).Item("TOTAL")&"-"
		tNumber=data
		tNumber = cint(Split(tNumber,"_")(1))
		followUpTransac="Transaction_"&String(5-len(""&(tNumber+1)),"0")&(tNumber+1)
		globalTable.Item(data).Item("TIME")=globalTable.Item(followUpTransac).Item("TIME")
	End  If
	
	If globalTable.Item(data).Exists("FLOATNO")=False Then
		Call UpdateNewValueInDictionary(objDictReceiptData,"TRANSACTION_COUNT","Add",1)
	Else
		Call UpdateNewValueInDictionary(objDictReceiptData,"TOTAL_SIGNON","Add",globalTable.Item(data).Item("FLOATNO"))
		Call UpdateNewValueInDictionary(objDictReceiptData,"SHIFTLIST","Append",cint(globalTable.Item(data).Item("SHIFT")))
	End  If
	
	If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL BOTTLE TRANSACTION") <> 0 Then
		Call UpdateNewValueInDictionary(objDictReceiptData,"TRANSACTION_COUNT","Add",-1)
		Call UpdateNewValueInDictionary(objDictReceiptData,"CONTAINERRETURN_COUNT","Add",1)
	End If
	
	If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CROSS OVER") <> 0 Then
		Call UpdateNewValueInDictionary(objDictReceiptData,"CROSSOVER_COUNT","Add",1)
		Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("SHIFT")&"_SALE_COUNT","Add",1)
		Call UpdateNewValueInDictionary(objDictReceiptData,"SALE_COUNT","Add",1)
		
		Call UpdateNewValueInDictionary(objDictReceiptData,"PromotionShiftList","Append",globalTable.Item(data).Item("SHIFT"))
		
	ElseIf instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"SALE") <> 0 or instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"FOLLOW") <> 0 Then
		Call UpdateNewValueInDictionary(objDictReceiptData,"SALE_COUNT","Add",1)
		Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("SHIFT")&"_SALE_COUNT","Add",1)
		Call UpdateNewValueInDictionary(objDictReceiptData,"PromotionShiftList","Append",globalTable.Item(data).Item("SHIFT"))
	End If
	
	
	If (globalTable.Item(data).Exists("CRITEMDETAILS0") and globalTable.Item(data).Exists("FUELDETAILS0")) Then
		Call UpdateNewValueInDictionary(objDictReceiptData,"MIXEDCOUNTNEW","Add",1)		
	ElseIf globalTable.Item(data).Exists("FUELDETAILS0") Then
		Call UpdateNewValueInDictionary(objDictReceiptData,"FUELCOUNTNEW","Add",1)
	ElseIf globalTable.Item(data).Exists("CRITEMDETAILS0") Then
	    Call UpdateNewValueInDictionary(objDictReceiptData,"CRCOUNTNEW","Add",1)
	End If
	
	
	
	If globalTable.Item(data).Exists("SHIFT")=True Then
		If instr(objDictReceiptData.Item("ShiftNoDetails"),globalTable.Item(data).Item("SHIFT"))=0 Then
		
		     
		     Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("SHIFT")&"_STARTTIME","Update",globalTable.Item(data).Item("TIME"))
'	    	Call UpdateNewValueInDictionary(objDictReceiptData,"ShiftsTimeDetail","Append",globalTable.Item(data).Item("TIME"))
	    End If
	    
	  Call UpdateNewValueInDictionary(objDictReceiptData,"ShiftNoDetails","Append",globalTable.Item(data).Item("SHIFT"))  
	End  If
	
	For Iterator = 0 To 3
		
		If Iterator = 0 Then
			counter=0
			sTypeOfItem = "CRITEMDETAILS"
		
		ElseIf  Iterator = 1  Then
			counter=0
			sTypeOfItem = "FUELDETAILS"
			
		ElseIf  Iterator = 2  Then
			sTypeOfItem="BOTTLEITEMDETAILS"
			
		ElseIf  Iterator = 3  Then
			sTypeOfItem="CONSOLIDATEDDATA"
			
		End If
		
	Do while True
	
	If (globalTable.Item(data).Exists(sTypeOfItem&counter) or (globalTable.Item(data).Exists("BOTTLEITEMDETAILS")) and Iterator = 2) and not(globalTable.Item(data).Exists("PUMPTEST")) Then
		
	
		 sTypeOfTrans = globalTable.Item(data).Item("TYPEOFTRANSACTION")
		 
		 If sTypeOfItem = "BOTTLEITEMDETAILS" Then
		 	strCRDetailsVal = globalTable.Item(data).Item(sTypeOfItem)
		 Else
		 	strCRDetailsVal = globalTable.Item(data).Item(sTypeOfItem&counter)
		 End If	
		 
		 sDetailVal = split(strCRDetailsVal,"||")
		 
		 'ITEM NAME
		 
		 sItemName = sDetailVal(0)
		 
		 Call UpdateNewValueInDictionary(objDictReceiptData,"ItemNameList","Append",sItemName)
		 
		 sItemReductions = 0
		 
		 'ITEM QTY (ITEM WISE)
		 
		 sItemQty = sDetailVal(1)
		 
		 sPerItemCost = replace(sDetailVal(2),"-","")
		 
		 
		  If sTypeOfItem = "CRITEMDETAILS"  or sTypeOfItem = "BOTTLEITEMDETAILS" Then

			 	sTagItem = "CR"
			 
			ElseIf sTypeOfItem = "FUELDETAILS" Then
			 	
			 	fuelTaxType = sDetailVal(4)
			 	sTagItem = "FUEL"
			 
			End If
		 If flagQtyRefund = False and instr(sTypeOfTrans,"CANCEL")<>0 and not(instr(sTypeOfTrans,"CANCEL BOTTLE TRANSACTION")<>0) Then
		 	
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"COUNTITEMREFUND","Add",1)
		 	flagQtyRefund = True
		 	
		 End If
		 
		 
		 If instr(sTypeOfTrans,"CANCEL")<>0 and not(instr(sTypeOfTrans,"REFUND")<>0) and not(instr(sTypeOfTrans,"CANCEL BOTTLE TRANSACTION")<>0) or instr(sTypeOfTrans,"VOID")<>0 Then
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_QTYITEMREFUND","Add",1)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_QTYITEMREFUNDVOID","Add",sItemQty)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_TOTALITEMREFUNDVOID","Add",cdbl(formatnumber(sPerItemCost * sItemQty)))
		 ElseIf instr(sTypeOfTrans,"CANCEL BOTTLE TRANSACTION")<>0 Then
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"QTYCONTAINERRETURN","Add",sItemQty)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"TOTALCONTAINERRETURN","Add",cdbl(formatnumber(sPerItemCost * sItemQty)))
		 End If 
		 
		 
		 If instr(sTypeOfTrans,"REFUND")<>0 and sTypeOfItem <> "BOTTLEITEMDETAILS" Then
		 	
		 	If objRefund.Exists(sItemName)=False Then
		 		Add_Dictionary objRefund,sItemName,"Refund"
		 	End If
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_REFUNDPRICES","Append",sPerItemCost)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"REFUND_"&sItemName&sPerItemCost&"_TOTAL","Add",sPerItemCost)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"REFUND_"&sItemName&sPerItemCost&"_TOTALWITHOUTTAX","Add",cdbl(formatnumber(sPerItemCost/1.15)))
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"REFUND_"&sItemName&"_TOTALWITHOUTTAX","Add",cdbl(formatnumber(sPerItemCost/1.15)))
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"REFUND_"&sItemName&"_TOTAL","Add",sPerItemCost)
		 End If
		 
		 If instr(sTypeOfTrans,"CANCEL")<>0 Then
		 	
		 	
		 	sItemQty = -sItemQty
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALITEMREFUND","Add",(formatnumber(sPerItemCost * sItemQty)))
		 	
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_TOTALITEMREFUND","Add",(formatnumber(sPerItemCost * sItemQty)))
		 	
		 	
		 End If
		 
		 
		 If globalTable.Item(data).Exists("BOTTLEITEMDETAILS") and Iterator = 2 Then
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"QTYCONTAINERITEM","Add",sItemQty)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,"TOTALCONTAINERITEM","Add",cdbl(formatnumber(sPerItemCost * sItemQty)))
		 End If
		 
		 Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_QTY","Add",sItemQty)
		 Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_QTY","Add",sItemQty)
		 
		 'PER ITEM COST (ITEM WISE)
		 
		 
		 If instr(sTypeOfTrans,"CANCEL")=0 Then
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_PERITEMCOST","Update",sPerItemCost)
		 End If
		 'TOTAL WITH TAX (ITEM WISE)
		 
		 sItemTotal = cdbl(formatnumber(sPerItemCost * sItemQty))
		 
		 'PRICE OVERRIDE VALUE OFF
		 
		 If globalTable.Item(data).Exists("PRICEOVERRIDEITEM") Then
		 	
		 	sPriceOverrideData = split(globalTable.Item(data).Item("PRICEOVERRIDEITEM"),"||") 
		 	
		 	If sPriceOverrideData(0) = sItemName  Then
		 		
		 		If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL FUEL") <> 0 or instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL MIX") <> 0 Then
		 			sPriceOverrideValue = -sPriceOverrideData(4)
		 		Else
		 			sPriceOverrideValue = sPriceOverrideData(4)
		 		End  If
		 		
		 		sItemTotal = cdbl(sItemTotal) + cdbl(sPriceOverRideValue)
		 		sItemReductions = sItemReductions + cdbl(sPriceOverRideValue)
		 		If sItemQty>0 Then
		 			Call UpdateNewValueInDictionary(objDictReceiptData,"PRICEOVERRIDE_QTY","Add",1)
				Else
					Call UpdateNewValueInDictionary(objDictReceiptData,"PRICEOVERRIDE_QTY","Add",-1)				
		 		End If
		 		
		 		Call UpdateNewValueInDictionary(objDictReceiptData,"PRICEOVERRIDE_TOTAL","Add",sPriceOverrideValue)
		 		
		 		
		 	End If
		 
		 End If

		 
		 If globalTable.Item(data).Exists("PROMOTIONDETAILS") Then
		 	
		 	sPromotionData = split(globalTable.Item(data).Item("PROMOTIONDETAILS"),"||") 
		 	
		 	If sPromotionData(0) = sItemName  Then

		 		If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL FUEL") <> 0 or instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL MIX") <> 0 Then
		 			sPromotionValue = -sPromotionData(3)
		 		Else
		 			sPromotionValue = sPromotionData(3)
		 		End  If
		 		
		 		sItemTotal = cdbl(sItemTotal) + cdbl(sPromotionValue)
		 		sItemReductions = sItemReductions + cdbl(sPromotionValue)
		 		
		 		If instr(sPromotionData(1),"Buy1Get1")<>0 or instr(sPromotionData(1),"BOGO")<>0 or instr(sPromotionData(1),"B1G1")<>0 Then
					tempQty=cint(abs(sItemQty)/2)
				ElseIf instr(sPromotionData(1),"BuyNGet1") Then
					tempQty=cint(abs(sItemQty)/3)
				Else
					tempQty=cint(abs(sItemQty))
				End If
		 		
		 		
		 		Call UpdateNewValueInDictionary(objDictReceiptData,"PROMOTION_QTY","Add",(tempQty))
		 		Call UpdateNewValueInDictionary(objDictReceiptData,"PROMOTION_TOTAL","Add",sPromotionValue)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,"PromotionNameList","Append",sPromotionData(1))
'				Call UpdateNewValueInDictionary(objDictReceiptData,"PromotionShiftList","Append",globalTable.Item(data).Item("SHIFT"))
				Call UpdateNewValueInDictionary(objDictReceiptData,"PromotionItemNameList","Append",sItemName)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,sPromotionData(1)&"_PROMOTIONTOTAL","Add",sPromotionValue)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,sPromotionData(1)&globalTable.Item(data).Item("SHIFT")&"_QTY","Add",sItemQty)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,sPromotionData(1)&"_QTY","Add",sItemQty)
		 		
		 		If instr(sTypeOfTrans,"CANCEL") = 0 Then
		 			Call UpdateNewValueInDictionary(objDictReceiptData,sPromotionData(1)&"_QTYWITHOUTREFUND","Add",sItemQty)
		 		End If
		 	End If
		 	
		 End If
		 	
		 If sTypeOfItem = "FUELDETAILS" Then
		 	If instr(globalTable.Item(data).Item("PUMPNO"),"||")=0 Then
		 		Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("PUMPNO")&"_"&sItemName&"_QTY","Add",sItemQty)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("PUMPNO")&"_"&sItemName&"_TOTALWITHTAX","Add",sItemTotal)
		 	Else
		 		pump=Split(globalTable.Item(data).Item("PUMPNO"),"||")
		 		Call UpdateNewValueInDictionary(objDictReceiptData,pump(counter)&"_"&sItemName&"_QTY","Add",sItemQty)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,pump(counter)&"_"&sItemName&"_TOTALWITHTAX","Add",sItemTotal)
		 	End If		 
		 End If
		 
		 If instr(sTypeOfTrans,"CANCEL") = 0 Then
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALREDUCTIONS","Add",sItemReductions)
		 	Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_TOTALREDUCTIONS","Add",sItemReductions)
		 End If
		
		 Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALWITHTAX","Add",sItemTotal)
		 Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_"&sTagItem&"_TOTALWITHTAX","Add",sItemTotal)
		 Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHTAX","Add",sItemTotal)
		 
		If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"MIX")<>0 and sTypeOfItem = "FUELDETAILS" Then
			counter=1
		End If
		
		If (globalTable.Item(data).Exists("CRITEMDETAILS0") and globalTable.Item(data).Exists("FUELDETAILS0")) and instr(sTypeOfTrans,"CANCEL MIX") = 0 Then
			Call UpdateNewValueInDictionary(objDictReceiptData,"MIX_"&sTagItem&"_TOTALWITHTAX","Add",sItemTotal)
			Call UpdateNewValueInDictionary(objDictReceiptData,"MIX_"&sTagItem&"_QTY","Add",sItemQty)
		End If
		 '---------------------------------TAX RELATED DETAILS-------------------------------------------
		 
		 If (globalTable.Item(data).Exists("TAXDETAILS"&counter) or (globalTable.Item(data).Exists("TAXDETAILS"&lCounter) and globalTable.Item(data).Exists("BOTTLEITEMDETAILS"))) or fuelTaxType = "NoTax" Then  
		 	
		 	If fuelTaxType = "NoTax" Then
		 		strTaxDetailsVal = sItemName&"||"&fuelTaxType&"||15.00"
		 
		 	ElseIf counter = -1 Then
		 		
		 		strTaxDetailsVal = globalTable.Item(data).Item("TAXDETAILS"&lCounter)
		 		
		 	Else
		 		strTaxDetailsVal = globalTable.Item(data).Item("TAXDETAILS"&counter)
		 	End If
		 	
			sTaxDetails = split(strTaxDetailsVal,"||")
		 	
		 	If sTaxDetails(0) = sItemName  Then
			
			'TAX NAME (ITEM WISE)
			
			 sTaxName = trim(sTaxDetails(2))&"%"
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TAXNAME","Update",sTaxName)
			 
			'TAX TOTAL (ITEM WISE)
			sTaxOnItem=0
			 
			 If globalTable.Item(data).Exists("PROMOTIONDETAILS") Then
				If instr(sPromotionData(1),"Buy1Get1")<>0 or instr(sPromotionData(1),"BOGO")<>0 or instr(sPromotionData(1),"B1G1")<>0 Then
					transQty=cint(sItemQty/2)
				ElseIf instr(sPromotionData(1),"BuyNGet1") Then
					transQty=cint(sItemQty/3)*2
				Else
					transQty=cint(sItemQty)
				End If
			Else
				transQty=cint(sItemQty)				
			End If
			 
			If instr(sTypeOfTrans,"CANCEL") <> 0 Then
				sPerItemCost=-sPerItemCost
				transQty=-transQty			 
			End If
			 
			 
			 If sTypeOfItem <> "FUELDETAILS" Then
			 	
				 
				 If not(globalTable.Item(data).Exists("PROMOTIONDETAILS")) and not(globalTable.Item(data).Exists("PRICEOVERRIDEITEM")) Then
				 
					sTaxOnItem = cdbl(formatnumber(formatnumber((sItemTotal/transQty) - ((sItemTotal/(1+(formatnumber(sTaxDetails(2))/100)))/transQty))*transQty))
				 ElseIf not(globalTable.Item(data).Exists("PRICEOVERRIDEITEM")) Then
				 	sTaxOnItem = cdbl(formatnumber(formatnumber((sItemTotal/transQty) - ((sItemTotal/(1+(formatnumber(sTaxDetails(2))/100)))/transQty))*transQty))
				 ElseIf (globalTable.Item(data).Exists("PROMOTIONDETAILS")) and (globalTable.Item(data).Exists("PRICEOVERRIDEITEM")) Then
				 	perDisc=-(cdbl(sPromotionValue/(sItemTotal-sPromotionValue)))
					poDisc=(cdbl(sPriceOverrideValue)+cdbl(sPerItemCost))*perDisc
				 	If sPriceOverrideData(0) = sItemName  Then
				 	
				 		tempAmount = formatnumber(cdbl(sPriceOverrideValue)+cdbl(sPerItemCost)-poDisc)
				 		sTaxOnItem = sTaxOnItem+cdbl(formatnumber(formatnumber((tempAmount) - ((tempAmount/(1+(formatnumber(sTaxDetails(2))/100)))))))
				 	End If
				 	
				 	If sPromotionData(0) = sItemName  Then
				 	
						tempAmount=cdbl(formatnumber(cdbl(sItemTotal)-cdbl(sPriceOverrideValue)-cdbl(sPerItemCost)+poDisc))			 	
				 		sTaxOnItem = sTaxOnItem+cdbl(formatnumber(formatnumber((tempAmount/(transQty-1)) - ((tempAmount/(1+(formatnumber(sTaxDetails(2))/100)))/(transQty-1)))*(transQty-1)))
				 	End If
				 Else
					sTaxOnItem = cdbl(formatnumber(formatnumber((sItemTotal/transQty) - ((sItemTotal/(1+(formatnumber(sTaxDetails(2))/100)))/transQty))*transQty))
				 	
				 End If
			Else
				sTaxOnItem = cdbl(formatnumber((sItemTotal) - ((sItemTotal/(1+(formatnumber(sTaxDetails(2))/100))))))
			End If
			
			If instr(sTypeOfTrans,"CANCEL") <> 0 Then
				 sPerItemCost=-sPerItemCost			
			End If
			 
			
			 Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALTAX","Add",sTaxOnItem)
			 
			 If sTypeOfItem = "CRITEMDETAILS" or sTypeOfItem = "BOTTLEITEMDETAILS" Then
			 	
			 	sTagT = sTaxName&"_CR"
			 	sTagItem = "CR"
			 
			 ElseIf sTypeOfItem = "FUELDETAILS" Then
			 	
			 	sTagT = sTaxName&"_FUEL"
			 	sTagItem = "FUEL"
			 
			 End If
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sTagT&"_TOTALTAX","Add",sTaxOnItem)		
			 
			 
			 'TOTAL WITH TAX (TAXCATEGORY WISE)
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sTagT&"_TOTALWITHTAX","Add",sItemTotal)'CHECK
			 
			 sItemTotalWithOutTax = sItemTotal - sTaxOnItem
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALWITHOUTTAX","Add",sItemTotalWithOutTax)
			 Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_"&sTagItem&"_TOTALWITHOUTTAX","Add",sItemTotalWithOutTax)
			 If globalTable.Item(data).Exists("PROMOTIONDETAILS") Then
				Call UpdateNewValueInDictionary(objDictReceiptData,sPromotionData(1)&"_TOTALWITHOUTTAX","Add",sItemTotalWithOutTax)
			 End If
			Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHOUTTAX","Add",sItemTotalWithOutTax)
			
			If (globalTable.Item(data).Exists("CRITEMDETAILS0") and globalTable.Item(data).Exists("FUELDETAILS0")) and instr(sTypeOfTrans,"CANCEL MIX") = 0 Then
				Call UpdateNewValueInDictionary(objDictReceiptData,"MIX_"&sTagItem&"_TOTALWITHOUTTAX","Add",sItemTotalWithOutTax)
			End If
			 
			 'TOTAL WITHOUT REFUND 
			  
			  If instr(sTypeOfTrans,"CANCEL")=0 Then
			  
			  	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_QTYWITHOUTREFUND","Add",sItemQty)
				Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_QTYWITHOUTREFUND","Add",sItemQty)			  	
			  	sNewTotalWithoutRefund = cdbl(formatnumber(sPerItemCost*sItemQty))
			 	
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALWITHOUTREFUND","Add",sNewTotalWithoutRefund)
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sTagItem&"_TOTALWITHOUTREFUND","Add",sNewTotalWithoutRefund)
			 	sNewTotalWithoutRefundAndTax = sNewTotalWithoutRefund - sTaxOnItem			 			 
			 	
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALWITHOUTTAXANDREFUND","Add",sNewTotalWithoutRefundAndTax)
			 	
			 	'TOTAL WITHOUT REFUND (ALL TRANSACTIONS)
			 			 
			 	Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHOUTREFUND","Add",sNewTotalWithoutRefund)
			 
			 	
			  End If
			 
			 'TOTAL WITHOUT TAX (TAXCATEGORY WISE)
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sTagT&"_TOTALWITHOUTTAX","Add",sItemTotalWithOutTax)
						
			 End If
			 
			 
			
		End If
		
		 '-------------------------------------CATEGORY DETAILS OF ONLY CR--------------------------------------------
		 
		 If sTypeOfItem = "CRITEMDETAILS" or sTypeOfItem = "BOTTLEITEMDETAILS" Then
		 
			 'CATEGORY QTY
			 
			 sCategoryName = sDetailVal(4)						
			 	
			 sCategoryQty = sItemQty
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYQTY","Add",sCategoryQty)
			 	
			 'CATEGORY TOTAL WITH TAX
			 
			 sCategoryTotal = sItemTotal
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHTAX","Add",sCategoryTotal)
			
			 'CATEGORY TOTAL WITHOUT TAX
			 
			 sCategoryTotalWithOutTax = sItemTotalWithOutTax
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHOUTTAX","Add",sCategoryTotalWithOutTax)
			 
			 'CATEGORY TOTAL WITH TAX WITHOUT REFUND
			 If instr(sTypeOfTrans,"CANCEL")=0 Then			
			  
			  	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYQTYWITHOUTREFUND","Add",sItemQty)
			 	
			 	sNewTotalWithoutRefund = cdbl(formatnumber(sPerItemCost*sItemQty))
			 	
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHOUTREFUND","Add",sNewTotalWithoutRefund)
			 	''''''''''TOTAL REDUCTIONS CATEGORY-WISE
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALREDUCTIONS","Add",sItemReductions)			 				 
			 	
			 	'TOTAL WITH TAX WITHOUT REFUND (ALL TRANSACTIONS)
			 
			 	Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_CR_TOTALWITHOUTTAXANDREFUND","Add",sNewTotalWithoutRefund)
			
			 	
			 Else
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALITEMREFUND","Add",(formatnumber(sPerItemCost * sItemQty)))
			  End If
			  
			 If Iterator = 2 Then
			 	Exit Do
			 End If
		
		ElseIf sTypeOfItem = "FUELDETAILS" Then		
		
			If fuelItemSetupFlag = False Then
		    					
				Call CreateFuelCategoryList
				fuelItemSetupFlag = True
				
			End If
			
			sCategoryName = Get_Dictionary(objFuelCategory,sDetailVal(0))
			
			sCategoryQty = sItemQty
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYQTY","Add",sCategoryQty)
			 	
			 'CATEGORY TOTAL WITH TAX
			 
			 sCategoryTotal = sItemTotal
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHTAX","Add",sCategoryTotal)
			
			 'CATEGORY TOTAL WITHOUT TAX
			 
			 sCategoryTotalWithOutTax = sItemTotalWithOutTax
			 
			 Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHOUTTAX","Add",sCategoryTotalWithOutTax)
			 
			 'CATEGORY TOTAL WITH TAX WITHOUT REFUND
			 If instr(sTypeOfTrans,"CANCEL")=0 Then			
			  
			  	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYQTYWITHOUTREFUND","Add",sItemQty)
			 	
			 	sNewTotalWithoutRefund = cdbl(formatnumber(sPerItemCost*sItemQty))
			 	
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHOUTREFUND","Add",sNewTotalWithoutRefund)
			 	
			 	'TOTAL WITH TAX WITHOUT REFUND (ALL TRANSACTIONS)
			 
			 	Call UpdateNewValueInDictionary(objDictReceiptData,"ALLTRANSACTIONS_FUEL_TOTALWITHOUTREFUND","Add",sNewTotalWithoutRefund)
			
			 	
			 	''''''''''TOTAL REDUCTIONS CATEGORY-WISE
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALREDUCTIONS","Add",sItemReductions)			 				 
			 	
			 	
			 Else
			 	Call UpdateNewValueInDictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALITEMREFUND","Add",(formatnumber(sPerItemCost * sItemQty)))
			  End If
		
		End If
		 counter=counter+1
	
	ElseIf globalTable.Item(data).Exists("PAIDIN") Then
	
		If objTransac.Exists(data)=False Then
			Call UpdateNewValueInDictionary(objDictReceiptData,"TRANSACTION_COUNT","Add",-1)
			Add_Dictionary objTransac,data,"Present"
			Call UpdateNewValueInDictionary(objDictReceiptData,"PAIDIN","Add",Split(globalTable.Item(data).Item("PAIDIN"),"||")(1))
			Call UpdateNewValueInDictionary(objDictReceiptData,"PAIDIN_COUNT","Add",1)
			
			paidinName = Split(globalTable.Item(data).Item("PAIDIN"),"||")(0)
			sLastList = Get_Dictionary(objDictReceiptData,"PAIDINLIST")
			 	
			 If instr(sLastList,paidinName)=0 Then
				sNewList = sLastList & "||" & paidinName
				Call Update_Dictionary(objDictReceiptData,"PAIDINLIST",sNewList)
				Call UpdateNewValueInDictionary(objDictReceiptData,paidinName&"_PAIDIN","Add",Split(globalTable.Item(data).Item("PAIDIN"),"||")(1))
				Call UpdateNewValueInDictionary(objDictReceiptData,paidinName&"_PAIDINQTY","Add",1)
			End  If
			
		End If
		Exit Do
	
	ElseIf globalTable.Item(data).Exists("PAIDOUT") Then
	
		If objTransac.Exists(data)=False Then
			Call UpdateNewValueInDictionary(objDictReceiptData,"TRANSACTION_COUNT","Add",-1)
			Add_Dictionary objTransac,data,"Present"
			Call UpdateNewValueInDictionary(objDictReceiptData,"PAIDOUT","Add",Split(globalTable.Item(data).Item("PAIDOUT"),"||")(1))
			Call UpdateNewValueInDictionary(objDictReceiptData,"PAIDOUT_COUNT","Add",1)
			
			paidoutName = Split(globalTable.Item(data).Item("PAIDOUT"),"||")(0)
			sLastList = Get_Dictionary(objDictReceiptData,"PAIDOUTLIST")
			 	
			 If instr(sLastList,paidoutName)=0 Then
				sNewList = sLastList & "||" & paidoutName
				Call Update_Dictionary(objDictReceiptData,"PAIDOUTLIST",sNewList)
				Call UpdateNewValueInDictionary(objDictReceiptData,paidoutName&"_PAIDOUT","Add",Split(globalTable.Item(data).Item("PAIDOUT"),"||")(1))
				Call UpdateNewValueInDictionary(objDictReceiptData,paidoutName&"_PAIDOUTQTY","Add",1)
			End  If
			
		End If
		Exit Do
		
	ElseIf globalTable.Item(data).Exists("PUMPTEST") and globalTable.Item(data).Exists(sTypeOfItem&counter) Then
	
		If objTransac.Exists(data)=False Then
			Call UpdateNewValueInDictionary(objDictReceiptData,"TRANSACTION_COUNT","Add",-1)
			Add_Dictionary objTransac,data,"Present"
			
		End If
		sItemDetails=Split(globalTable.Item(data).Item("FUELDETAILS"&counter),"||")
		sItemName=sItemDetails(0)
		sItemQTY=sItemDetails(1)
		sItemTotal=sItemDetails(3)
		counter=counter+1
		Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_TOTALPUMPTEST","Add",sItemTotal)
		Call UpdateNewValueInDictionary(objDictReceiptData,sItemName&"_QTYPUMPTEST","Add",sItemQTY)
		Call UpdateNewValueInDictionary(objDictReceiptData,"PUMPTESTCOUNT","Add",1)
		Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("PUMPNO")&"_"&sItemName&"_QTY","Add",sItemQTY)
		 		Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("PUMPNO")&"_"&sItemName&"_TOTALWITHTAX","Add",sItemTotal)
		 	
	
	Else
		If counter<>0 Then
			lCounter=counter
		End If
		
		 counter=counter-1
		 
		 If sTypeOfItem="CONSOLIDATEDDATA" and not(globalTable.Item(data).Exists("PUMPTEST")) Then
			 
			 'MOP TOTAL
			 
			 sMOPName = globalTable.Item(data).Item("MOP")
			 
			 sMOPName = trim(replace(sMOPName," ",""))
			 sMOPName = split(sMOPName,"&")
			 
			 For id = 0 to ubound(sMOPName)
			 
			 If sMOPName(id)="LETTEROFINTENT" Then
			 	sMOPName(id)="LOI"
			 End If
			 
			 If globalTable.Item(data).Exists(sMOPName(id)&"TOTAL") Then
			 
			 	sLastList = Get_Dictionary(objDictReceiptData,"MOPLIST")
			 	
			 	If instr(sLastList,sMOPName(id)&"TOTAL")=0 Then
			 		
			 		sNewList = sLastList&"||"&sMOPName(id)&"TOTAL"
			 		Call Update_Dictionary(objDictReceiptData,"MOPLIST",sNewList)
			 	
			 	End If
				If instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL FUEL") <> 0 or instr(globalTable.Item(data).Item("TYPEOFTRANSACTION"),"CANCEL MIX") <> 0 Then
					globalTable.Item(data).Item(sMOPName(id)&"TOTAL") = globalTable.Item(data).Item(sMOPName(id)&"TOTAL")&"-"
				End  If
			 	sMOPTotal = globalTable.Item(data).Item(sMOPName(id)&"TOTAL")
			 				 	
			 	
			 	If (sMOPName(id) = "CASH") Then
			 		If right(sMOPTotal,1) <> "-" Then
			 			If int(right(sMOPTotal,1)) <> 0 Then
			 				cashVal=(mid(sMOPTotal,1,len(sMOPTotal)-1))
					 		roundVal=formatnumber(sMOPTotal-cashVal)
					 		Call UpdateNewValueInDictionary(objDictReceiptData,"ROUNDINGTOTAL","Add",roundVal)
							Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("SHIFT")&"_ROUNDINGTOTAL","Add",roundVal)
					 		Call UpdateNewValueInDictionary(objDictReceiptData,"ROUNDINGTOTALQTY","Add",1)
					 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTAL","Add",cashVal)
					 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTALQTY","Add",1)
					 	Else
					 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTAL","Add",sMOPTotal)
					 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTALQTY","Add",1)
			 			End If			 			
			 		ElseIf int(left(right(sMOPTotal,2),1)) <> 0 Then
			 			cashVal=(mid(sMOPTotal,1,len(sMOPTotal)-2)&"-")
			 			roundVal=formatnumber(sMOPTotal-cashVal)
			 			Call UpdateNewValueInDictionary(objDictReceiptData,"ROUNDINGTOTAL","Add",roundVal)
						Call UpdateNewValueInDictionary(objDictReceiptData,globalTable.Item(data).Item("SHIFT")&"_ROUNDINGTOTAL","Add",roundVal)
			 			Call UpdateNewValueInDictionary(objDictReceiptData,"ROUNDINGTOTALQTY","Add",1)
				 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTAL","Add",cashVal)	
				 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTALQTY","Add",1)
					
					Else
			 			Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTAL","Add",sMOPTotal)
						Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTALQTY","Add",1)			 			
			 		End If
			 		
			 	Else
			 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTAL","Add",sMOPTotal)
			 		Call UpdateNewValueInDictionary(objDictReceiptData,sMOPName(id)&"TOTALQTY","Add",1)
			 	End If
			 	
			 	
			 
			 End If
			 
			 Next
			 
		 End If
		 counter = -1
		 Exit Do
		 
		 End if
		
	 Loop
	
	Next

 Next


End Function


Public Function CreateConsolidatedReceiptDictionary()
	
	Dim strTranData
	Dim bFlagNext
	Dim xmlDoc
	Dim nodes
	Dim iMixedCount
	Dim iFuelCount
	Dim objDictQty
	Dim objDictVATCAT
	
	Set objDictQty = CreateObject("Scripting.Dictionary")
	Set objDictVATCAT = CreateObject("Scripting.Dictionary")
	Set objFuelCategory = CreateObject("Scripting.Dictionary")
	
	Set CurrentItemReceiptDetails = CreateObject("Scripting.Dictionary")
	
	iMixedCount=0
	iFuelCount=0
	strTranData=""
	bFlagNext = False
	VoucherFlag = False
	
	sReversalFlag = True'Keep

	
	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
	xmlDoc.setProperty "SelectionLanguage", "XPath"
	
	strXMLFILE = GetBOSInstruction("XMLFILE")
	xmlDoc.load(strXMLFILE)
	
	set nodes = xmlDoc.selectNodes("//*")    
	
	Add_Dictionary objDictReceiptData,"TOTALSALES",""
	Add_Dictionary objDictReceiptData,"NETSALES",""
	Add_Dictionary objDictReceiptData,"TOTALTAXES",""
	Add_Dictionary objDictReceiptData,"FUEL_V-Power Nitro_SALES",""
	
	Add_Dictionary objDictReceiptData,strNewTranID,""
	
	
	for i = 0 to nodes.length-1

	    
	    If bFlagNext = False Then
		    bFuelFlag = False
		    NoTotalCost = False

		    bPumpFlagFlag = False
			ShellVoucherItem = False
			bMixedSalesFlag = False
			PriceOverrideFlag = False
			bPayInOutFlag = False
			BottleItemFlag = False
			strAlreadyChecked = empty
			sReversalFlag = False'Keep
			sPriceOverRideFlag = False 'Keep
			sPumpTestFlag = False 'Keep
			sCRItemFlag = False 'Keep
			sFuelItemFlag = False 'Keep
			sPromotionTranFlag = False 'Keep
			strCurrentFuelQty = 0
	
		    If instr(1,nodes(i).nodeName,"Transaction_") <> 0 Then
		    	bFlagNext = True
		    	strNewTranID = nodes(i).nodeName
		    	Add_Dictionary objDictReceiptData,strNewTranID,""
		    End If
		    
	   	Else
	    	strTag = nodes(i).nodeName
			strTagVal = nodes(i).text
	    	
	    	If instr(1,strTag,"SHIFT") = 0 Then
	    		
	    		If strTranData="" Then
	    			strTranData = strTag & ":" & strTagVal
	    		Else
	    			strTranData = strTranData & "," & strTag & ":" & strTagVal
	    		End If
	    			    	
	    		
	    		If instr(1,strTag,"TYPEOFTRANSACTION")<>0 and instr(1,strTagVal,"CANCEL")<>0  Then
'	    			
'	    			sReversalFlag = True
'	    			
'	    		ElseIf instr(1,strTag,"PRICEOVERRIDEITEM")<>0 Then
'	    		
'	    			sPriceOverRideFlag = True
'	    			
'	    			sPriceOverRideValue = split(strTagVal,"||")(4)
'	    		
'	    		ElseIf instr(1,strTag,"PROMOTIONDETAILS") <> 0  Then
'				
'					sPromoAmount = split(strTagVal,"||")(2)
'					
'					sPromotionTranFlag = True
'					
'					'''''''''''''''''''''''''''''''''''''
'					
'					sItemNameTOTAL = sItemNameTOTAL + sPromoAmount
'					
'					Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sItemName&"_TOTALWITHTAX",sItemNameTOTAL)
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sItemName&"_TOTALWITHTAX",sItemNameTOTAL)
'
'
'					strPromoInAllTotalWithTax = Get_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHTAX")
'						
'					strPromoExcAllTotalWithTax = cdbl(strPromoInAllTotalWithTax) + cdbl(sPromoAmount)
'					
'					Call Update_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHTAX",strPromoExcAllTotalWithTax) 
'
'					
'    				Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sCategoryName&"_CATEGORYTOTALWITHTAX",sItemNameTOTAL)
'    			
'	    			strLastCategoryTOTAL = Get_Dictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHTAX")
'	    			
'	    			strNewCategoryTOTAL = GetNewConsolidatedValue(strLastCategoryTOTAL,sCurrentCategoryTotal) 				
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHTAX",strNewCategoryTOTAL)
'					
'	    		   '-------------------------
'	    		ElseIf instr(1,strTag,"PUMPTEST")<>0 Then
'	    		
'	    			sPumpTestFlag = True
'	    		
'	    		ElseIf (instr(1,strTag,"FUELDETAILS")<>0 OR instr(1,strTag,"CRITEMDETAILS")<>0) and sPumpTestFlag = False Then
'	    				    			
'	    			sValues = split(strTagVal,"||")
'	    			
'	    			'NAME (ITEMWISE)
'	    			
'		    		sItemName = sValues(0)
'	
'	    			
'	    			
'	    			'QUANTITY (ITEMWISE)
'	    				
'	    			sItemNameQTY = sValues(1)
'	    			
'	    			Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sItemName&"_QTY",sItemNameQTY) 'New Current
'	    			
'					strLastQTY = Get_Dictionary(objDictReceiptData,sItemName&"_QTY")
'	    			
'	    			If sReversalFlag = True and instr(1,strTag,"CRITEMDETAILS")<>0 Then
'	    				
'	    				sItemNameQTY = -sItemNameQTY
'	    				
'	    				strNewQTY = GetNewConsolidatedValue(strLastQTY,-sItemNameQTY)
'	    			Else
'	    				strNewQTY = GetNewConsolidatedValue(strLastQTY,sItemNameQTY)
'	    			End If
'	    			  				
'	    			Call Update_Dictionary(objDictReceiptData,sItemName&"_QTY",strNewQTY) 
'	    			
'
'					
'					'TOTAL WITH TAX (ALL TRANSACTIONS)	    			
'		    			
'	    			sPerItemCost = replace(sValues(2),"-","")
'	    			
'	    			sItemNameTOTAL = sPerItemCost*sItemNameQTY
'	    			
'	    			sCurrentTotal = sItemNameTOTAL
'	    			
'	    			If sPriceOverRideFlag = True Then
'						
'						sCurrentTotal = cdbl(sCurrentTotal) + cdbl(sPriceOverRideValue)
'					
'					End If
'					
'					strLastAllTaxWithTOTAL = Get_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHTAX")
'				
'					strNewAllTaxWithTOTAL = GetNewConsolidatedValue(strLastAllTaxWithTOTAL,sCurrentTotal)  
'				
'					Call Update_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHTAX",strNewAllTaxWithTOTAL) 
'	    			
'	    			
'	    			'TOTAL WITH TAX (ITEMWISE)
'	    				
'    				Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sItemName&"_TOTALWITHTAX",sItemNameTOTAL) 'New Current
'    			
'	    			strLastTOTAL = Get_Dictionary(objDictReceiptData,sItemName&"_TOTALWITHTAX")
'	    			
'	    			strNewTOTAL = GetNewConsolidatedValue(strLastTOTAL,sItemNameTOTAL)
'
'					If sPriceOverRideFlag = True Then
'						strNewTOTAL = cdbl(strNewTOTAL) + cdbl(sPriceOverRideValue)
'					End If
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sItemName&"_TOTALWITHTAX",strNewTOTAL) 
'	    				    			
'	    			
'	    			'CATEGORY TOTAL WITH TAX (ITEMWISE) 
'	    			
'	    			If instr(1,strTag,"CRITEMDETAILS")<>0 Then
'	    				
'	    				sCRItemFlag = True
'	    				
'	    				'CATEGORY NAME 	    				
'	    		
'    					sCategoryName = sValues(4)
'    					
'    					Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_CATEGORYNAME",sCategoryName) 
'    			
'	    				
'						'CATEGORY QTY
'						
'	    				sCurrentCategoryQty = sItemNameQTY
'						
'						Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sCategoryName&"_CATEGORYQTY",sCurrentCategoryQty) 'New Current
'						
'						strLastCategoryQTY = Get_Dictionary(objDictReceiptData,sCategoryName&"_CATEGORYQTY")
'		    			
'		    			strNewCategoryQTY = GetNewConsolidatedValue(strLastCategoryQTY,sCurrentCategoryQty) 				
'		    			
'		    			Call Update_Dictionary(objDictReceiptData,sCategoryName&"_CATEGORYQTY",strNewCategoryQTY)
'		    	
'	    				
'	    				'CATEGORY TOTAL WITH TAX
'	    				
'	    				sCurrentCategoryTotal = sItemNameTOTAL
'	    				
'	    				Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sCategoryName&"_CATEGORYTOTALWITHTAX",sCurrentCategoryTotal) 'New Current
'	    			
'		    			strLastCategoryTOTAL = Get_Dictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHTAX")
'		    			
'		    			strNewCategoryTOTAL = GetNewConsolidatedValue(strLastCategoryTOTAL,sCurrentCategoryTotal) 				
'		    			
'		    			Call Update_Dictionary(objDictReceiptData,sCategoryName&"_CATEGORYTOTALWITHTAX",strNewCategoryTOTAL)
'			    		
'	    			
'	    			ElseIf instr(1,strTag,"FUELDETAILS")<>0 Then
'	    			
'	    				sFuelItemFlag = True
'	    			
'	    			End If
'	    			
'	    			
'	    		ElseIf instr(1,strTag,"TAXDETAILS")<>0 and sPumpTestFlag = False Then
'	    		
'	    			sTaxValues = split(strTagVal,"||")
'	    			
'	    			'ITEM NAME (ITEMWISE)
'	    			
'	    			sTaxItemName = sTaxValues(0)
'	    			
'	    			'TAX CATEGORY (ITEMWISE)
'	    			
'	    			sTaxCategory = sTaxValues(1)
'	    			
'	    			Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sTaxItemName&"_TAXCATEGORY",sTaxCategory) 'New Current
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sItemName&"_TAXCATEGORY",sTaxCategory)
'	    			
'	    			'TAX NAME (ITEMWISE)
'	    			
'	    			sTaxName = trim(sTaxValues(2))&"%"
'	    			
'	    			Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sTaxItemName&"_TAXNAME",sTaxName) 'New Current
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sItemName&"_TAXNAME",sTaxName)
'	    			
'	    			'TAX TOTAL (ITEMWISE)
'	    			
'	    			sCurrentTaxTotal = sTaxValues(3)
'	    			
'	    			Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sTaxItemName&"_TAXTOTAL",sCurrentTaxTotal) 'New Current
'	    			
'	    			strLastTaxTOTAL = Get_Dictionary(objDictReceiptData,sTaxItemName&"_TAXTOTAL")
'	    			
'	    			strNewTaxTOTAL = GetNewConsolidatedValue(strLastTaxTOTAL,sCurrentTaxTotal)  				
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sTaxItemName&"_TAXTOTAL",strNewTaxTOTAL)
'	    			
'	    			'TAX TOTAL (TAXNAMEWISE)
'	    			
'	    			sTaxName = trim(sTaxValues(2))&"%"
'	    			
'	    			If sCRItemFlag = True Then
'	    			
'	    				sTaxNameCategoryTag = sTaxName&"_CR_TAXNAMETOTAL"
'	    			
'	    			ElseIf sFuelItemFlag = True Then
'	    				
'	    				sTaxNameCategoryTag = sTaxName&"_FUEL_TAXNAMETOTAL"
'	    			Else
'	    			
'	    				sTaxNameCategoryTag = sTaxName&"_NAMOSVAT_TAXNAMETOTAL"
'	    				
'	    			End If
'	    			
'	    			strLastTaxNameTOTAL = Get_Dictionary(objDictReceiptData,sTaxNameCategoryTag)
'	    			
'	    			strNewTaxNameTOTAL = GetNewConsolidatedValue(strLastTaxNameTOTAL,sCurrentTaxTotal)  				
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sTaxNameCategoryTag,strNewTaxNameTOTAL)
'	    			
'	    			
'	    			'TAX TOTAL (ALL TRANSACTIONS)
'	    			
'	    			strLastAllTaxTOTAL = Get_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TAXTOTAL")
'	    			
'	    			strNewAllTaxTOTAL = GetNewConsolidatedValue(strLastAllTaxTOTAL,sCurrentTaxTotal)  				
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TAXTOTAL",strNewAllTaxTOTAL)
'
'					'TOTAL WITHOUT TAX (ITEMWISE)
'	    			
'	    			sItemNameTOTAL = Get_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sTaxItemName&"_TOTALWITHTAX")
'	    			
'	    			strTOTALWITHOUTTAX =  cdbl(sItemNameTOTAL) - cdbl(sCurrentTaxTotal)
'	    			
'	    			If sPriceOverRideFlag = True Then
'						strTOTALWITHOUTTAX = cdbl(strTOTALWITHOUTTAX) + cdbl(sPriceOverRideValue)
'					End If
'					
'					Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&sTaxItemName&"_TOTALWITHOUTTAX",strTOTALWITHOUTTAX) 'New Current
'					
'					strLastTaxWithOutTOTAL = Get_Dictionary(objDictReceiptData,sTaxItemName&"_TOTALWITHOUTTAX")					
'					
'					strNewTaxWithOutTOTAL = GetNewConsolidatedValue(strLastTaxWithOutTOTAL,strTOTALWITHOUTTAX)  	   				
'
'	    			Call Update_Dictionary(objDictReceiptData,sTaxItemName&"_TOTALWITHOUTTAX",strNewTaxWithOutTOTAL) 
'	    			
'	    			'CATEGORY TOTAL WITHOUT TAX (ITEMWISE)
'	    			
'	    			If sCRItemFlag = True Then
'	    				
'	    				'sCategoryName = sValues(4)
'	  					
'	  					sCurrentCategoryTOTALWITHOUTTAX = strTOTALWITHOUTTAX
'	    				
'	    				strCurrentCategoryName = Get_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_CATEGORYNAME")
'	    				
'	    				Call Update_Dictionary(CurrentItemReceiptDetails,"CURRENTRECEIPT_"&strCurrentCategoryName&"_CATEGORYTOTALWITHOUTTAX",sCurrentCategoryTotal) 'New Current
'	    			
'		    			strLastCategoryTOTALWITHOUTTAX = Get_Dictionary(objDictReceiptData,strCurrentCategoryName&"_CATEGORYTOTALWITHOUTTAX")
'		    			
'		    			strNewCategoryTOTALWITHOUTTAX = GetNewConsolidatedValue(strLastCategoryTOTALWITHOUTTAX,sCurrentCategoryTOTALWITHOUTTAX) 				
'		    			
'		    			Call Update_Dictionary(objDictReceiptData,strCurrentCategoryName&"_CATEGORYTOTALWITHOUTTAX",strNewCategoryTOTALWITHOUTTAX) 
'	    			
'	    				sCRItemFlag = False
'	    				
'	    			End If
'	    				    			
'	    			'TOTAL WITHOUT TAX (ALL TRANSACTIONS)
'	    			
'	    			strLastAllTaxWithOutTOTAL = Get_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHOUTTAX")
'	    			
'	    			strNewAllTaxWithOutTOTAL = GetNewConsolidatedValue(strLastAllTaxWithOutTOTAL,strTOTALWITHOUTTAX)  
'
'					Call Update_Dictionary(objDictReceiptData,"ALLTRANSACTIONS_TOTALWITHOUTTAX",strNewAllTaxWithOutTOTAL)
'					
'	    			
'	    		ElseIf (instr(1,strTag,"CASHTOTAL") <> 0 or instr(1,strTag,"OTHERCARD1TOTAL") <> 0 or instr(1,strTag,"OTHERCARD2TOTAL") <> 0 or instr(1,strTag,"LOCALACCOUNTTOTAL") <> 0 or instr(1,strTag,"MINIPOSMANUALTOTAL") <> 0 or instr(1,strTag,"CROSSOVERTOTAL") <> 0 or instr(1,strTag,"DRIVEOFFTOTAL") <> 0  or instr(1,strTag,"LOITOTAL") <> 0) Then	
'		
'    				'MOP TOTAL (MOPWISE)	    				
'    				
'    				sCurrentTag = strTag
'    				
'    				sCurrentMOPTotal = strTagVal
'    				
'    				strLastMOPTOTAL = Get_Dictionary(objDictReceiptData,sCurrentTag)
'	    			
'	    			strNewMOPTOTAL = GetNewConsolidatedValue(strLastMOPTOTAL,sCurrentMOPTotal)  
'	    			
'	    			Call Update_Dictionary(objDictReceiptData,sCurrentTag,strNewMOPTOTAL) 
'	    			
'	    	'-----------------------------------------------------------------------------------------------		
	    			If instr(1,strTag,"FUELDETAILS")<>0 Then
	    				bFuelFlag = True
	    				
	    				strTagValT = split(strTagVal,"||")(0)
	    				
						If strTagValT="V-Power Nitro+ DSL" Then
	    					strTagValT="Shell V-Power Diesel"
	    				else
	    					strTagValT=strTagValT
	    				End If
	    				
	    				
		    			If objDictQty.Exists (strTagValT)=False then
		    				Call Update_Dictionary(objDictQty,strTagValT,strCurrentFuelQty) 
		    			Else
		    				
		    				strCQty = get_dictionary(objDictQty,strTagValT)
		    				strFinVol = cdbl(strCQty) + cdbl(strCurrentFuelQty)
		    				Call Update_Dictionary(objDictQty,strTagValT,strFinVol) 
		    			
		    			End if
	    			
	    			else

			    			if objDictQty.Exists (strTagVal)=False then
			    				Call Update_Dictionary(objDictQty,strTagVal,1) 
			    			else
			    				strCQty = get_dictionary(objDictQty,strTagVal)
			    				strCQty = strCQty + 1
			    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
			    			End if
'			    		else
'			    			if objDictQty.Exists (strTagVal)=False then
'			    				Call Update_Dictionary(objDictQty,strTagVal,1) 
'			    			else
'			    				strCQty = get_dictionary(objDictQty,strTagVal)
'			    				strCQty = strCQty + 1
'			    				Call Update_Dictionary(objDictQty,strTagVal,strCQty) 
'			    			End if
'			    		End if
	    			End if
	    			
	    			strQTYKeyItem1 = strTagVal & "_" & "QTY"
		    		strTOTALKeyItem2 = strTagVal & "_" & "TOTAL"
		    		strVATKeyItem3 = strTagVal & "_" & "VAT"
		    		strPerItemKeyItem4 = strTagVal & "_" & "PerItemCost" ''''Delcy added
		    		strTAXKeyItem5 = strNewTranID & "_" & "TAXPERCENT"
		    		
		    		'SHAFI ADDED TESTING----------------
		    		strMOP = "@MOP@"
		    		strQTYKeyItemMOP1 = strTagVal & "_" & strMOP & "_" & "QTY"
		    		strTOTALKeyItemMOP2 = strTagVal & "_" & strMOP & "_" & "TOTAL"
		    		strVATKeyItemMOP3 = strTagVal & "_" & strMOP & "_" & "VAT"	    		
		    		'SHAFI ADDED TESTING----------------
		    		
		    		sInitialValue=""
		    		
		    		if objDictReceiptData.Exists (strQTYKeyItem1)=False then
		    			if bFuelFlag = False then
		    				Call Add_Dictionary(objDictReceiptData,strQTYKeyItem1,1) 
		    				'SHAFI ADDED TESTING----------------		
		    				Call Add_Dictionary(objDictReceiptData,strQTYKeyItemMOP1,1) 
		    				'SHAFI ADDED TESTING----------------
		    			else
		    				strCQty = get_dictionary(objDictQty,strTagVal)
		    				Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
		    				'SHAFI ADDED TESTING----------------
		    				Call Update_Dictionary(objDictReceiptData,strQTYKeyItemMOP1,strCQty)
		    				'SHAFI ADDED TESTING----------------
		    			End if
		    		else
		    			strCQty = get_dictionary(objDictQty,strTagVal)
		    			Call Update_Dictionary(objDictReceiptData,strQTYKeyItem1,strCQty)
		    			'SHAFI ADDED TESTING----------------
		    			Call Update_Dictionary(objDictReceiptData,strQTYKeyItemMOP1,strCQty)
		    			'SHAFI ADDED TESTING----------------
		    		End if
			    	
		    		if objDictReceiptData.Exists (strTOTALKeyItem2)=False then
		    			Call Add_Dictionary(objDictReceiptData,strTOTALKeyItem2,sInitialValue) 
		    			'SHAFI ADDED TESTING----------------
		    			Call Add_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,sInitialValue) 
		    			'SHAFI ADDED TESTING----------------
		    		End if
		    		if objDictReceiptData.Exists (strVATKeyItem3)=False then
		    			Call Add_Dictionary(objDictReceiptData,strVATKeyItem3,sInitialValue)
						'SHAFI ADDED TESTING----------------
		    			Call Add_Dictionary(objDictReceiptData,strVATKeyItemMOP3,sInitialValue) 
		    			'SHAFI ADDED TESTING----------------		    			
		    		End if
		    	ElseIf instr(1,strTag,"CostItemName_")<>0 Then
	    			PriceOverrideFlag = True
	    		End If
	    		
	    		If instr(1,strTag,"PumpTestItem")<> 0 Then
	    			bPumpFlagFlag = True
	    		End If
				
	    		If instr(1,strTag,"TYPE")<> 0 Then
	    			If instr(1,strTagVal,"Paid-in")<> 0 or instr(1,strTagVal,"Paid-out")<> 0 Then
	    				bPayInOutFlag = True
	    			End If
	    		End If
	    		
	    		If instr(1,strTag,"TOTALNET") <> 0 Then
	    			strCurrentNetSale = strTagVal
	    			
	    			'''''' PERITEMCOST LOGIC
	    		elseIf instr(1,strTag,"PerItemCost_") <> 0 Then
	    			If objDictReceiptData.Exists (strPerItemKeyItem4)=False Then
	    				'strLastIndItemTotal = Get_Dictionary(objDictReceiptData,strPerItemKeyItem4)
	    				Call Update_Dictionary(objDictReceiptData,strPerItemKeyItem4,strTagVal)
	    				'Saves all the items
	    				
	    				'ItemsPresent = split(strPerItemKeyItem4,"_PerItemCost")
	    				
	    				If strPerItemKeyItem4 = "" Then
								strAlreadyChecked  = strVATKeyItem3
						elseif instr(1,strAlreadyChecked,strVATKeyItem3) =0 then
								strAlreadyChecked = strAlreadyChecked & strVATKeyItem3
						End If

	    			End If
				elseIf instr(1,strTag,"TAX_") <> 0 Then
	    			'If objDictReceiptData.Exists (strTAXKeyItem5)=False Then
	    				'strLastIndItemTotal = Get_Dictionary(objDictReceiptData,strPerItemKeyItem4)
	    				strTagValPercent = cdbl(replace(strTagVal,"%",""))
	    				Call Update_Dictionary(objDictReceiptData,strTAXKeyItem5,strTagValPercent)
    			
	    		elseIf instr(1,strTag,"_TAXCAT") <> 0 Then
					strCategory = Mid(strTag,1,1)
					strITEM_NAME = strTagVal		
					Call Update_Dictionary(objDictVATCAT,strCategory,strITEM_NAME) 

				elseIf instr(1,strTag,"Litres") <> 0 Then
	    			strCurrentFuelQty = trim(replace(strTagVal,"litre",""))
	    			bFuelFlag = True		    			
		    	elseIf instr(1,strTag,"FUELTOTAL") <> 0 Then
	    			strCurrentFuelTotal =strTagVal
	    			bFuelFlag =True
	    			'Item wise Total calculation
    				strLastFuelTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
    				If strLastFuelTotal="" Then
    					strNewFuelTotal = cdbl(strCurrentFuelTotal)
    				else
'    					If strCurrentFuelTotal = objDictReceiptData.Item ("VOUCHER") Then
'    						strNewFuelTotal = cdbl(strLastFuelTotal)+cdbl(strCurrentFuelTotal)+cdbl(strCurrentFuelTotal)
'    					Else
    					strNewFuelTotal = cdbl(strLastFuelTotal)+cdbl(strCurrentFuelTotal)
'    					End If
    				End If	    				
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewFuelTotal) 
    				'SHAFI ADDED TESTING----------------
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,strNewFuelTotal) 
    				'SHAFI ADDED TESTING----------------
    				'All fuel total calculation-shabana
    				
    			If bmixedsalesflag = False and bPumpFlagFlag = False and  bPayInOutFlag = False  Then
    				iFuelCount=iFuelCount + 1	
						
						strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
		    			If strCurrentTotalSales="" Then
		    				
		    				strFuelTotalSale = strCurrentFuelTotal
						Else
		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
		    				End if
						objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
		    			strFuelTotalSale=""
		    			
		    			
						strPreviousQty = objDictReceiptData.Item ("FUELQTY")
		    			If strPreviousQty="" Then
		    				strNewFuelSaleQty = strCurrentFuelQty
		    			else
							strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				End If
		    			objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
		    			strNewFuelSaleQty = ""
		    			
				End If
    				NoTotalCost = True
		    	elseIf instr(1,strTag,"CRCOST") <> 0 and bFuelFlag =True Then
		    	
		    		'Item wise Total calculation
		    		bMixedSalesFlag = True

					If bMixedSalesFlag = True Then
		    			strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
		    			strPreviousQty = objDictReceiptData.Item ("FUELQTY")
		    			'iFuelCount=iFuelCount - 1	
		    			strFuelTotalSale = cdbl(strCurrentTotalSales) - cdbl(strCurrentFuelTotal)
		    			strNewFuelSaleQty = cdbl(strPreviousQty) - cdbl(strCurrentFuelQty)
		    		End If
		    		objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
		    		strFuelTotalSale=""
		    		objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
		    		strNewFuelSaleQty = ""

    				strLastCRTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
    				If strLastCRTotal="" Then
    					strNewCRTotal = cdbl(strTagVal)
    				else
    					strNewCRTotal = cdbl(strLastCRTotal)+cdbl(strTagVal)
    				End If	    				
					objDictReceiptData.Item ( "CRCOST") = strNewCRTotal
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewCRTotal) 
    				
    				'SHAFI ADDED TESTING----------------
    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,strNewFuelTotal) 
    				'SHAFI ADDED TESTING----------------
    				
    				'MIXED FUEL DATA TOTAL
    				strCurrentTotalSales = objDictReceiptData.Item ("MIXEDFUEL_TOTAL")
    				
	    			If strCurrentTotalSales="" Then
	    				strMixTotalSale = strCurrentFuelTotal
	    				iMixedCount=iMixedCount + 1
	    			else
	    				strMixTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
	    				iMixedCount=iMixedCount + 1
	    			End If
	    			objDictReceiptData.Item ("MIXEDFUEL_TOTAL") = strMixTotalSale
	    			strMixTotalSale=""
	    			strPreviousQty = objDictReceiptData.Item ("MIXEDFUEL_QTY")
	    			If strPreviousQty="" Then
	    				strNewMixSaleQty = strCurrentFuelQty
	    			else
	    				strNewMixSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
	    			End If
	    			objDictReceiptData.Item ("MIXEDFUEL_QTY") = strNewMixSaleQty
	    			bFuelFlag =False
    				
		    	elseIf instr(1,strTag,"CURRENCY") <> 0 Then
					
					'Calculating Rounding					
					strRoundingTagVal = split(strTagVal," ")
		    		
		    		strRoundingVal = strRoundingTagVal(ubound(strRoundingTagVal))
		    		
		    		If instr(strRoundingVal,"Rand") Then
		    			strRoundingT =empty
		    		ElseIf instr(strRoundingVal,"R") and instr(strRoundingVal,".") and instr(strRoundingVal,"-")=0 Then
		    			
		    			strRoundingT = replace(strRoundingVal,"R","")
		    			strRoundingTT = split(strRoundingT,".")
		    			strRoundingTTT = strRoundingTT(1)
		    			
		    			strCurrentTransactionTotalT = split(strCurrentTransactionTotal,".")
		    			strCurrentTransactionTotalTT = strCurrentTransactionTotalT(1)
		    					    			
		    			strTotalRoundingVal = objDictReceiptData.Item ("ROUNDING")
		    		
			    		If strTotalRoundingVal="" Then
			    			strTotalRoundingVal = (strCurrentTransactionTotalTT - strRoundingTTT)/100
						Else		    		
			    			strTotalRoundingVal = cdbl(strTotalRoundingVal) + (strCurrentTransactionTotalTT - strRoundingTTT)/100		
			    		End If
			    		
			    		objDictReceiptData.Item ("ROUNDING")=strTotalRoundingVal
			    			
		    		End If
			    		
					If bFuelFlag =True Then
						'FUEL DATA TOTAL
			    		
		    			strCurrentTotalSales = objDictReceiptData.Item ("FUELTOTAL")
		    			If strCurrentTotalSales="" Then
		    				iFuelCount=iFuelCount + 1
		    				strFuelTotalSale = strCurrentFuelTotal
		    			
		    			Else 
						    If strTagVal="Voucher" Then
'						    iFuelCount=iFuelCount + 3
'		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal) +  cdbl(strCurrentFuelTotal)
'		    				objDictReceiptData.Item ("VOUCHER") = cdbl(strCurrentFuelTotal)
							 iFuelCount=iFuelCount + 1'3
						    
		    			
						strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
		    				strTOTALKeyVoucher = strTOTALKeyItem2&"Voucher"
		    				Call Add_Dictionary(objDictVoucher,strTOTALKeyVoucher,strCurrentFuelTotal)


		    				objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
		    				objDictReceiptData.Item ("VOUCHER") = cdbl(strCurrentFuelTotal)
			    			strFuelTotalSale=""
			    			VoucherFlag = True
			    			'-----------------------------Shabana Testing---------------------------Commented
'		    				Else
'		    				iFuelCount=iFuelCount + 1

'		    				strFuelTotalSale = cdbl(strCurrentTotalSales) + cdbl(strCurrentFuelTotal)
							'-----------------------------Shabana Testing---------------------------Commented									
		    				End if
		    				
		    			End If
'		    			objDictReceiptData.Item ("FUELTOTAL") = strFuelTotalSale
'		    			strFuelTotalSale=""
		    			strPreviousQty = objDictReceiptData.Item ("FUELQTY")
		    			If strPreviousQty="" Then
		    				strNewFuelSaleQty = strCurrentFuelQty
		    			else
		    				If strTagVal="Voucher" Then
		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				strTOTALKeyVoucherQty = strTOTALKeyItem2&"VoucherQty"
		    				Call Add_Dictionary(objDictVoucher,strTOTALKeyVoucherQty,strCurrentFuelQty)



		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
'		    				Else
'		    				strNewFuelSaleQty = cdbl(strPreviousQty) + cdbl(strCurrentFuelQty)
		    				End If
		    			End If
		    			'objDictReceiptData.Item ("FUELQTY") = strNewFuelSaleQty
	'	    		End if

		    	End If
		    	
		    	elseIf strcomp(strTag,"TOTAL") = 0 Then
		    		strCI=objDictReceiptData.item("CURRENTITEM")
'		    		If instr(1,strCI,"Diesel Extra")<>0 Then
'		    					''msgbox "CUrrent Total -"&strTagVal 			
'		    		End If
		    		strCurrentTransactionTotal = strTagVal
	    			strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
	    			'Shift Total transaction calculation
	    			if bPumpFlagFlag = False and bPayInOutFlag = False  Then
	    			If strCurrentTotalSales="" Then
	    				strNewTotalSale = strTagVal
	    			else
	    				strNewTotalSale = cdbl(strCurrentTotalSales) + cdbl(strTagVal)
	    			End If
	    			
	    			objDictReceiptData.Item ( "TOTAL") = strNewTotalSale
	    			
	    			If NoTotalCost = False Then
		    			'Item wise Total calculation
	    				strLastTotal = Get_Dictionary(objDictReceiptData,strTOTALKeyItem2)
	    				If strLastTotal="" Then
	    					strNewTotal = cdbl(strTagVal)
	    				else
	    					strNewTotal = cdbl(strLastTotal)+cdbl(strTagVal)
	    				End If	    				
	    				Call Update_Dictionary(objDictReceiptData,strTOTALKeyItem2,strNewTotal) 
	    				strRefundCashtotal = strTagVal	    				
	    				'SHAFI ADDED TESTING----------------
    					Call Update_Dictionary(objDictReceiptData,strTOTALKeyItemMOP2,strRefundCashtotal) 
    				'SHAFI ADDED TESTING----------------
	    			End if
	    			End If
'	    		elseIf instr(1,strTag,"TAXAMOUNT") <> 0 Then
'	    			strCategory1 = Mid(strTag,len(strTag))
'					strITEM_NAME1 = Get_Dictionary(objDictVATCAT,strCategory1)
'    				strTemp = strVATKeyItem3
'	    			'Item wise Total calculation
'    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem3)
'    				If strLastVAT="" Then
'    					strNewVATTotal = cdbl(strTagVal)
'    				else
'    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
'    				End If	    				
'    				Call Update_Dictionary(objDictReceiptData,strVATKeyItem3,strNewVATTotal) 
'		    		strVATKeyItem3 = strTemp

				elseIf instr(1,strTag,"TAX_") <> 0 Then
					
					strCategory = Mid(strTag,len(strTag))
					strITEM_NAME = Get_Dictionary(objDictVATCAT,strCategory)
    				
    				strAlreadyCheckedTemp = strAlreadyChecked
    				If instr(strAlreadyCheckedTemp,"VAT")<> 0 Then
    				strAlreadyCheckedTemp = split(strAlreadyCheckedTemp,"_VAT")
					Else
					strAlreadyCheckedTemp = split(strAlreadyCheckedTemp," ")					
    				End If
    				
    				
    				strTemp = strVATKeyItem
   				
   				For ihead = 0 To ubound(strAlreadyCheckedTemp)
   				strITEM_NAME = strAlreadyCheckedTemp(ihead)	
   					strTagVal = cdbl(replace(strTagVal,"%",""))
    				If strITEM_NAME="V-Power Nitro+ DSL" Then
	    				strITEM_NAME="Shell V-Power Diesel"
	    			else
	    				strITEM_NAME=strITEM_NAME
	    			End If
	    			strVATKeyItem = strITEM_NAME & "_TAXPERCENT"
	    			
	    			'Item wise Total calculation
    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem)
    				If strLastVAT="" Then
    					strNewVATTotal = cdbl(strTagVal)
    				else
    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
    				End If	    				
					Call Update_Dictionary(objDictReceiptData,strVATKeyItem,strNewVATTotal)

					'SHAFI ADDED TESTING----------------
		    		Call Update_Dictionary(objDictReceiptData,strVATKeyItemMOP,strNewVATTotal) 
		    		'SHAFI ADDED TESTING----------------
		    		Next	

		    		strVATKeyItem = strTemp
					elseIf instr(1,strTag,"TAXAMOUNT") <> 0 Then
	    			'Item wise Total calculation
    				strLastVAT = Get_Dictionary(objDictReceiptData,strVATKeyItem3)
    				If strLastVAT="" Then
    					strNewVATTotal = cdbl(strTagVal)
    				else
    					strNewVATTotal = cdbl(strLastVAT)+cdbl(strTagVal)
    				End If	    				
    				Call Update_Dictionary(objDictReceiptData,strVATKeyItem3,strNewVATTotal) 
					'SHAFI ADDED TESTING----------------
		    		Call Update_Dictionary(objDictReceiptData,strVATKeyItemMOP3,strNewVATTotal) 
		    		'SHAFI ADDED TESTING----------------
					strVATKeyItem3 = strTemp
		    	elseIf instr(1,strTag,"MOP") <> 0 Then
		    		'strCurrentTranTotal = objDictReceiptData.Item ("TOTAL")
		    		strCurrentTranTotal = strCurrentTransactionTotal
		    		
		    		'SHAFI ADDED LOGIC TESING-------------
		    		strCI=objDictReceiptData.item("CURRENTITEM")
		    		MOPSearch = strCI & "_@MOP@_" 
		    		
		    		For each MopKey in objDictReceiptData.keys
		    			strMopTemp1=Get_Dictionary(objDictReceiptData,MopKey)
		    			
		    			If instr(1,MopKey,MOPSearch)<>0 Then
		    				
		    				if instr(1,MopKey,"_QTY") then
			    				NewMOPKey = strCI & "_" & strTagVal & "_QTY" 
			    			ElseIf instr(1,MopKey,"_TOTAL") then 
			    				NewMOPKey = strCI & "_" & strTagVal & "_TOTAL" 
			    			ElseIf instr(1,MopKey,"_VAT") then 
			    				NewMOPKey = strCI & "_" & strTagVal & "_VAT" 
			    			End if
			    			Call Add_Dictionary(objDictReceiptData,NewMOPKey,strMopTemp1) 
		    			End If
		    			
		    		Next
		    	
		    		'SHAFI ADDED LOGIC TESING-------------
		    		Select Case strTagVal
		    			Case "CASH"
		    				strTotalCashSales = objDictReceiptData.Item ("CASH")
		    				
		    				If strRoundingT<>0 Then
			    				strCurrentTranTotal = strRoundingT
								strRoundingT = empty    				
			    			End If		    				
		    				
		    				If VoucherFlag =False and bPumpFlagFlag = False Then
		    					If strTotalCashSales="" Then
			    					strCashTotal = strCurrentTranTotal
				    			else
				    				strCashTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCashSales)
				    			End If	
		    				objDictReceiptData.Item ("CASH") = strCashTotal
			    			strCashTotal=""	    				
		    				Else
			    				VoucherFlag =False 
		    				End If
		    				
							If strCurrentNetSale>0 and bPumpFlagFlag = False Then
								stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    				'stCashVATTotal = strNewNetSale
			    				If stTotalOfNetSales = "" Then		    					
			    					stTotalOfNetSales = strCurrentNetSale
			    				Else
			    					stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(strCurrentNetSale)		    						
			    				End If
				    			'stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stCashTotalNetSales)
				    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales	
							End If
		    				
			    				
		    			Case "DRIVEOFF"
		    				
		    				strTotalDOSales = objDictReceiptData.Item ("DRIVEOFF")
		    				If strTotalDOSales="" Then
		    					strDOTotal = strCurrentTranTotal
			    			else
			    				strDOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalDOSales)
			    			End If
			    			objDictReceiptData.Item ("DRIVEOFF") = strDOTotal
							
							'Total Net Sales
			    			stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    			stDOVATTotal = (strDOTotal/1.15)*(0)
			    			stDOTotalNetSales = cdbl(strDOTotal) - cdbl(stDOVATTotal)
			    			stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stDOTotalNetSales)
			    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales
			    			strDOTotal=""
		    			
		    			Case "LOI"
		    				strTotalLOISales = objDictReceiptData.Item ("LOI")
		    				If strTotalLOISales="" Then
		    					strLOITotal = strCurrentTranTotal
			    			else
			    				strLOITotal = cdbl(strCurrentTranTotal) + cdbl(strTotalLOISales)
			    			End If
			    			objDictReceiptData.Item ("LOI") = strLOITotal
							
							'Total Net Sales
			    			stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    			stLOIVATTotal = (strLOITotal/1.15)*(0.15)
			    			stLOITotalNetSales = cdbl(strLOITotal) - cdbl(stLOIVATTotal)
			    			stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stLOITotalNetSales)
			    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales			    			
			    			
			    			strLOITotal=""
		    			Case "CROSSOVER"
		    				strTotalCOSales = objDictReceiptData.Item ("CROSSOVER")
		    				If strTotalCOSales="" Then
		    					strCOTotal = strCurrentTranTotal
			    			else
			    				strCOTotal = cdbl(strCurrentTranTotal) + cdbl(strTotalCOSales)
			    			End If
							objDictReceiptData.Item ("CROSSOVER") = strCOTotal		    			
			    			
			    			'Total Net Sales
			    			stTotalOfNetSales = objDictReceiptData.Item ("NETSALES")
			    			stCOVATTotal = (strCOTotal/1.15)*(0)
			    			stCOTotalNetSales = cdbl(strCOTotal) - cdbl(stCOVATTotal)
			    			stTotalOfNetSales = cdbl(stTotalOfNetSales) + cdbl(stCOTotalNetSales)
			    			objDictReceiptData.Item ("NETSALES") = stTotalOfNetSales	
			    			
			    			strCOTotal=""
			    			
			    		Case "REFUND","REFUND AS REVERSAL"
		    				If PriceOverrideFlag = True  Then
								
								'Subtracting the price override reversal items from the CASHTOTAL TOTAL
			    				TotalCashVal = objDictReceiptData.Item ("CASH")
			    				If TotalCashVal = "" Then
			    					TotalCashValueT = cdbl(0) - cdbl(sCurrentMOPVal)
			    				Else
			    					TotalCashValueT = cdbl(TotalCashVal)-cdbl(sCurrentMOPVal)
			    				End If			    				
								objDictReceiptData.Item ("CASH") = TotalCashValueT
																
								'Subtracting the price override reversal items from the NETSALES TOTAL
								TotalNetSalesValue = objDictReceiptData.Item ("NETSALES")
								If TotalNetSalesValue = "" Then
									TotalNetSalesValueT = cdbl(0) - cdbl(strCurrentItemNetSales)								
								Else
									TotalNetSalesValueT = cdbl(TotalNetSalesValue)-cdbl(strCurrentItemNetSales)
								End If																
                                objDictReceiptData.Item ("NETSALES") = TotalNetSalesValueT
								
			    			End If
			    			
		    				If bFuelFlag = True and bPumpFlagFlag = False Then 
		    					strCurrentTotalSales = objDictReceiptData.Item ("MIXEDFUEL_TOTAL")
		    					
		    					If strCurrentTotalSales = "" Then
		    						strFRefundTotal = cdbl(strCurrentFuelTotal)
		    					Else
		    						strFRefundTotal = cdbl(strCurrentFuelTotal) +  cdbl(strCurrentTotalSales)
		    					End If	    					
		    					'objDictReceiptData.Item ("MIXEDFUEL_TOTAL")=strFRefundTotal
		    					
		    					strCurrentTotalQty = objDictReceiptData.Item ("MIXEDFUEL_QTY")
		    					'strFRefundQty = cdbl(strCurrentFuelQty) +  cdbl(strCurrentTotalQty)
		    					If strCurrentTotalQty = "" Then
		    						strFRefundQty = cdbl(strCurrentFuelQty)
		    					Else
		    						strFRefundQty = cdbl(strCurrentFuelQty) +  cdbl(strCurrentTotalQty)
		    					End If
		    					'objDictReceiptData.Item ("MIXEDFUEL_QTY")=strFRefundQty
		    				    					
		    					
		    					strTOTALKeyRefundFuel = strTOTALKeyItem2&"RefundFuel"

		    					Call Add_Dictionary(objDictVoucher,strTOTALKeyRefundFuel,strCurrentFuelTotal)
		    					
		    					strTOTALKeyRefundFuelQty = strTOTALKeyItem2&"RefundFuelQty"
		    					Call Add_Dictionary(objDictVoucher,strTOTALKeyRefundFuelQty,strCurrentFuelQty)
		    	
		    					
		    				End If
		    		End Select
		    	End if
    			
	    	else
	    		strTranData = strTranData & "," & strTag & ":" & strTagVal
	    		bFlagNext = False
	'    		objDictReceiptData.Item(strNewTranID) = strTranData
	    		Update_Dictionary objDictReceiptData,strNewTranID,strTranData
	    		strNewTranID=""
	    		strTranData=""
	  			bPumpFlagFlag = False
	  			VoucherFlag =False
	    		'Set dictionary tax category to null
	    		Set objDictVATCAT = Nothing
	    		Set objDictVATCAT = CreateObject("Scripting.Dictionary")
	    	End if  '21
	    		    
	    End If '1
	    'bFuelFlag = False
	Next
	objDictReceiptData.Item ("MIXEDCOUNT") = iMixedCount
	objDictReceiptData.Item ("FUELCOUNT") = iFuelCount
	
	'Calculate the total taxes paid from net and total sales
	strCurrentNetSales = objDictReceiptData.Item ( "NETSALES")
	strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
	strCurrentTotalSales = objDictReceiptData.Item ( "TOTAL")
	
	If strCurrentNetSales<>"" Then
		strTaxesPaid = cdbl(strCurrentTotalSales)-cdbl(strCurrentNetSales)
		objDictReceiptData.Item ("TOTALTAXES") = round(strTaxesPaid,2)
	End If
	
	
End Function
'
'Public Function CreateBOSInstructionFile()
'	
'	Dim strStatusUpdate
'	
'	Set FileSysObj = CreateObject("Scripting.FileSystemObject")
'	Set dataFile = FileSysObj.OpenTextFile(BOSInstructionFile,2,true)
'	
'	If objDictBOSExecInstruction.count <> 0 Then
'		
'		For each key in objDictBOSExecInstruction.keys
'			
'			strText = objDictBOSExecInstruction.item(key)
'			strLine = key & "-" & strText & vbCrLf
'			dataFile.Write strLine
'		Next
'		
'	End If
'	
'	dataFile.Close
'	
'	Set FileSysObj = Nothing
'	Set dataFile = Nothing
'
'
'End Function
'
'Function GetBOSInstruction(strKey)
'	
'	Set FileSysObj = CreateObject("Scripting.FileSystemObject")
'	Set dataFile = FileSysObj.OpenTextFile(BOSInstructionFile,1)
'	
'	Do
'		strLine = dataFile.ReadLine
'		If instr(1,strLine,strKey) <> 0 Then
'			strArr = Split(strLine,"-")
'			strKeyValue = strArr(1)
'			strInstruction = strKeyValue
'			Exit do
'		End If
'	Loop While NOT dataFile.AtEndOfStream
'
'	dataFile.Close
'	Set FileSysObj = Nothing
'	Set dataFile = Nothing
'
'	GetBOSInstruction = strInstruction
'
'End Function
'

		'******************************************** HEADER ******************************************
' Name : GetTodaydate
' Description : to Get today's date
' Creator : Madhusmitta Pal
' Date :11th Nov,2017
' Last Modified On : 
' Last Modified By : 
' Input Parameter : 
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
  Function GetTodaydate()
       sdate = Date() 
    sy=year(sdate)
    sm=Month(sdate)
      If len(sm)=1 Then
      	sm="0"&sm
      End If
     
    sd=day(sdate)
      If len(sd)=1 Then
      	sd="0"&sd
      End If
    ddate= sd&"/"&sm&"/"&sy
    '''msgbox ddate
      GetTodaydate=ddate
 End Function
 
 
 		'******************************************** HEADER ******************************************
' Name : GetYesterdate
' Description : to Get yesterday's date
' Creator : Madhusmitta Pal
' Date :11th Nov,2017
' Last Modified On : 
' Last Modified By : 
' Input Parameter : 
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
  Function GetYesterdate()
       sdate = Date() 
    sy=year(sdate)
    sm=Month(sdate)
    
      If len(sm)=1 Then
      	sm="0"&sm
      End If
     
    sd=day(sdate)
    sd= sd-1
      If len(sd)=1 Then
      	sd="0"&sd
      End If
    ddate= sd&"/"&sm&"/"&sy
   ' ''msgbox ddate
      GetYesterdate=ddate
 End Function
 
 
  		'******************************************** HEADER ******************************************
' Name : GetYesterdate
' Description : to Get yesterday's date
' Creator : Madhusmitta Pal
' Date :11th Nov,2017
' Last Modified On : 
' Last Modified By : 
' Input Parameter : 
' Output Parameter : Boolean
'******************************************** HEADER ******************************************
  Function GetTomordate()
       sdate = Date() 
    sy=year(sdate)
    sm=Month(sdate)
    
      If len(sm)=1 Then
      	sm="0"&sm
      End If
     
    sd=day(sdate)
    sd= sd+1
      If len(sd)=1 Then
      	sd="0"&sd
      End If
    ddate= sd&"/"&sm&"/"&sy
   ' ''msgbox ddate
      GetTomordate=ddate
 End Function
 
 
'******************************************** HEADER ******************************************
' Name : ValidateDtWtColValue1
' Description : 
' Creator : Syed Shafi
' Date : 10th Dec,2017
' Last Modified On : 
' Last Modified By :  
' Input Parameter : sBrowser,sPage,sFrame,sObject,byref DictTbl
' Output Parameter : 
'******************************************** HEADER ******************************************
 Public Function ValidateDtWtColValue1Shabana(sBrowser,sPage,sFrame,sObject,byref DictTbl)
	
 On error resume next  

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  LastRow=2
	
'  set TableObject=FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count
	
	For Each Key in DictTbl
		If instr(1,Key,"Item" )<>0 OR instr(1,Key,"Business Date" )<>0 OR instr(1,Key,"Special Name" )<>0 or instr(1,Key,"Org Level - Dynamic" )<>0 or instr(1,Key,"Tank" )<>0 Then
			KeyItemColumn=Key
			Exit for
		End If
	next
	
	 Do
		 	For Each elem in DictTbl
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						
					 	if strcomp(elem,KeyItemColumn)=0 then
					 		counter = counter+1
						 	rowNum=TableObject.GetRowWithCellText(KeyItemColumn)
						 	colNumb=GetWTColumnNumber(TableObject,KeyItemColumn) 		
						 	rCount=TableObject.RowCount	
						 	
						 	For row = LastRow To rCount Step 1
								strCellValue = TableObject.GetCellData(row,colNumb)
								
								If instr(1,trim(strCellValue), TbleColValue ) <> 0 Then	
								   AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue)&"</B>","PASSED"	
									'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
									writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									LastRow=row
									TblrowNum=row
									DNLoopRowflag=True
									Exit For
								End If
							next
						End if
						If DNLoopRowflag=True Then
							RefreshDicObj=True
'							counter=0
							Exit for
						End If
					else
						if strcomp(elem,KeyItemColumn)<>0 then
							counter = counter+1
							colNumb=GetWTColumnNumber(TableObject,elem) 
							strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
							strCellValue2 = replace(strCellValue1,",","")
							If instr(1,TbleColValue,"litre")<>0 Then
								strCellValue1 =strCellValue1 & " litre"
							End If
							If instr(1,trim(strCellValue1),TbleColValue)<>0 Then
								'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue2),"PASSED"
								AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue2)&"</B>","PASSED"	
								writelog "","","","","Report(): Column Value matched. Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue2)
								bMatchedRow = True
							else
							    AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue2)&"</B>","FAILED"	
								'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue2),"FAILED"
								writelog "","","","","Report(): Column Value not matched!! Expected Value : " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue2)
'								bMatchedRow = False
'								DNLoopRowflag=False
'								RefreshDicObj=True
'								TblrowNum=0
'								LastRow=LastRow+1
'								counter=0
'								Exit for
							End if
						End if
					End If
		 	Next
	 		
	 		If counter=totalDictNum Then
	 			RefreshDicObj=False
	 		End If
	 		
	 Loop While (RefreshDicObj=True )
  
	If ErrNumber<>0 Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
  
  ValidateDtWtColValue1 = bMatchedRow
	
End Function

Public Function ValidateDtWtColValue1(sBrowser,sPage,sFrame,sObject,byref DictTbl)
	
' On error resume next  
bMatchedRow1=true
r=true
  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  bMatchedRow = False
  flag=False
  
  'bMatchedRow1=True
  LastRow=2
	
'  set TableObject=FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count
'	instr(1,Key,"Org Level - Dynamic" )<>0

	For Each Key in DictTbl
			If  instr(1,Key,"Activity Type" )<>0  or instr(1,Key,"Reference" )<>0 or instr(1,Key,"Date and Time" )<>0 or instr(1,Key,"Paid" )<>0 or instr(1,Key,"Sales" )<>0 or instr(1,Key,"Time" )<>0 or instr(1,Key,"Hour" )<>0 or instr(1,Key,"Trans" )<>0 or instr(1,Key,"Shift" )<>0 or instr(1,Key,"Category" )<>0  or instr(1,Key,"Timestamp" )<>0 or instr(1,Key,"Coupon" )<>0 or instr(1,Key,"Time Frequency - Dynamic" )<>0 or instr(1,Key,"Item" )<>0 OR instr(1,Key,"VAT Rate") OR instr(1,Key,"Business Date" )<>0 OR instr(1,Key,"Method of Payment") OR instr(1,Key,"Special Name" )<>0 or instr(1,Key,"Org Level - Dynamic" )<>0 or instr(1,Key,"Tank" )<>0 or instr(1,Key,"Description" )<>0 or instr(1,Key,"Tank / Business Date" )<>0 Then
			KeyItemColumn=Key
			If KeyItemColumn = "Hose: Fuel Item"  Then
				pumpNo = "2"
			End If
			Exit for
		End If
	next
'	For Each Key in DictTbl
'		If instr(1,Key,"Item" )<>0 OR instr(1,Key,"Dynamic") OR instr(1,Key,"Method of Payment") OR instr(1,Key,"Description")  OR  instr(1,Key,"Reference" )<>0 OR instr(1,Key,"Sales Range Start" )<>0 OR instr(1,Key,"VAT Rate") OR instr(1,Key,"Hose: Fuel Item")  OR instr(1,Key,"Event Timestamp") Then	
'		KeyItemColumn=Key
'		If KeyItemColumn = "Hose: Fuel Item"  Then
'			pumpNo = "2"
'		End If
'			Exit for
'		End If
'	next
	
	 Do
		 	For Each elem in DictTbl
		 			bln= false
		 			match =false
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						''''''''Delcy Added
						If elem = "Hose: Fuel Item" Then
							For Iterator = 1 To TableObject.rowcount Step 1
									
								If instr(TableObject.getcelldata(Iterator,2), DictTbl.Item(elem)) <>0 Then
																		
									If TableObject.getcelldata(Iterator,1) = pumpNo Then
										rowNumber =Iterator
							exit for
									End If
									
								End If
						
							Next
						
						End If	
						If  elem = "Tank / Business Date" Then
						
							For Iterator = 1 To TableObject.rowcount Step 1
									
								If instr(TableObject.getcelldata(Iterator,1), DictTbl.Item(elem)) <>0 Then
																		
									rNumber=Iterator
									flag=True
							exit for
									
									
								End If
						
							Next
						
						End If
						If  elem = "Time Frequency - Dynamic" Then
						
							For Iterator = 1 To TableObject.rowcount Step 1
									
								If instr(TableObject.getcelldata(Iterator,1), DictTbl.Item(elem)) <>0 Then
																		
									rNumber=Iterator
							exit for
									
									
								End If
						
							Next
						
						End If						
						''''''Delcy Added
					 	if strcomp(elem,KeyItemColumn)=0 then
					 		counter = counter+1
						 	rowNum=TableObject.GetRowWithCellText(KeyItemColumn)
						 	colNumb=GetWTColumnNumber(TableObject,KeyItemColumn) 		
						 	rCount=TableObject.RowCount	
						 	'''''Delcy Added
							If elem = "Hose: Fuel Item" Then
								rowNum = rowNumber
							ElseIf elem="Tank / Business Date" Then
								rowNum=rNumber
							ElseIf elem="Time Frequency - Dynamic" Then
								rowNum=rNumber
							End If
							'''''Delcy Added
						 	
						 	For row = LastRow To rCount Step 1
						 		If rowNumber <> "" or rNumber<>"" Then
						 			strCellValue = TableObject.GetCellData(rowNum,colNumb)
							 	else
							 		strCellValue = TableObject.GetCellData(row,colNumb)
							 	End If
'								strCellValue = TableObject.GetCellData(row,colNumb)
'								If instr(TbleColValue,".") <> 0 Then
'									tableColVal = split(TbleColValue,".")
'									'strCellVal = strCellValue(0) & "." & strCellValue(1)
'									table_length = len(tableColVal(1))
'									If table_length = 1 Then
'										TbleColValue = TbleColValue +0
''										tableColVal(0) & "." & tableColVal(1) & "0"
''									else
''										tableColVal = TbleColValue
'									End If
'								End if 
								If instr(strCellValue,".") <> 0 and instr(strCellValue,"%")=0 Then
									strCellVal = split(strCellValue,".")
									'strCellVal = strCellValue(0) & "." & strCellValue(1)
									cell_length = len(strCellVal(1))
								If mid(TbleColValue,1,1) = "R" Then
									TbleColValue = mid(TbleColValue,2)	
								End If
								If mid(strCellValue,1,1) = "R" Then
									strCellValue = mid(strCellValue,2)	
								End If
									TbleColValue=round(TbleColValue, cell_length)							
									If cstr(strCellValue) = cstr(TbleColValue) Then
										bln=true
									End If
								End If
								
								If isnumeric(mid(strCellValue, 2) )Then
									If cstr(strCellValue) = cstr(TbleColValue) Then
										bln=true
									End If
								elseif instr(1,trim(cstr(strCellValue)), cstr(TbleColValue) ) <> 0 then
										bln = true
								Else 
										bln=false
								End If	
								
								
								If bln = true Then
									'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED")	
									AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"
									'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
									'writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									writelog "Info","BOS_"& environment.Value("sCurrentDirectory") & "Generic_Function_Rp\ValidateDtWtColValue1",Key & "-" & "Expected Value :" & TbleColValue & " " & "Actual Value :" & strCellValue
									If rowNumber <> "" Then
										LastRow=rowNum
										TblrowNum=rowNum
										DNLoopRowflag=True
									elseIf rNumber <> "" and flag=true Then
										LastRow=rNumber + 1
										TblrowNum=rNumber + 1
										DNLoopRowflag=True
									elseIf rNumber <> "" Then
										LastRow=rowNum
										TblrowNum=rowNum
										DNLoopRowflag=True
									else
										LastRow=row
										TblrowNum=row
										DNLoopRowflag=True
									End if 
									Exit For
									else
									r=false
								End If
							next
							
						End if
						If DNLoopRowflag=True Then
							RefreshDicObj=True
'							counter=0
							Exit for
						End If
					else
						if strcomp(elem,KeyItemColumn)<>0 then
							counter = counter+1
							colNumb=GetWTColumnNumber(TableObject,elem) 
							strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
							
							If instr(TbleColValue,".") <> 0 Then
										TbleColValue1 = split(TbleColValue,".")
									 	'cell_leng = len(strCellValue2(1))
									 
									If right(TbleColValue1(1),2) ="00" Then
										TbleColValue=left(TbleColValue,len(TbleColValue)-2)
							
								elseIf  right(TbleColValue1(1),1) ="0" Then
										TbleColValue=left(TbleColValue,len(TbleColValue)-1)
									End If
									If right(TbleColValue,1) = "." Then
										TbleColValue = left(TbleColValue,len(TbleColValue)-1)
									End If
								End if
								''''''Delcy Added
									If instr(TbleColValue,",") <> 0 Then
										TbleColValue = replace(TbleColValue,",","")
										TbleColValue = trim(TbleColValue)
									End If
							'''''''''Delcy Added	
							If instr(strCellValue1,".") <> 0 Then
										strCellValue2 = split(strCellValue1,".")
									 	cell_leng = len(strCellValue2(1))
									If right(strCellValue2(1),4) ="0000" Then
										strCellValue1=left(strCellValue1,len(strCellValue1)-4)
									ElseIf right(strCellValue2(1),3) ="000" Then
										strCellValue1=left(strCellValue1,len(strCellValue1)-3)
									elseIf right(strCellValue2(1),2) ="00" Then
										strCellValue1=left(strCellValue1,len(strCellValue1)-2)
									elseIf  right(strCellValue2(1),1) ="0" Then
										strCellValue1=left(strCellValue1,len(strCellValue1)-1)
									End If
									If right(strCellValue1,1) = "." Then
										strCellValue1 = left(strCellValue1,len(strCellValue1)-1)
									End If
								End if
							''''''Delcy Added
							If instr(strCellValue1,",") <> 0 Then
								strCellValue1 = replace(strCellValue1,",","")
								strCellValue1 = trim(strCellValue1)
							End If
							'''''''''Delcy Added
							If mid(TbleColValue,1,1) = "R" Then
								TbleColValue = mid(TbleColValue,2)

								if TbleColValue = "" then
									tblCol =  "R"

								else
									
								tblCol = "R" & round(TbleColValue, cell_leng)			
								End If
								
								
'								If mid(TbleColValue,1,1) <> "£" Then
'									tblCol = "£" & round(mid(TbleColValue,2), cell_leng)	
'								else
'									tblCol = TbleColValue							
'								End If
								
								If mid(strCellValue1,1,1) = "R" Then
									strCellValue1 = mid(strCellValue1,2)
									cellVal = "R" & strCellValue1
								else
									cellVal	= strCellValue1						
								End If
								If TbleColValue <> "" Then
									TbleColValue=round(TbleColValue, cell_leng)
'								else
'									TbleColValue = TbleColValue
								End If							
							else
								tblCol = TbleColValue
								cellVal	= strCellValue1		
							End If

							If cstr(strCellValue1) = cstr(TbleColValue) Then
									match=true
							End If

								
								If isnumeric(mid(strCellValue1, 2))Then
									If mid(strCellValue1, 2) = TbleColValue Then
										match=true
									End If
								elseif instr(1,trim(cstr(strCellValue1)), cstr(TbleColValue) ) <> 0 then
										match = true
								Else 
										match=false
								End If	
							
							
							If instr(1,TbleColValue,"litre")<>0 Then
								strCellValue1 =strCellValue1 & " litre"
							End If
							If match = true Then
								'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal),"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &tblCol&"</B>","<B>" & trim(cellVal) &"</B>","PASSED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal),"PASSED"
								'writelog "","","","","Report(): Column Value matched. Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal)
								writelog "Info","BOS_"& environment.Value("sCurrentDirectory") & "ValidateDtWtColValue1"& err.number,Key & "-" & "Expected Value :" & tblCol & " " & "Actual Value :" & cellVal
								bMatchedRow = True
							else
'								Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal),"FAILED")
                                AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &tblCol&"</B>","<B>" & trim(cellVal) &"</B>","FAILED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal),"FAILED"
								'writelog "","","","","Report(): Column Value not matched!! Expected Value : " & tblCol & VBCRLF & "Actual Value: " & trim(cellVa)
								writelog  "Error","BOS_"& environment.Value("sCurrentDirectory") & "ValidateDtWtColValue1"& err.number,Key & "-" & "Expected Value :" & tblCol & " " & "Actual Value :" & cellVal
								bMatchedRow1=false
						 		
						 		'bMatchedRow = False
								'DNLoopRowflag=True
								'RefreshDicObj=True
'								TblrowNum=0
'								LastRow=LastRow+1
'								counter=0
'								Exit for
							End if
							
						End if
					End If
					
		 	Next
	 		
	 		If counter=totalDictNum Then
	 			RefreshDicObj=False
	 		End If
	 		
	 Loop While (RefreshDicObj=True )
	'If ErrNumber<>0 Then
	If bMatchedRow1=false Then	 	
		'writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
		ValidateDtWtColValue1=false
	else
		ValidateDtWtColValue1=true
		'writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
'  r=true
'  bMatchedRow1=true
'  ValidateDtWtColValue1 = bMatchedRow
	
End Function

Public Function ValidateDtWtColValue1ZA(sBrowser,sPage,sFrame,sObject,byref DictTbl)
	
' On error resume next  
bMatchedRow1=true
r=true
  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  bMatchedRow = False
  
  'bMatchedRow1=True
  LastRow=2
	
'  set TableObject=FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count
	
	For Each Key in DictTbl
			If instr(1,Key,"Item" )<>0 OR instr(1,Key,"Business Date" )<>0 OR instr(1,Key,"Method of Payment") OR instr(1,Key,"Special Name" )<>0 or instr(1,Key,"Org Level - Dynamic" )<>0 or instr(1,Key,"Tank" )<>0 or instr(1,Key,"Description" )<>0 Then
			KeyItemColumn=Key
			Exit for
		End If
	next
	
	 Do
		 	For Each elem in DictTbl
		 			bln= false
		 			match =false
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						
					 	if strcomp(elem,KeyItemColumn)=0 then
					 		counter = counter+1
						 	rowNum=TableObject.GetRowWithCellText(KeyItemColumn)
						 	colNumb=GetWTColumnNumber(TableObject,KeyItemColumn) 		
						 	rCount=TableObject.RowCount	
						 	
						 	For row = LastRow To rCount Step 1
								strCellValue = TableObject.GetCellData(row,colNumb)
'								If instr(TbleColValue,".") <> 0 Then
'									tableColVal = split(TbleColValue,".")
'									'strCellVal = strCellValue(0) & "." & strCellValue(1)
'									table_length = len(tableColVal(1))
'									If table_length = 1 Then
'										TbleColValue = TbleColValue +0
''										tableColVal(0) & "." & tableColVal(1) & "0"
''									else
''										tableColVal = TbleColValue
'									End If
'								End if 
								If instr(strCellValue,".") <> 0 Then
									strCellVal = split(strCellValue,".")
									'strCellVal = strCellValue(0) & "." & strCellValue(1)
									cell_length = len(strCellVal(1))
								If mid(TbleColValue,1,1) = "R" Then
									TbleColValue = mid(TbleColValue,2)	
								End If
								If mid(strCellValue,1,1) = "R" Then
									strCellValue = mid(strCellValue,2)	
								End If
									TbleColValue=round(TbleColValue, cell_length)							
									If cstr(strCellValue) = cstr(TbleColValue) Then
										bln=true
									End If
								End If
								
								If isnumeric(mid(strCellValue, 2) )Then
									If cstr(strCellValue) = cstr(TbleColValue) Then
										bln=true
									End If
								elseif instr(1,trim(cstr(strCellValue)), cstr(TbleColValue) ) <> 0 then
										bln = true
								Else 
										bln=false
								End If	
								
								
								If bln = true Then		
                                    AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"								
									'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
									writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									LastRow=row
									TblrowNum=row
									DNLoopRowflag=True
									Exit For
									else
									r=false
								End If
							next
							
						End if
						If DNLoopRowflag=True Then
							RefreshDicObj=True
'							counter=0
							Exit for
						End If
					else
						if strcomp(elem,KeyItemColumn)<>0 then
							counter = counter+1
							colNumb=GetWTColumnNumber(TableObject,elem) 
							strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
							
							If instr(TbleColValue,".") <> 0 Then
										TbleColValue1 = split(TbleColValue,".")
									 	'cell_leng = len(strCellValue2(1))
									 
									If right(TbleColValue1(1),2) ="00" Then
										TbleColValue=left(TbleColValue,len(TbleColValue)-3)
									elseIf  right(TbleColValue1(1),1) ="0" Then
										TbleColValue=left(TbleColValue,len(TbleColValue)-1)
									End If
								End if
							If instr(strCellValue1,".") <> 0 Then
										strCellValue2 = split(strCellValue1,".")
									 	cell_leng = len(strCellValue2(1))
									 
									If right(strCellValue2(1),2) ="00" Then
										strCellValue1=left(strCellValue1,len(strCellValue1)-3)
									elseIf  right(strCellValue2(1),1) ="0" Then
										strCellValue1=left(strCellValue1,len(strCellValue1)-1)
									End If
								End if
							If mid(TbleColValue,1,1) = "R" Then
								TbleColValue = mid(TbleColValue,2)

								if TbleColValue = "" then
									tblCol =  "R"

								else
									
								tblCol = "R" & round(TbleColValue, cell_leng)			
								End If
								
								
'								If mid(TbleColValue,1,1) <> "£" Then
'									tblCol = "£" & round(mid(TbleColValue,2), cell_leng)	
'								else
'									tblCol = TbleColValue							
'								End If
								
								If mid(strCellValue1,1,1) = "R" Then
									strCellValue1 = mid(strCellValue1,2)
									cellVal = "R" & strCellValue1
								else
									cellVal	= strCellValue1						
								End If
								If TbleColValue <> "" Then
									TbleColValue=round(TbleColValue, cell_leng)
'								else
'									TbleColValue = TbleColValue
								End If							
							else
								tblCol = TbleColValue
								cellVal	= strCellValue1		
							End If

							If cstr(strCellValue1) = cstr(TbleColValue) Then
									match=true
							End If

								
								If isnumeric(mid(strCellValue1, 2))Then
									If mid(strCellValue1, 2) = TbleColValue Then
										match=true
									End If
								elseif instr(1,trim(cstr(strCellValue1)), cstr(TbleColValue) ) <> 0 then
										match = true
								Else 
										match=false
								End If	
							
							
							If instr(1,TbleColValue,"litre")<>0 Then
								strCellValue1 =strCellValue1 & " litre"
							End If
							If match = true Then
								'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal),"PASSED"
								AppendTestHTMLNEW StepCounter,elem,"<B>" &tblCol&"</B>","<B>" & trim(cellVal) &"</B>","PASSED"	
								writelog "","","","","Report(): Column Value matched. Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal)
								bMatchedRow = True
							else
								'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & tblCol & VBCRLF & "Actual Value: " & trim(cellVal),"FAILED"
								AppendTestHTMLNEW StepCounter,elem,"<B>" &tblCol&"</B>","<B>" & trim(cellVal) &"</B>","FAILED"	
								writelog "","","","","Report(): Column Value not matched!! Expected Value : " & tblCol & VBCRLF & "Actual Value: " & trim(cellVa)
								bMatchedRow1=false
						 		
						 		'bMatchedRow = False
								'DNLoopRowflag=True
								'RefreshDicObj=True
'								TblrowNum=0
'								LastRow=LastRow+1
'								counter=0
'								Exit for
							End if
							
						End if
					End If
					
		 	Next
	 		
	 		If counter=totalDictNum Then
	 			RefreshDicObj=False
	 		End If
	 		
	 Loop While (RefreshDicObj=True )
	'If ErrNumber<>0 Then
	If bMatchedRow1=false Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
		ValidateDtWtColValue1=false
	else
		ValidateDtWtColValue1=true
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
'  r=true
'  bMatchedRow1=true
'  ValidateDtWtColValue1 = bMatchedRow
	
End Function


Public Function ValidateDtWtColValue1Delcy(sBrowser,sPage,sFrame,sObject,byref DictTbl)
	
' On error resume next  

bMatchedRow1=true
r=true

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  bMatchedRow = False
  'bMatchedRow1=True

  LastRow=2
	
'  set TableObject=FindWTObjectInPage(sBrowser,sPage,sFrame,sWebTable)

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count
	
	For Each Key in DictTbl
		If instr(1,Key,"Item" )<>0 OR instr(1,Key,"Business Date" )<>0 OR instr(1,Key,"Method of Payment") OR instr(1,Key,"Special Name" )<>0 or instr(1,Key,"Org Level - Dynamic" )<>0 or instr(1,Key,"Tank" )<>0 or instr(1,Key,"Description" )<>0 Then
			KeyItemColumn=Key
			Exit for
		End If
	next
	
	 Do
		 	For Each elem in DictTbl
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						
					 	if strcomp(elem,KeyItemColumn)=0 then
					 		counter = counter+1
						 	rowNum=TableObject.GetRowWithCellText(KeyItemColumn)
						 	colNumb=GetWTColumnNumber(TableObject,KeyItemColumn) 		
						 	rCount=TableObject.RowCount	
						 	
						 	For row = LastRow To rCount Step 1
								strCellValue = TableObject.GetCellData(row,colNumb)
								
								If instr(1,trim(strCellValue), TbleColValue ) <> 0 Then	
									'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
									AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue)&"</B>","PASSED"	
									writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									LastRow=row
									TblrowNum=row
									DNLoopRowflag=True
									Exit For
									else
									r=false

								End If
							next

						End if
						If DNLoopRowflag=True Then
							RefreshDicObj=True
'							counter=0
							Exit for
						End If
					else
						if strcomp(elem,KeyItemColumn)<>0 then
							counter = counter+1
							colNumb=GetWTColumnNumber(TableObject,elem) 
							strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
							strCellValue1 = replace(strCellValue1,",","")
							If instr(1,TbleColValue,"litre")<>0 Then
								strCellValue1 =strCellValue1 & " litre"
							End If
							If instr(1,trim(strCellValue1),TbleColValue)<>0 Then
							AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1)&"</B>","PASSED"	
								'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"PASSED"
								writelog "","","","","Report(): Column Value matched. Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)
								bMatchedRow = True
							else
							AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1)&"</B>","FAILED"	
								'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"FAILED"
								writelog "","","","","Report(): Column Value not matched!! Expected Value : " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)

								bMatchedRow1=false
						 		
						 		'bMatchedRow = False
								'DNLoopRowflag=True
								'RefreshDicObj=True
'								TblrowNum=0
'								LastRow=LastRow+1
'								counter=0
'								Exit for
							End if

						End if
					End If

		 	Next
	 		
	 		If counter=totalDictNum Then
	 			RefreshDicObj=False
	 		End If
	 		
	 Loop While (RefreshDicObj=True )
  
	'If ErrNumber<>0 Then
	If bMatchedRow1=false Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
		ValidateDtWtColValue1=false
	else
		ValidateDtWtColValue1=true
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If

'  r=true
'  bMatchedRow1=true
'  ValidateDtWtColValue1 = bMatchedRow
	
End Function


 '******************************************** HEADER ******************************************
' Name : Format Date to dd/mm/yyyy
' Description : generic function to format date
' Creator : Shruthi Lokesh
' Date : 9 Jan 2018
' Last Modified On : 
' Last Modified By : 
' Input Parameter : Date
' Output Parameter : Date
'******************************************** HEADER ******************************************
Public Function FormatDate(dtDate, strFormat)
	
	Dim dDate, dDay, dMonth
	
	Select Case strFormat
						
	Case "DD/MM/YYYY"
						
		dDate = cDate(dtDate)
	
		dDay = day(dDate)
		if dDay < 10 then
			dDay = "0" & dDay
		End if
	
		dMonth = month(dDate)
	
		If dMonth < 10 Then
			dMonth = "0" & dMonth
		End If
	
		Dim lookupDate
		lookupdate = dDay & "/" & dMonth & "/" & year(dtDate)
	
	Case "MM/DD/YYYY"
	
		dDate = cDate(dtDate)
	
		dDay = day(dDate)
		if dDay < 10 then
			dDay = "0" & dDay
		End if
	
		dMonth = month(dDate)
	
		If dMonth < 10 Then
			dMonth = "0" & dMonth
		End If
	
		'Dim lookupDate
		lookupdate = dDay & "/" & dMonth & "/" & year(dtDate)
		
	Case "DD/MM/YY"
	
	''	dDate = cDate(dtDate)  'commented by Shabana
	
		dDay = day(dtDate)
'		if dDay < 10 then
'			dDay = "0" & dDay
'		End if
	
		dMonth = month(dtDate)
	
'		If dMonth < 10 Then
'			dMonth = "0" & dMonth
'		End If
	
		'Dim lookupDate
		lookupdate = dDay & "/" & dMonth & "/" & Right(Year(dtDate),2)
	
	End Select
	
	FormatDate = lookupdate
End Function


'******************************************** HEADER ******************************************
' Description : The function for validating the webtable column value based on Item Name
' Creator :  Shruthi L 
' Date : 9 Jan 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : Browser, Page, Frame, Object, Item Name
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateDtWtColValueBasedOnItemName(sBrowser,sPage,sFrame,sObject,ItemName, ColumnName, Column, RowAddition)
	
  On error resume next  
  set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)

	TableObject.Highlight
  
  	FirstItemRowNo=GetExactItemRowNumber(TableObject,ItemName,"",Column)
  	
  	FirstItemRowNo = FirstItemRowNo + RowAddition

	colNumb=GetWTColumnNumber(TableObject,ColumnName) 		
	strCellValue = TableObject.GetCellData(FirstItemRowNo,colNumb)
	 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then	 	
	 	writelog "","","","","ValidateDtWtColValueBasedOnItemName(): ValidateDtWtColValueBasedOnItemName is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","ValidateDtWtColValueBasedOnItemName(): ValidateDtWtColValueBasedOnItemName is Successful. Object Name : "&sWebTable
	 End If
	 
  ValidateDtWtColValueBasedOnItemName = strCellValue
	
End Function


'******************************************** HEADER ******************************************
' Description : The function for getting exact Item Row
' Creator :  Shruthi L 
' Date : 9 Jan 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function GetExactItemRowNumber(byref objTableObject,ItemName,startrow, column)
  
  On error resume next  
	
	If startrow="" Then
		startrow=1
	End If
	
	objTableObject.Highlight
	
	rCount=objTableObject.RowCount
	
	For row = cint(startrow) To rCount Step 1	  		
		strActual = objTableObject.GetCellData(row,column)
		strActual = trim(strActual)
		
		If strActual=ItemName Then
			ItemRow = row
			Exit for
		End If
		
	Next 
	
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then	 	
	 	writelog "","","","","GetExactPumpRowNumber(): GetExactPumpRowNumber is unsucessfull! Object Name : "&PumpNumber & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","GetExactPumpRowNumber(): GetExactPumpRowNumber is Successful. Object Name : "&PumpNumber
	 End If
  
  GetExactItemRowNumber = ItemRow
	
	
End Function

'******************************************** HEADER ******************************************
' Description : Validate column value for Item
' Creator :  Shruthi L 
' Date : 9 Jan 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function ValidateDtWtColValueForItemShruthi(sBrowser,sPage,sFrame,sObject,byref DictTbl, AddRowCount, aKey)
	
 On error resume next  

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  LastRow=2

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count

	 rowNum=TableObject.GetRowWithCellText(aKey)
		 	For Each elem in DictTbl
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						 
						 	colNumb=GetWTColumnNumber(TableObject,elem) 		
								strCellValue = TableObject.GetCellData((rowNum + AddRowCount),colNumb)
							
								If instr(1,trim(strCellValue), TbleColValue ) <> 0 Then	
								    
									AppendTestHTMLNEW StepCounter,"Verify report Value - "&akey,"<B>" &"Item Name: " & akey&"</B>","","PASSED"	
									AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue)&"</B>","PASSED"	
									'Append_TestHTML StepCounter,"Verify report Value - "&akey,"Item Name: " & akey,"PASSED"
									'Append_TestHTML StepCounter,"Verify report Value - "&elem,"Expected Value : " & TbleColValue & VBCRLF & " Actual Value: " & trim(strCellValue) ,"PASSED"
									writelog "","","","","Report(): Expected Value : " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									LastRow=row
									TblrowNum=rowNum
									DNLoopRowflag=True
								End If
						If DNLoopRowflag=True Then
							RefreshDicObj=True
						End If
					else
							colNumb=GetWTColumnNumber(TableObject,elem) 
							strCellValue1 = TableObject.GetCellData((TblrowNum + AddRowCount),colNumb)
							If instr(1,trim(strCellValue1),TbleColValue)<>0 Then
							    AppendTestHTMLNEW StepCounter,elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1)&"</B>","PASSED"	
								'Append_TestHTML StepCounter,"Verify report Value"&elem,"Expected Value :" & TbleColValue & " Actual Value: " & trim(strCellValue1),"PASSED"
								writelog "","","","","Report(): Column Value matched. Expected Value :" & TbleColValue & " Actual Value: " & trim(strCellValue1)
								bMatchedRow = True
							else
							     AppendTestHTMLNEW StepCounter,"Verify report Value","<B>" &TbleColValue&"</B>","<B>"& trim(strCellValue1)&"</B>","PASSED"	
								'Append_TestHTML StepCounter,"Verify report Value","Expected Value :" & TbleColValue & " Actual Value: " & trim(strCellValue1),"FAILED"
								writelog "","","","","Report(): Column Value not matched!! Expected Value :" & TbleColValue & " Actual Value: " & trim(strCellValue1)
							End if
					End If
		 	Next
	 		
  
	If ErrNumber<>0 Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
  
  ValidateDtWtColValueForItem = bMatchedRow
	
End Function

Function RecoveryItemSearch(Object)
 
Set bBrowser = CreateDescriptionObject(Comm_Browser3)
      Set bDialog = CreateDescriptionObject("Dialog_Popup")
  
     Set bButton = CreateDescriptionObject("WebBtn_OK")
	
Browser(bBrowser).Dialog(bDialog).WinButton(bButton).Click
 
End Function 

'******************************************** HEADER ******************************************
' Description : Validate column value for Item
' Creator :  Shruthi L 
' Date : 9 Mar 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

Public Function OperateOnWeb2DialogPage(sBrowser,sWindow,sWindow1,sPage,sFrame,sObject,sOperation,strData)

	On error resume next
		
	Set bBrowser = CreateDescriptionObject(sBrowser)
    Set bWindow = CreateDescriptionObject(sWindow)
    Set bWindow1 = CreateDescriptionObject(sWindow1)
    Set bPage = CreateDescriptionObject(sPage)
	Set bFrame = CreateDescriptionObject(sFrame)    
    Set bObject = CreateDescriptionObject(sObject)
    
    If sWindow1 = "NoWindow1" and sFrame = "NoFrame" Then
    	set ochild = Browser(bBrowser).Window(bWindow).Page(bPage).ChildObjects(bObject)
    ElseIf sWindow1 = "NoWindow1" Then
       	set ochild = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).ChildObjects(bObject)
    ElseIf sFrame = "NoFrame" Then
    	set ochild = Browser(bBrowser).Window(bWindow).Window(bWindow1).Page(bPage).ChildObjects(bObject)
    Else
    	set ochild = Browser(bBrowser).Window(bWindow).Window(bWindow1).Page(bPage).Frame(bFrame).ChildObjects(bObject)
	End If
	
	If StrChildIndex <> "" Then
		Set sochild= ochild(StrChildIndex)
	Else
		Set sochild= ochild(0)
	End If
'	Set sochild= ochild(0)


	Select Case sOperation
  
  	Case "Click"
  	
	  	If  sochild.Count <> 0 Then
			sochild.Click			
		End If
		
	Case "Highlight"
  	
	  	If  sochild.Count<>0 Then	  	
			sochild.Highlight		
		End If
	
	Case "DoubleClick"
  	
	  	If  sochild.Count <> 0 Then
			sochild.DoubleClick
		End If
			
	Case "Set"	
		
		If sochild.Count <> 0 Then
			
			sochild.Set strData
						
		End If 
		
	Case "SendKeys"	
		
		Set WshShell = CreateObject("WScript.Shell")
		If sochild.Count <> 0 Then
			
'			sochild.Set strData
			WshShell.SendKeys strData
			wait 2
'			WshShell.SendKeys "{ENTER}"
						
		End If 		
		
	Case "Select"
  		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count <> 0 Then
			sochild.Click
			wait 1
			sochild.Set strData
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1
			sochild.Click
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1
		End If 
		
	Case "Select1"
  		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count <> 0 Then
			sochild.Click
			wait 1
			WshShell.SendKeys strData
			wait 1
			WshShell.SendKeys "{ENTER}"
		End If 
		
	Case "GetRoProperty"
'  		Set WshShell = CreateObject("WScript.Shell")
		
		If sochild.Count <> 0 Then
			strRoProp = sochild.GetROProperty(strData)
			OperateOnWeb2DialogPage = strRoProp
			wait 1
'			WshShell.SendKeys strData
'			wait 1
'			WshShell.SendKeys "{ENTER}"
		End If 

 End Select 
 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"OperateOnWeb2DialogPage:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage	
	 	writelog "","","","","OperateOnWeb2DialogPage(): OperateOnWeb2DialogPage is unsucessfull! Object Name : "&sObject & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","OperateOnWeb2DialogPage(): OperateOnWeb2DialogPage is Successful. Object Name : "&sObject
	 End If
 
 StrChildIndex = Empty
 strRoProp = Empty
	
End Function

'******************************************** HEADER ******************************************
' Description : Function to Click on the cell in webtable
' Creator : Rocky
' Date : 26th Feb 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
sub webtable_clickcell(sBrowser,sWindow,sPage,sFrame,sWebTable,obj_typ,row,col)
	On error resume next
	
Set bBrowser = CreateDescriptionObject(sBrowser)

  Set bPage = CreateDescriptionObject(sPage)

  Set bWebTable = CreateDescriptionObject(sWebTable)

If sWindow = "No Window" and sFrame = "No Frame"  Then
	
	Set child_item = Browser(bBrowser).Page(bPage).WebTable(bWebTable)
	
	ElseIf sFrame = "No Frame" Then
	
	Set bWindow = CreateDescriptionObject(sWindow)
	Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).WebTable(bWebTable)
	
	ElseIf sWindow = "No Window" Then
	
	Set bFrame = CreateDescriptionObject(sFrame)
	Set child_item = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable)
	
	else

Set bWindow = CreateDescriptionObject(sWindow)
Set bFrame = CreateDescriptionObject(sFrame)
Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable)


End If


Set child_item1 = child_item.ChildItem(row,col,obj_typ,0)
child_item1.highlight
child_item1.click


Set bBrowser = nothing
Set bWindow =nothing
Set bPage = nothing
Set bFrame = nothing
Set bWebTable = nothing
set child_item=nothing
Set child_item1=nothing
	
If Err.Number<>0 Then
	 	Reporter.ReportEvent micPass,"webtable_clickcell:Error","Error Number : " & Err.Number & " -- Error Description : " & Err.description
	 	writelog "","","","","webtable_clickcell(): webtable_clickcell is unsucessfull! Object type : "&obj_typ & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","webtable_clickcell(): webtable_clickcell is Successful. Object Name : "&obj_typ
	 End If
 
	
End sub 
'******************************************** HEADER ******************************************
' Description : Function to enter values in edit box
' Creator : Rocky
' Date : 26th Feb 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************

sub WebtableEnterEdtbx(sBrowser,sWindow,sPage,sFrame,sWebTable,row,col,str_val)
	On error resume next
	
Set bBrowser = CreateDescriptionObject(sBrowser)

  Set bPage = CreateDescriptionObject(sPage)

  Set bWebTable = CreateDescriptionObject(sWebTable)

If sWindow = "No Window" and sFrame = "No Frame"  Then
	
	Set child_item = Browser(bBrowser).Page(bPage).WebTable(bWebTable)
	
	ElseIf sFrame = "No Frame" Then
	
	Set bWindow = CreateDescriptionObject(sWindow)
	Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).WebTable(bWebTable)
	
	ElseIf sWindow = "No Window" Then
	
	Set bFrame = CreateDescriptionObject(sFrame)
	Set child_item = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable)
	
	else

Set bWindow = CreateDescriptionObject(sWindow)
Set bFrame = CreateDescriptionObject(sFrame)
Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable)


End If


Set child_item1 = child_item.ChildItem(row,col,"WebEdit",0)
child_item1.highlight
child_item1.set str_val


Set bBrowser = nothing
Set bWindow =nothing
Set bPage = nothing
Set bFrame = nothing
Set bWebTable = nothing
set child_item=nothing
Set child_item1=nothing
	
If Err.Number<>0 Then
	 	Reporter.ReportEvent micPass,"WebtableEnterEdtbx:Error","Error Number : " & Err.Number & " -- Error Description : " & Err.description
	 	writelog "","","","","WebtableEnterEdtbx(): WebtableEnterEdtbx is unsucessfull! Object type : "&obj_typ & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","WebtableEnterEdtbx(): WebtableEnterEdtbx is Successful. Object Name : "&obj_typ
	 End If
 
	
End sub 
'******************************************** HEADER ******************************************
' Description : Function to enter values in edit box
' Creator : Rocky
' Date : 26th Feb 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************


function webtable_getrocell(sBrowser,sWindow,sPage,sFrame,sWebTable,obj_typ,row,col,str_propval)
	On error resume next
	
Set bBrowser = CreateDescriptionObject(sBrowser)

  Set bPage = CreateDescriptionObject(sPage)

  Set bWebTable = CreateDescriptionObject(sWebTable)

If sWindow = "No Window" and sFrame = "No Frame"  Then
	
	Set child_item = Browser(bBrowser).Page(bPage).WebTable(bWebTable)
	
	ElseIf sFrame = "No Frame" Then
	
	Set bWindow = CreateDescriptionObject(sWindow)
	Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).WebTable(bWebTable)
	
	ElseIf sWindow = "No Window" Then
	
	Set bFrame = CreateDescriptionObject(sFrame)
	Set child_item = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable)
	'Set child_item = GetFrameObject(sBrowser,sPage,sFrame,bWebTable)
	else

Set bWindow = CreateDescriptionObject(sWindow)
Set bFrame = CreateDescriptionObject(sFrame)
Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable)


End If


Set child_item1 = child_item.ChildItem(row,col,obj_typ,0)
child_item1.highlight
webtable_getrocell=child_item1.getroproperty(str_propval)


Set bBrowser = nothing
Set bWindow =nothing
Set bPage = nothing
Set bFrame = nothing
Set bWebTable = nothing
set child_item=nothing
Set child_item1=nothing
	
If Err.Number<>0 Then
	 	Reporter.ReportEvent micPass,"webtable_clickcell:Error","Error Number : " & Err.Number & " -- Error Description : " & Err.description
	 	writelog "","","","","webtable_clickcell(): webtable_clickcell is unsucessfull! Object type : "&obj_typ & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","webtable_clickcell(): webtable_clickcell is Successful. Object Name : "&obj_typ
	 End If
 
	
End function 


'******************************************** HEADER ******************************************
' Description : Validate column value for Item
' Creator :  Shruthi L 
' Date : 9 Jan 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateDtWtColValueForItem(sBrowser,sPage,sFrame,sObject,byref DictTbl, AddRowCount, aKey)
	
 On error resume next  

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  LastRow=2

	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject)
  
  TableObject.Highlight
  counter=0
  totalDictNum = DictTbl.Count
  
  	If aKey = "Total" Then
		For Iterator = 2 To TableObject.rowcount   Step 1
			If TableObject.getcelldata(Iterator,1)="Total" Then
				 rowNum = Iterator
			End If
		Next
	Else	
	rowNum=TableObject.GetRowWithCellText(aKey)
	End If	

	' rowNum=TableObject.GetRowWithCellText(aKey)
		 	For Each elem in DictTbl
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						 
						 	colNumb=GetWTColumnNumber(TableObject,elem) 		
								strCellValue = TableObject.GetCellData((rowNum+AddRowCount),colNumb)
								
								If instr(1,trim(strCellValue), TbleColValue ) <> 0 Then	
										sWetItemDetails =split(aKey,":")									
										sWetItemDetails_1 = sWetItemDetails(1)
									'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED")
									AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"
									'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
									writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									LastRow=row
									TblrowNum=rowNum
									DNLoopRowflag=True
									bMatchedRow = True
								End If
						If DNLoopRowflag=True Then
							RefreshDicObj=True
						End If
					else
							colNumb=GetWTColumnNumber(TableObject,elem) 
							strCellValue1 = TableObject.GetCellData((TblrowNum+AddRowCount),colNumb)
							If instr(1,trim(strCellValue1),TbleColValue)<>0 Then
								'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"PASSED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"PASSED"
								writelog "","","","","Report(): Column Value matched. Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)
								bMatchedRow = True
							else
								'Call HTMLTableMessage("Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"FAILED")
								AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","FAILED"
								
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"FAILED"
								writelog "","","","","Report(): Column Value not matched!! Expected Value : " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)							
							End if
					End If
		 	Next
	 		
  
	If ErrNumber<>0 Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
  
  ValidateDtWtColValueForItem = bMatchedRow
	
End Function

'******************************************** HEADER ******************************************
' Description : Validate column value for Item
' Creator :  Shruthi L 
' Date : 9 Mar 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function OperateOnWeb2DialogPage1(sBrowser,sWindow,sWindow1,sPage,sFrame,sObject,sOperation,strData)

	On error resume next
		
	Set bBrowser = CreateDescriptionObject(sBrowser)
    Set bWindow = CreateDescriptionObject(sWindow)
    Set bWindow1 = CreateDescriptionObject(sWindow1)
    Set bPage = CreateDescriptionObject(sPage)
	Set bFrame = CreateDescriptionObject(sFrame)    
    Set bObject = CreateDescriptionObject(sObject)
    
    If sOperation =  "ClickWithAtrr" Then
    	Dim strAddAtt
    	strAddAtt = "innerhtml:" & strData
    	Set bObject = CreateDescriptionObjectV2(sObject, strAddAtt)
    	sOperation = "Click"
    End If
  		
    If sWindow1 = "NoWindow1" and sFrame = "NoFrame" Then
    	set ochild = Browser(bBrowser).Window(bWindow).Page(bPage).ChildObjects(bObject)
    ElseIf sWindow1 = "NoWindow1" Then 	
    	set ochild = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).ChildObjects(bObject)
    ElseIf sFrame = "NoFrame" Then
    	set ochild = Browser(bBrowser).Window(bWindow).Window(bWindow1).Page(bPage).ChildObjects(bObject)
    Else
    	set ochild = Browser(bBrowser).Window(bWindow).Window(bWindow1).Page(bPage).Frame(bFrame).ChildObjects(bObject)
	End If
	
	Set sochild= ochild(0)


	Select Case sOperation
  
  	Case "Click"
			sochild.highlight
			sochild.Click			
		
	Case "Set"				
			sochild.Set strData
		
	Case "Select"
  		Set WshShell = CreateObject("WScript.Shell")
		
			sochild.Click
			wait 1
			sochild.Set strData
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1
			sochild.Click
			wait 1
			WshShell.SendKeys "{TAB}"
			wait 1

 End Select 
 
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"OperateOnWeb2DialogPage:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage	
	 	writelog "","","","","OperateOnWeb2DialogPage(): OperateOnWeb2DialogPage is unsucessfull! Object Name : "&sObject & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","OperateOnWeb2DialogPage(): OperateOnWeb2DialogPage is Successful. Object Name : "&sObject
	 End If
 
	
End Function


	
Public Function VerifyObjectExistInWebTbl(sBrowser,sPage,sFrame,sWebTable,sObject,PropertyName,PropertyValue)
 
  On error resume next  
  
  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bPage = CreateDescriptionObject(sPage)
  Set bWebTable = CreateDescriptionObject(sWebTable)
  Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
  If sFrame = "NoFrame" Then
  	Set sochild = Browser(bBrowser).Page(bPage).WebTable(bWebTable).ChildObjects(bObject)   
  Else
  	Set bFrame = CreateDescriptionObject(sFrame)
  	Set sochild = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable).ChildObjects(bObject)   
  End If
  

  If sochild.count<>0 Then 
  	 sochild(0).Highlight  	 
     VerifyObjectExistInWebTbl = True   
 
  End If
  
 ErrNumber = Err.Number
 ErrMessage = Err.Description
	 
 If ErrNumber<>0 Then	 	
 	writelog "","","","","VerifyObjectExistInWebTbl(): VerifyObjectExistInWebTbl is unsucessfull! Object Name : "&PropertyValue & " Error Number : "&err.number & " Error Description : "&Err.description
 else
 	writelog "","","","","VerifyObjectExistInWebTbl(): VerifyObjectExistInWebTbl is Successful. Object Name : "&PropertyValue
 End If
 
 End Function
 
 
 '******************************************** HEADER ******************************************
' Description : Function to Opertate on the objects in the webtable
' Creator : Ramesh Kumar KB
' Date : 28th Feb 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Function OperateOnWebtable(sBrowser,sWindow,sPage,sFrame,sWebTable,obj_typ,row,col,act_type,str_val)

	On error resume next
	
	sWindowTemp = split(sWindow,"||")
	cWincount = ubound(sWindowTemp)
		
	Set bBrowser = CreateDescriptionObject(sBrowser)
	Set bWindow = CreateDescriptionObject(sWindow)
	Set bPage = CreateDescriptionObject(sPage)
	Set bFrame = CreateDescriptionObject(sFrame)
	Set bWebTable = CreateDescriptionObject(sWebTable)
	
''	GetFrameWebWinObject(sBrowser,sWindow,sPage,sFrame,sObject)
	
	If sWindow = "No_Window" and sFrame = "No_Frame"  Then
		Set child_item = Browser(bBrowser).Page(bPage).WebTable(bWebTable)
	ElseIf sBrowser = "No_Browser" Then
		Set child_item = Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable)
'''''		Set child_item = GetFrameWebWinObject(sBrowser,sWindow,sPage,sFrame,sWebTable)
	ElseIf sFrame = "No_Frame" Then
		Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).WebTable(bWebTable)
	ElseIf sWindow = "No_Window" Then
		Set child_item = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable)
''''''		Set child_item = GetFrameWebWinObject(sBrowser,sWindow,sPage,sFrame,sWebTable)
	ElseIf cWincount = 1 Then
		Set bWindow0 = CreateDescriptionObject(sWindowTemp(0))
		Set bWindow1 = CreateDescriptionObject(sWindowTemp(1))
		Set child_item = Browser(bBrowser).Window(bWindow0).Window(bWindow1).Page(bPage).Frame(bFrame).WebTable(bWebTable)
''''''		Set child_item = GetFrameWebWinObject(sBrowser,sWindow,sPage,sFrame,sWebTable)
	else
		Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable)
'''''		Set child_item = GetFrameWebWinObject(sBrowser,sWindow,sPage,sFrame,sWebTable)
	End If
	
	Set child_item1 = child_item.ChildItem(row,col,obj_typ,0)
	
	Select Case act_type
	
		Case "RowCount"
			OperateOnWebtable = child_item1.rowcount()
		
		Case "Click"
			child_item1.highlight
			child_item1.click
			
		Case "Set"
			child_item1.Set str_val
			
		Case "AutoSelect"
			Set WshShell = CreateObject("WScript.Shell")
			child_item1.click
			WshShell.SendKeys str_val
			WshShell.SendKeys "{ENTER}"
			
'''			'If You dont want row or col values, then pass 0
'''''     Use Webtable object not child of webtable
		Case "GetRowWithCellText"
			 StrGrwct = child_item.GetRowWithCellText(str_val,col,row)
			 OperateOnWebtable = StrGrwct
			 
		Case "GetCellData"
			 StrGcd = child_item.GetCellData(row,col)
			 OperateOnWebtable = StrGcd
			 
	End Select
	
		
	If Err.Number <> 0 Then
		 	Reporter.ReportEvent micPass,"OperateOnWebtable:Error","Error Number : " & Err.Number & " -- Error Description : " & Err.description
		 	writelog "","","","","OperateOnWebtable(): OperateOnWebtable is unsucessfull! Object type : "&obj_typ & " Error Number : "&err.number & " Error Description : "&Err.description
		 else
		 	writelog "","","","","OperateOnWebtable(): OperateOnWebtable is Successful. Object Name : "&obj_typ
	End If
	 
	Set bBrowser = nothing
	Set bWindow =nothing
	Set bPage = nothing
	Set bFrame = nothing
	Set bWebTable = nothing
	set child_item=nothing
	Set child_item1=nothing
	
	StrGrwct = Empty
	StrGcd = Empty
	cWincount =Empty
	
End Function


Public Function VerifyWebdialogObjectExist(sBrowser,sDialog,sObject)
 
  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bDialog = CreateDescriptionObject(sDialog)
  Set bObject = CreateDescriptionObject(sObject)
  
 	VerifyWebdialogObjectExist = false


 	set sochild = Browser(bBrowser).Dialog(bDialog).WinButton(bObject) 	 	

	If Isobject(sochild) Then
		
	 	If  sochild.Exist(5) then 
	       VerifyWebdialogObjectExist = True  
	 	End If
	Else
	 		VerifyWebdialogObjectExist = false
	End If

  	Set sochild = nothing
 
 Set bBrowser = nothing
 Set bDialog = nothing
  
 Set bObject=nothing
 
 End Function
 
 
'******************************************** HEADER ******************************************
' Description : The function for validating the data of the webtable row value
' Creator : Niharika  
' Date : 19th Dec, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateDtWtRowValueModified(sBrowser,sPage,sFrame,sObject1,byref DictTbl,colMappingVal,keyColNum)
	
	On error resume next  

	bflag=True
	DNLoopRowflag=False
	RefreshDicObj=False
	strAlreadyChecked = ""
	LastRow=2
	
	Set colMapVal = CreateObject("Scripting.Dictionary")
	tempArr1 = split(colMappingVal,",")
	For a=0 to ubound(tempArr1)
		tempArr = split(tempArr1(a),":")
		Call Add_Dictionary( colMapVal,tempArr(0),tempArr(1))
	Next
	
	set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sObject1)
	  
	TableObject.Highlight
	counter=0
	totalDictNum = DictTbl.Count
	colCount = TableObject.ColumnCount(1)

	For Iterator = 1 To colCount-1 Step 1
		If Iterator <> keyColNum Then
			reqVal = TableObject.GetCellData(1,Iterator)
			If strcomp(reqVal,"Amount") = 0 or instr(reqVal,"Amount") <> 0 or instr(reqVal,"R") = 1 Then
				colType = "Amt"
			ElseIf instr(reqVal,"Qty") <> 0 Then
				colType = "QTY"
			ElseIf strcomp(reqVal,"Actual") = 0 Then
				colType = "Actual"
			ElseIf strcomp(reqVal,"Expected") = 0 Then
				colType = "Exp"
			ElseIf strcomp(reqVal,"Variance") = 0 Then
				colType = "Var"
			ElseIf strcomp(reqVal,"Value") = 0 Then
				colType = "Val"
			ElseIf strcomp(reqVal,"% Variance") = 0 Then
				colType = "VarPer"
			ElseIf instr(reqVal,"%") <> 0 Then
				colType = "Percent"
			ElseIf reqVal = "" Then
				n = cstr(Iterator)
				colType = colMapVal.item(n)
			End If
	
		 	For Each elem in DictTbl
		 		counter = counter+1
				TblRowName=elem
				TbleRowValue=DictTbl(elem)
			 	colNumb = Iterator
			 	test1 = split(TblRowName,"_")
			 	If test1(1) = colType Then
			 		
				 	If instr(TblRowName,colType) <> 0 Then
				 		TblRowName = replace(TblRowName,"_","")
				 		TblRowName = replace(TblRowName,colType,"")
				 	End If
				 	
				 	tempVar1 = replace(TblRowName,"(L)","")
					tempArr1 = split(tempVar1," ")
					tempVar2 = tempArr1(0)
					If isNumeric(tempVar2) = True Then
						s = tempArr1(1)
						For r=2 to ubound(tempArr1)
							h = tempArr1(r)
							s = s & " " & h						
						Next
						TblRowName = s
					ElseIf isNumeric(tempVar2) = False Then
						TblRowName = tempVar1
					End If
				 	
				 	rowNum=TableObject.GetRowWithCellText(TblRowName)
				 	rCount=TableObject.RowCount	
					strCellValue = TableObject.GetCellData(rowNum,colNumb)
					
					TblRowNameType = TblRowName & "_" & colType
					
					If rowNum <> -1 and strCellValue <> empty Then
						If instr(strCellValue,"R")<>0 or isNumeric(strCellValue) = True Then
							valType = "numType"
						ElseIf instr(strCellValue,"%")<>0 Then
							tempVal = replace(strCellValue,"%","")
							If isNumeric(tempVal) = True Then
								valType = "numType"
							else
								valType = "strType"
							End If
						else
							valType = "strType"
						End If
						
						If valType = "numType" Then
							strCellValue = ChangeNumberFormat(strCellValue)
							TbleRowValue = ChangeNumberFormat(TbleRowValue)
							If cdbl(strCellValue)=cdbl(TbleRowValue) Then
								bMatchedRow = True
							else
								bMatchedRow = false
							End If
						ElseIf valType = "strType" Then
							If strCellValue=TbleRowValue Then
								bMatchedRow = True
							else
								bMatchedRow = false
							End If
						End If
						
						If bMatchedRow = True then
						    AppendTestHTMLNEW StepCounter,"Report Validation Value - "&TblRowNameType,"<B>" &TbleRowValue&"</B>","<B>" & trim(strCellValue)&"</B>","PASSED"	
							'Append_TestHTML StepCounter,"Report Validation Value - "&TblRowNameType,"Expected Value: " & TbleRowValue & VBCRLF & " Actual Value: " & trim(strCellValue),"PASSED"
							writelog "","","","","Report Validation(): Column Value matched. Expected Value: " & TbleRowValue & VBCRLF &" Actual Value: " & trim(strCellValue)
						else
						    AppendTestHTMLNEW StepCounter,"Report Validation Value - "&TblRowNameType,"<B>" &TbleRowValue&"</B>","<B>" & trim(strCellValue)&"</B>","FAILED"	
							'Append_TestHTML StepCounter,"Report Validation Value - "&TblRowNameType,"Expected Value: " & TbleRowValue & VBCRLF & " Actual Value: " & trim(strCellValue),"FAILED"
							writelog "","","","","Report Validation(): Column Value not matched!! Expected Value: " & TbleRowValue & VBCRLF & " Actual Value: " & trim(strCellValue)
							counter=0
						End if
					End If				
					strCellValue = empty
				End if 
			Next
		
		ElseIf Iterator = keyColNum Then
			For Each elem in DictTbl
				keyColVar = elem
				keyColVarArr1 = split(keyColVar,"_")
				keyColVar1 = keyColVarArr1(0)
				keyColVar1Temp = replace(keyColVar1,"(L)","")
				keyColVarArr2 = split(keyColVar1Temp," ")
				keyColVar2 = keyColVarArr2(0)
				If isNumeric(keyColVar2) = True Then
					s = keyColVarArr2(1)
					For r=2 to ubound(keyColVarArr2)
						h = keyColVarArr2(r)
						s = s & " " & h						
					Next
					reqKeyVal = s
				ElseIf isNumeric(keyColVar2) = False Then
					reqKeyVal = keyColVar1
				End If
				
				If instr(strAlreadyChecked,reqKeyVal) = 0 Then
					strAlreadyChecked = strAlreadyChecked & " " & reqKeyVal
					reqRowNum = TableObject.GetRowWithCellText(reqKeyVal)
					reqColNum = Iterator
					reqCellValue = TableObject.GetCellData(reqRowNum,reqColNum)
					
					If reqRowNum <> -1 Then
						If instr(reqCellValue,"R")<>0 or isNumeric(reqCellValue) = True Then
							valType = "numType"
						ElseIf instr(reqCellValue,"%")<>0 Then
							tempVal = replace(reqCellValue,"%","")
							If isNumeric(tempVal) = True Then
								valType = "numType"
							else
								valType = "strType"
							End If
						else
							valType = "strType"
						End If

						If valType = "numType" Then
							reqCellValue = ChangeNumberFormat(reqCellValue)
							keyColVar1 = ChangeNumberFormat(keyColVar1)
							If cdbl(reqCellValue)=cdbl(keyColVar1) Then
								bMatchedRow = True
							else
								bMatchedRow = false
							End If
						ElseIf valType = "strType" Then
							If reqCellValue=keyColVar1 Then
								bMatchedRow = True
							else
								bMatchedRow = false
							End If
						End If
						If bMatchedRow = True then
							'Append_TestHTML StepCounter,"Report Validation Value - "&reqKeyVal,"Expected Value: " & keyColVar1 & VBCRLF & " Actual Value: " & trim(reqCellValue),"PASSED"
							AppendTestHTMLNEW StepCounter,"Report Validation Value - "&reqKeyVal,"<B>" &keyColVar1&"</B>","<B>" & trim(reqCellValue)&"</B>","PASSED"	
							writelog "","","","","Report Validation(): Column Value matched. Expected Value: " & keyColVar1 & VBCRLF &" Actual Value: " & trim(reqCellValue)
						else
							'Append_TestHTML StepCounter,"Report Validation Value - "&reqKeyVal,"Expected Value: " & keyColVar1 & VBCRLF & " Actual Value: " & trim(reqCellValue),"FAILED"
							AppendTestHTMLNEW StepCounter,"Report Validation Value - "&reqKeyVal,"<B>" &keyColVar1&"</B>","<B>" & trim(reqCellValue)&"</B>","FAILED"	
							writelog "","","","","Report Validation(): Column Value not matched!! Expected Value: " & keyColVar1 & VBCRLF & " Actual Value: " & trim(reqCellValue)
							LastRow=LastRow+1
							counter=0
						End if
					End if

				End If				
			Next
		End If		
	Next
  
	If ErrNumber<>0 Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
  
  ValidateDtWtRowValueModified = bMatchedRow
	
End Function


'******************************************** HEADER ******************************************
' Description : The function is to generate an HTML report steps where the messages are parameterised
' Creator :  Shabana Ashrafi
' Date : 20 Jan,2018
' Last Modified On : 
' Last Modified By : 
' Input Parameter : Column1Msg,Column2Msg,PassOrFailMsg
' Output Parameter :
'******************************************** HEADER ******************************************
Public Function HTMLTableMessage(Column1Msg,Column2Msg,PassOrFailMsg)
	
 If PassOrFailMsg ="PASSED" Then

	Reporter.ReportEvent micPass,Column1Msg,Column2Msg
	
	If instr(Column2Msg,"Validation")<>0 or instr(Column2Msg,"PDF")<>0 or instr(Column1Msg,"Report Time")<>0 Then
		'Append_TestHTML StepCounter,Column1Msg,Column2Msg,"PASSED"
		AppendTestHTMLNEW StepCounter,"","<B>" &Column1Msg&"</B>","<B>" & Column2Msg&"</B>","PASSED"	
	Else
		'Append_TestHTMLWithNoMultipleScreenshot StepCounter,Column1Msg,Column2Msg,"PASSED"
		AppendTestHTMLNEWwithoutSS StepCounter,"","<B>" &Column1Msg&"</B>","<B>" & Column2Msg &"</B>","PASSED"
	End If
	
	writelog "","","","",Column1Msg&" is Successful"
	
 Else
 
	Reporter.ReportEvent micFail,Column1Msg,Column2Msg
	
	If instr(Column2Msg,"Validation")<>0 or instr(Column2Msg,"PDF")<>0 or instr(Column1Msg,"Report Time")<>0 Then
		'Append_TestHTML StepCounter,Column1Msg,Column2Msg,"FAILED"
		AppendTestHTMLNEW StepCounter,"","<B>" &Column1Msg&"</B>","<B>" & Column2Msg&"</B>","FAILED"
	Else
	    AppendTestHTMLNEWwithoutSS StepCounter,"","<B>" &Column1Msg&"</B>","<B>" & Column2Msg &"</B>","FAILED"
		'Append_TestHTMLWithNoMultipleScreenshot StepCounter,Column1Msg,Column2Msg,"FAILED"
	End If
	
	writelog "","","","",Column1Msg&" is UnSuccessful"
	bRunFlag=False
	scriptFlag=False
 End If
 
 bCheckFlag = False
	
 End Function
 

Public Function HTMLTableMessageOLD(Column1Msg,Column2Msg,PassOrFailMsg)
	
 If PassOrFailMsg ="PASSED" Then

	Reporter.ReportEvent micPass,Column1Msg,Column2Msg
	AppendTestHTMLNEW StepCounter,"","<B>" &Column1Msg&"</B>","<B>" & Column2Msg&"</B>","PASSED"	
	'Append_TestHTML StepCounter,Column1Msg,Column2Msg,"PASSED"
	writelog "","","","",Column1Msg&" is Successful"
	
 Else
 
	Reporter.ReportEvent micFail,Column1Msg,Column2Msg
		AppendTestHTMLNEW StepCounter,"","<B>" &Column1Msg&"</B>","<B>" & Column2Msg&"</B>","FAILED"	
	'Append_TestHTML StepCounter,Column1Msg,Column2Msg,"FAILED"
	writelog "","","","",Column1Msg&" is UnSuccessful"
	bRunFlag=False
	scriptFlag=False
 End If
 
 bCheckFlag = False
	
 End Function
 
 
Public Function GetNewConsolidatedValue(strLastValue,sCurrentValue)
	
	If strLastValue="" Then
		strNewValue = cdbl(sCurrentValue)
	Else
		strNewValue = cdbl(strLastValue)+cdbl(sCurrentValue)
	End If
	
	GetNewConsolidatedValue = strNewValue

End Function


Public Function UpdateNewValueInDictionary(sDictName,sTagName,sOperation,sCurrentValue)
	
	
	strLastValue = Get_Dictionary(sDictName,sTagName)
	
	Select Case sOperation
		
		
		Case "Add"
		
			If strLastValue="" Then
				strNewValue = cdbl(formatnumber(sCurrentValue))
			Else
				strNewValue = cdbl(formatnumber(strLastValue))+cdbl(formatnumber(sCurrentValue))
			End If
		
		Case "Subtract"
		
			If strLastValue="" Then
				strNewValue = cdbl(sCurrentValue)
			Else
				strNewValue = cdbl(strLastValue)-cdbl(sCurrentValue)
			End If		
		
		Case "Multiply"
		
			If strLastValue="" Then
				strNewValue = cdbl(sCurrentValue)
			Else
				strNewValue = cdbl(strLastValue)*cdbl(sCurrentValue)
			End If	
		
		Case "Update"
		
			strNewValue = sCurrentValue
			
		Case "Append"

            If strLastValue="" Then
				strNewValue = sCurrentValue
			ElseIf instr(strLastValue,sCurrentValue)=0 Then
				strNewValue =strLastValue&","&sCurrentValue
			Else
                strNewValue = strLastValue			
			End If
		
		Case "AllValue"

            If strLastValue="" Then
				strNewValue = sCurrentValue
		
			Else
               strNewValue =strLastValue&","&sCurrentValue		
			End If
		
		
	End Select
	
	Call Update_Dictionary(sDictName,sTagName,strNewValue)
	
	
End Function


function webtable_rowcount(sBrowser,sWindow,sPage,sFrame,sWebTable)
	
	On error resume next
	
	Set bBrowser = CreateDescriptionObject(sBrowser)

 	Set bPage = CreateDescriptionObject(sPage)

  	Set bWebTable = CreateDescriptionObject(sWebTable)

	If sWindow = "No Window" and sFrame = "No Frame"  Then
	
		Set child_item = Browser(bBrowser).Page(bPage).WebTable(bWebTable)
	
	ElseIf sFrame = "No Frame" Then
	
		Set bWindow = CreateDescriptionObject(sWindow)
		Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).WebTable(bWebTable)
	
	ElseIf sWindow = "No Window" Then
	
		Set bFrame = CreateDescriptionObject(sFrame)
		Set child_item = Browser(bBrowser).Page(bPage).Frame(bFrame).WebTable(bWebTable)
	
	else

		Set bWindow = CreateDescriptionObject(sWindow)
		Set bFrame = CreateDescriptionObject(sFrame)
		Set child_item = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebTable(bWebTable)
End If


child_item.highlight
webtable_rowcount = child_item.RowCount


Set bBrowser = nothing
Set bWindow =nothing
Set bPage = nothing
Set bFrame = nothing
Set bWebTable = nothing
set child_item=nothing

	
	If Err.Number<>0 Then
		 Reporter.ReportEvent micPass,"webtable_rowcount:Error","Error Number : " & Err.Number & " -- Error Description : " & Err.description
		 writelog "","","","","webtable_rowcount: webtable_rowcount is unsucessfull! Object type : "&obj_typ & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		 writelog "","","","","webtable_rowcount(): webtable_rowcount is Successful. Object Name : "&obj_typ
	End If
 
	
End function 

'******************************************** HEADER ******************************************
' Description : Function to set data to webedit under dialog
' Creator : Ramesh Kumar KB
' Date : 20 th February,2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function WaitForWeb2DialogObject(sBrowser,sWindow,sWindow1,sPage,sFrame,sObject,strData)

	On error resume next

	TimeCounter = 0
	strObjExt = False
	bExistFlag = False
	
	Do
	
		Set bBrowser = CreateDescriptionObject(sBrowser)
	    Set bWindow = CreateDescriptionObject(sWindow)
	    Set bWindow1 = CreateDescriptionObject(sWindow1)
	    Set bPage = CreateDescriptionObject(sPage)
		Set bFrame = CreateDescriptionObject(sFrame)    
	    Set bObject = CreateDescriptionObject(sObject)
    
	    If sWindow1 = "NoWindow1" and sFrame = "NoFrame" Then
	    	set ochild = Browser(bBrowser).Window(bWindow).Page(bPage).ChildObjects(bObject)
	    	Set sochild= ochild(0)
	    ElseIf sWindow1 = "NoWindow1" and sBrowser = "NoBrowser" Then
	    	set ochild = Window(bWindow).Page(bPage).Frame(bFrame).ChildObjects(bObject)
	    	Set sochild= ochild(0)
	'''		set ochild = GetFrameWeb2DiaObject(sBrowser,sWindow,sWindow1,sPage,sFrame,sObject)
	    ElseIf sWindow1 = "NoWindow1" Then
	       	set ochild = Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).ChildObjects(bObject)
	       	Set sochild= ochild(0)
	'''       	set ochild = GetFrameWeb2DiaObject(sBrowser,sWindow,sWindow1,sPage,sFrame,sObject)
	    ElseIf sFrame = "NoFrame" Then
	    	set ochild = Browser(bBrowser).Window(bWindow).Window(bWindow1).Page(bPage).ChildObjects(bObject)
	    	Set sochild= ochild(0)


	    Else
	    	set ochild = Browser(bBrowser).Window(bWindow).Window(bWindow1).Page(bPage).Frame(bFrame).ChildObjects(bObject)
	    	Set sochild= ochild(0)
	'''		set ochild = GetFrameWeb2DiaObject(sBrowser,sWindow,sWindow1,sPage,sFrame,sObject)
		End If

'		Do	

			If isobject(sochild) = True Then
				
				strObjExt = sochild.Exist(5)
				intobjhieght = cint(sochild.GetROProperty("height"))
				
				If strObjExt = "True" and intobjhieght > 0 Then
					bExistFlag = True
''''					Wait 1
					Exit do
				End If


			End If
'''			Else
				strData = cint(strData)
				If TimeCounter > strData Then
					Exit do
				End If
				TimeCounter = cint(TimeCounter + 1)
'''''				Wait 1

'''			End If


	
		Loop While bExistFlag = False
	
	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"WaitForWeb2DialogObject:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage	
	 	writelog "","","","","WaitForWeb2DialogObject(): WaitForWeb2DialogObject is unsucessfull! Object Name : "&sObject & " Error Number : "&err.number & " Error Description : "&Err.description
	 else
	 	writelog "","","","","WaitForWeb2DialogObject(): WaitForWeb2DialogObject is Successful. Object Name : "&sObject
	 End If
 
 	strObjExt = ""
	intobjhieght = ""

End Function

'******************************************** HEADER ******************************************
' Description : The function for validating the webtable cell value
' Creator :  Hemaraja 
' Date : 14th July, 2017
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateWebTableCellValue(sBrowser,sPage,sFrame,sWebTable,strCellSearchValue)

  On error resume next
  
  Dim rCount
  Dim cCount
  Dim row
  Dim col

  Set bBrowser = CreateDescriptionObject(sBrowser)
  Set bPage = CreateDescriptionObject(sPage)
  Set bFrame = CreateDescriptionObject(sFrame)
  Set bWebTable = CreateDescriptionObject(sWebTable)
  
  
  If sFrame = "No_Frame" Then
  	Set objchild = Browser(bBrowser).Page(bPage).ChildObjects(bWebTable)
  Else
  	Set objchild = GetFrameObject(sBrowser,sPage,sFrame,sWebTable) 
  End If

  ValidateWebTableCellValue = False
  
  If isobject(objchild) = True Then
  	
  
		  If cint(objchild.count) > 0 Then
		  	 objchild(0).Highlight
'			objchild.Highlight
		  	 
		  	 rCount=objchild(0).RowCount
'		  	 rCount=objchild.RowCount
		  	   	 
		  	 For row = 2 To rCount Step 1
		  	 	
		  	 	cCount=objchild(0).Columncount(row)  	 	
'		  	 	cCount=objchild.Columncount(row)  	 	
		  	 	 	 	
		  	 	For col = 1 To cCount Step 1 	  	 		
		  	 			
		  	 		If objchild(0).GetCellData(row,col)=strCellSearchValue Then	  
'					If objchild.GetCellData(row,col)=strCellSearchValue Then	  	 		  	 		
		  	 		  	 	 		  	
		  	 			ValidateWebTableCellValue = True 
		  	 			
		  	 			Exit for 
		  	 			
		  	 		End If 	 		
		  	 		
		  	 	Next
  
  	 Next
  	 
  End If
  End If
   	 ErrNumber = Err.Number
	 ErrMessage = Err.Description
	 
	 If ErrNumber<>0 Then
	 	Reporter.ReportEvent micPass,"OperateOnObject:Error","Error Number : " & ErrNumber & " -- Error Description : " & ErrMessage
	 End If


End Function


Public Function ChangeNumberFormat(reportNumber)
	
	reportNumber = replace(reportNumber,"£","")
	reportNumber = replace(reportNumber,"%","")
	reportNumber = replace(reportNumber,"R","")
	reportNumber = replace(reportNumber,",","")
	
	If reportNumber = "" or reportNumber = empty Then
		reportNumber = 0
	ElseIf instr(reportNumber,"(") > 0 and instr(reportNumber,")") > 0 Then
		reportNumber = replace(reportNumber,"(","")
		reportNumber = replace(reportNumber,")","")
		reportNumber = replace(reportNumber,",","")
		reportNumber = 0 - cdbl(reportNumber)
		ChangeNumberFormat = reportNumber
	End If
	reportNumber = cdbl(reportNumber)
	ChangeNumberFormat = reportNumber
	
End Function


Public Function ClickCashierReconWebBtn(sBrowser,sWindow,sPage,sFrame,sButton)
    
    Set bBrowser = CreateDescriptionObject(sBrowser)
    Set bWindow = CreateDescriptionObject(sWindow)
    Set bPage = CreateDescriptionObject(sPage)
    Set bFrame = CreateDescriptionObject(sFrame)
    Set bButton = CreateDescriptionObject(sButton)
      
    Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebButton(bButton).Highlight
    Browser(bBrowser).Window(bWindow).Page(bPage).Frame(bFrame).WebButton(bButton).Click
      
    Set bBrowser = nothing
    Set bWindow = nothing
    Set bPage = nothing
    Set bFrame = nothing
    Set bButton = nothing      
       
End Function


Function send_keys(str_string)
	Set str = createobject("Wscript.Shell")
	wait 1
	str.SendKeys(str_string)
	wait 1
	Set str=nothing
End Function

'******************************************** HEADER ******************************************
' Description : Validate function where columns are present in diff table and rows are present in diff table
' Creator :   
' Date : 9 Jan 2018
' Last Modified On :
' Last Modified By : 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function ValidateDtWtColValueForItemDiffTables(sBrowser,sPage,sFrame,sRowObject,sColumnObject,byref DictTbl, AddRowCount, aKey)
	
 On error resume next  

  bflag=True
  DNLoopRowflag=False
  RefreshDicObj=False
  LastRow=2

	'Column Table
	Set ColTableObject=GetFrameObject(sBrowser,sPage,sFrame,sColumnObject)
	ColTableObject.Highlight

	'Row Table
	 Set TableObject=GetFrameObject(sBrowser,sPage,sFrame,sRowObject)
	 TableObject.Highlight
	 counter=0
  
  	totalDictNum = DictTbl.Count
  
  	If aKey = "Total" Then
		For Iterator = 2 To TableObject.rowcount   Step 1
			If TableObject.getcelldata(Iterator,1)="Total" Then
				 rowNum = Iterator
			End If
		Next
	Else	
	rowNum=TableObject.GetRowWithCellText(aKey)
	End If	

		 	For Each elem in DictTbl
		 			
		 			TbleColValue=DictTbl(elem)
		 			
				 	If DNLoopRowflag=false Then
						 
						 	colNumb=GetWTColumnNumber(ColTableObject,elem) 		
								strCellValue = TableObject.GetCellData(rowNum,colNumb)
								
								If instr(1,trim(strCellValue), TbleColValue ) <> 0 Then	
										sWetItemDetails =split(aKey,":")									
										sWetItemDetails_1 = sWetItemDetails(1)
									AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue) &"</B>","PASSED"
									'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue) ,"PASSED"
									writelog "","","","","Report(): Expected Value: " & TbleColValue & VBCRLF &" Actual Value: " & trim(strCellValue) 
									LastRow=row
									TblrowNum=rowNum
									DNLoopRowflag=True
									bMatchedRow = True
								End If
						If DNLoopRowflag=True Then
							RefreshDicObj=True
						End If
					else
							colNumb=GetWTColumnNumber(ColTableObject,elem) 
							strCellValue1 = TableObject.GetCellData(TblrowNum,colNumb)
							If instr(1,trim(strCellValue1),TbleColValue)<>0 Then
							    AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","PASSED"
									'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"PASSED"
								writelog "","","","","Report(): Column Value matched. Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)
								bMatchedRow = True
							else
							AppendTestHTMLNEWwithoutSS StepCounter,"Verify report Value - "&elem,"<B>" &TbleColValue&"</B>","<B>" & trim(strCellValue1) &"</B>","FAILED"
								'Append_TestHTMLWithOutSS StepCounter,"Verify report Value - "&elem,"Expected Value: " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1),"FAILED"
								writelog "","","","","Report(): Column Value not matched!! Expected Value : " & TbleColValue & VBCRLF & "Actual Value: " & trim(strCellValue1)							
							End if
					End If
		 	Next
	 		
  
	If ErrNumber<>0 Then	 	
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is unsucessfull! Object Name : "&sWebTable & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","ValidateDtWtColValue(): ValidateDtWtColValue is Successful. Object Name : "&sWebTable
	End If
  
  ValidateDtWtColValueForItemDiffTables = bMatchedRow
	
End Function

'******************************************** HEADER ******************************************
' Name : AppendTestHTMLNEW
' Description : 
' Creator : Madhusmitta
' Date : 20th,March,2020
' Last Modified On : 
' Last Modified By : 
' Input Parameter : iStep,sType,sFieldName,sReceiptData,sAppData,sStatus
' Output Parameter : 
'******************************************** HEADER ******************************************
Public Function AppendTestHTMLNEW(iStep,sFieldName,sReceiptData,sAppData,sStatus)

	Dim sScreenshot
	
	If sStatus = "PASSED" Then
	    sScreenshot = TakeScreenshot
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sFieldName & "</a></td><td align=""center""><font color=""blue"">"&sReceiptData&"</td></font><td align=""center""><font color=""blue"">"&sAppData&"</td></font><td BGCOLOR=""#3ADF00"">PASSED</td><td align=""center""><a href=""" & sScreenshot & """>Screenshot</a></td></tr>"
	else
		sScreenshot = TakeScreenshot
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sFieldName & "</a></td><td align=""center""><font color=""blue"">"&sReceiptData&"</td></font><td align=""center""><font color=""blue"">"&sAppData&"</td></font><td BGCOLOR=""#FF0000"">FAILED</td><td align=""center""><a href=""" & sScreenshot & """>Screenshot</a></td></tr>"
	End If	
	
	StepCounter = StepCounter + 1
	
End Function

'******************************************** HEADER ******************************************
' Name : AppendTestHTMLNEWwithoutSS
' Description : 
' Creator : Madhusmitta
' Date : 20th,March,2020
' Last Modified On : 
' Last Modified By : 
' Input Parameter : iStep,sType,sFieldName,sReceiptData,sAppData,sStatus
' Output Parameter : 
'******************************************** HEADER ******************************************	
Public Function AppendTestHTMLNEWwithoutSS(iStep,sFieldName,sReceiptData,sAppData,sStatus)

	Dim sScreenshot
	
	If sStatus = "PASSED" Then
'	    sScreenshot = TakeScreenshot
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sFieldName & "</a></td><td align=""center""><font color=""blue"">"&sReceiptData&"</td></font><td align=""center""><font color=""blue"">"&sAppData&"</td></font><td BGCOLOR=""#3ADF00"">PASSED</td><td></td></tr>"
	else
'		sScreenshot = TakeScreenshot
		objFile_Testhtml.writeline  "<tr><td align=""center"">" & iStep & "</td><td>" & sFieldName & "</a></td><td align=""center""><font color=""blue"">"&sReceiptData&"</td></font><td align=""center""><font color=""blue"">"&sAppData&"</td></font><td BGCOLOR=""#FF0000"">FAILED</td><td></td></tr>"
	End If	
	
	StepCounter = StepCounter + 1
	
End Function

'******************************************** HEADER ******************************************
' Name : InitializeTestHTMLNEW
' Description : 
' Creator :
' Date : 2nd,August,2019
' Last Modified On : 
' Last Modified By : Madhusmitta
' Input Parameter : strTC
' Output Parameter : 
'******************************************** HEADER ******************************************	
'Public Function InitializeTestHTMLNEW(strTC)
'	
'	strGetOnlyBOS = GetBOSInstruction("ONLYBOS")
'	
'	
'	If strGetOnlyBOS="YES" Then
'		StepCounter = 1
'		strTestFile = TestResults & sFolderName & "\" & strTC & ".html"
'		
'		set objFile_Testhtml= fsObject.OpenTextFile(strTestFile,2,True)	
'		objFile_Testhtml.writeline "<head><style>body {    background-color: #E0FFFF;}</style></head>"
'		objFile_Testhtml.writeline "<h4 align='center'><u>"&strTC&"</u></h4>"
'		objFile_Testhtml.writeline "<br><br>"
'		objFile_Testhtml.writeline  "<table border=""1"" align='center' width=900><tr><th BGCOLOR='879773'><font size=2 align='center'>No.</font></th><th BGCOLOR='879773'><font size=2 >Field Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>POS Data</font></th><th BGCOLOR='879773'><font size=2 align='center'>Report Data</font></th><th BGCOLOR='879773'><font size=2>Status</font></th><th BGCOLOR='879773'><font size=2>Screenshot</font></th></tr>"
''		<th BGCOLOR='879773'><font size=2 align='center'>Report</font></th>
'	else
'	
'		StepCounter = 1
'		strTestFile = strPOSResultsPath & strTC & ".html"
'		
'		set objFile_Testhtml= fsObject.OpenTextFile(strTestFile,2,True)
'		objFile_Testhtml.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>No.</font></th><th BGCOLOR='879773'><font size=2 align='center'>Req/Response</font></th><th BGCOLOR='879773'><font size=2 >Field Name</font></th><th BGCOLOR='879773'><font size=2 align='center'>Receipt Data</font></th><th BGCOLOR='879773'><font size=2 align='center'>App Data</font></th><th BGCOLOR='879773'><font size=2>Status</font></th><th BGCOLOR='879773'><font size=2>Screenshot</font></th></tr>"
'	End if 
'	
'End Function
'

Public Function DateOperation(strOperation,strdate,strFormat)

	On error resume next

	Select Case strOperation

		Case "FormateDate"

			 DateOperation=FormatDate(strDate, strFormat)

		Case "YesterdayDate"

			If strDate<>"" Then
				sdate=strDate
			else
				sdate = Date()
			End If

			sy=year(sdate)
			sm=Month(sdate)

			If len(sm)=1 Then
			  sm="0"&sm
			End If

			sd=day(sdate)
			sd= sd-1
			If len(sd)=1 Then
			sd="0"&sd
			End If
			ddate= sd&"/"&sm&"/"&sy
			' ''msgbox ddate
			DateOperation=ddate

		Case "TommorowDate"

			If strDate<>"" Then
				sdate=strDate
			else
				sdate = Date()
			End If

			sy=year(sdate)
			sm=Month(sdate)

			If len(sm)=1 Then
				sm="0"&sm
			End If

			sd=day(sdate)
			sd= sd+1
			If len(sd)=1 Then
				sd="0"&sd
			End If
			ddate= sd&"/"&sm&"/"&sy
			' ''msgbox ddate
			DateOperation=ddate

		Case "TodayDate"

			If strDate<>"" Then
				sdate=strDate
			else
				sdate = Date()
			End If

			sy=year(sdate)
			sm=Month(sdate)
			If len(sm)=1 Then
				sm="0"&sm
			End If

			sd=day(sdate)
			If len(sd)=1 Then
				sd="0"&sd
			End If
			ddate= sd&"/"&sm&"/"&sy
			'''msgbox ddate
			DateOperation=ddate

	End Select


End Function


Public Function WriteConfig()

 

    'Convert it into YYYY-MM-DD
    sDateT=DateOperation("TodayDate","","")
    sDateTT = split(sDateT,"/")
    sDate = sDateTT(2)&"-"&sDateTT(1)&"-"&sDateTT(0)

 

    'config
    Config = "Results\Config.txt"
    ConfigDest = sCurrentDirectory & Config

 

    Set objFSO = createobject("Scripting.FileSystemObject")

 

      If objFSO.FileExists(ConfigDest) = False Then
           Set fsoConfig = fsObject.CreateTextFile (ConfigDest)
           sBatchNo = 1
      Else
        
       sBatchNo = 1
       Set fsoConfig = objFSO.OpenTextFile(ConfigDest,1)

 

       Do Until fsoConfig.AtEndOfStream
            strLine = fsoConfig.ReadLine
            If instr(strLine,"BatchNo")> 0 Then
                  sCurrentBatch = split(strLine,"==>")(1)
                  sBatchNo = sCurrentBatch+1
            End If
        Loop

 

      End If
        
        fsoConfig.Close
        Set fsoConfig = Nothing
        Set objFSO = Nothing
        
        wait 5
        
        Set objFSO1 = createobject("Scripting.FileSystemObject")
        
        Set fsoConfig1 = objFSO1.OpenTextFile(ConfigDest,2)

 

        'Timestamp==>
        fsoConfig1.writeline "Timestamp==>" & sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000)

 

        'Path==>
        fsoConfig1.writeline "Path==>" & Environment.Value("Solution") & "\" & Environment.Value("Market") & "\" & Environment.Value("ReleaseName") & "\" & sDate

 

        'BatchNo==>
        fsoConfig1.writeline "BatchNo==>"&sBatchNo

 

        'ReleaseName==>
        fsoConfig1.writeline "Release==>"&Environment.Value("ReleaseName")

 

        'ModuleName==>
        fsoConfig1.writeline "Module==>"&Environment.Value("ModuleName")
        
        'ResultFolder==>
        fsoConfig1.writeline "ResultFolder==>"&sFolderName
        
        fsoConfig1.Close
        Set fsoConfig1 = Nothing
        Set objFSO1 = Nothing
        
End Function




Public Function WriteLog(sLogLevel,sSource,sLogDescription)

 

    'Convert it into YYYY-MM-DD
    sDateT=DateOperation("TodayDate","","")
    sDateTT = split(sDateT,"/")
    sDate = sDateTT(2)&"-"&sDateTT(1)&"-"&sDateTT(0)

 


    If sLogDescription <> "" Then
        If sScreenShotpng = empty or sScreenShotpng = "" Then
            fsoFile.writeline sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000) & " => " & sLogLevel & ":" & sTCScriptName & " => " & sSource & " => " & sLogDescription
        Else
            fsoFile.writeline sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000) & " => " & sLogLevel & ":" & sTCScriptName & " => " & sSource & " => " & sLogDescription & " => " & sScreenShotpng
        End If
    else
        If sScreenShotpng = empty or sScreenShotpng = "" Then
            fsoFile.writeline sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000) & " => " & sLogLevel & ":" & sTCScriptName & " => " & sSource & " => " & sLogDescription
        Else
            fsoFile.writeline sDate & " " & FormatDateTime (Now,vbShortTime) & ":" & second(Now) & ":" & Int((Timer-Int(Timer)) * 1000000) & " => " & sLogLevel & ":" & sTCScriptName & " => " & sSource & " => " & sLogDescription & " => " & sScreenShotpng
        End If
        'fsoFile.writeline Now &  " => " & iStep & " : " & " : " & sStepName & " : " & sStepDescrption & " : " & sStatus
    End If

 


End Function

'******************************* HEADER ******************************************
' Description : The function to execute database queries
' Creator :  Pradeep Kumar
' Date : 12th November, 2021
' Last Modified On : 12th November, 2021
' Last Modified By : Pradeep Kumar 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function execute_db_query1(query)

	On error resume next
	
	Dim connect, sql, resultSet, path, text

	Set fs = CreateObject("Scripting.FileSystemObject")
	Set connect = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.RecordSet")
	Set execute_db_query = empty
	
	'connect.Open "Provider=sqloledb; Server=aewnw00235jm01.europe.shell.com; User Id=ASIA-PAC\Niharika.Vinjamuri; Password=; Database=SFN_SHELL_SPRINTQA_ID_REPORTS;Trusted_Connection=Yes;"
	connect.Open "provider=sqloledb;Server=aewnw00235jm01.europe.shell.com;User Id=ASIA-PAC\Niharika.Vinjamuri;Password=;Database=SFN_SHELL_SPRINTQA_ID_REPORTS;Trusted_Connection=Yes"
	
	objRecordSet.open query, connect

	Set execute_db_query1 = AddDBValuesToDict(objRecordSet)

	objRecordSet.Close
	connect.Close

End Function


'******************************* HEADER ******************************************
' Description : The function to add database values to dictionary
' Creator :  Niharika Vinjamuri
' Date : 30th November, 2021
' Last Modified On : 30th November, 2021
' Last Modified By :  
' Input Parameter : 
' Output Parameter : Dictionary
'******************************************** HEADER ******************************************
Public Function AddDBValuesToDict1(dbObj)

	Set objDict = CreateObject("Scripting.Dictionary")

 	For Iterator = 0 To dbObj.fields.count-1 Step 1
 		strKeyName = dbObj.fields.Item(Iterator).name
 		strValueName = dbObj.fields.Item(Iterator).value
		Call Add_Dictionary(objDict,strKeyName,strValueName)
	Next

   	Set AddDBValuesToDict = objDict
 
End Function



'******************************* HEADER ******************************************
' Description : The function to add db values to dictionary
' Creator :  Niharika Vinjamuri
' Date : 12th November, 2021
' Last Modified On : 
' Last Modified By :
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
Public Function AddDBValuesToDict(dbObj, rows)
	'Set objDict = empty
	Set objDict = CreateObject("Scripting.Dictionary")
	Dim curRowString
	dbObj.MoveFirst
	count_fields = dbObj.fields.count-1
	If rows="" Then
		rows = 1
	End if 
	curRow = 0
	Do Until dbObj.EOF
		curRowString = ""
		curRow = curRow + 1
		For ii = 0 to count_fields
			if instr(dbObj.fields(ii).Name,"Stamp")=0 then
				If objDict.Exists(dbObj.fields(ii).Name) and objDict.item(dbObj.fields(ii).Name) <> "" Then
					strValue = objDict(dbObj.fields(ii).Name) & "|"&dbObj.fields(ii).Value
					Call Update_Dictionary(objDict,dbObj.fields(ii).Name,strValue)
				else	
					Call Add_Dictionary(objDict,dbObj.fields(ii).Name,dbObj.fields(ii).Value)
				End If
			End  If
		Next
 		dbObj.MoveNext
 		
		If curRow = rows Then
			Exit Do
		End If	
 	Loop
   	Set AddDBValuesToDict = objDict
   	
End Function


'******************************* HEADER ******************************************
' Description : The function to execute database queries
' Creator :  Pradeep Kumar
' Date : 12th November, 2021
' Last Modified On : 17th December, 2021
' Last Modified By : Niharika Vinjamuri 
' Input Parameter : 
' Output Parameter : None
'******************************************** HEADER ******************************************
''
Public Function execute_db_query(query, rows, database)

	On error resume next
	
	Dim connect, sql, resultSet, path, text

	Set fs = CreateObject("Scripting.FileSystemObject")
	Set connect = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.RecordSet")
	'Set execute_db_query = empty
	mydbnames = Split(database,"_")
	dbname = mydbnames(ubound(mydbnames))
	
	If dbname = "BATCH" or  dbname = "OLTP" or dbname = "REPORTS" or dbname = "WWW" Then
		ccDBname = countryCode &"_"& dbname
	Else
		ccDBname = dbname
	End If
	
	envdbServerName = "aewnw00235jm01.europe.shell.com"
	
	If appName = "SFN" Then
		If appEnvName = "SPRINTQA" Then
			dbrepStr = "SFN_SHELL_SPRINTQA_"  & ccDBname
		ElseIf appEnvName = "RELEASEQA" Then
			dbrepStr = "SFN_SHELL_" & ccDBname
		End If
	ElseIf appName = "GFN" Then
		If appEnvName = "SPRINTQA" Then
			dbrepStr = "GFN_SHELL_SPRINTQA_" & ccDBname
		ElseIf appEnvName = "RELEASEQA" Then
			dbrepStr = "GFN_SHELL_" & ccDBname
			envdbServerName = "aewnw00235jm02.europe.shell.com"
		ElseIf appEnvName = "RD" Then
'			dbrepStr = "GFN_SHELL_RD_" & ccDBname
			dbrepStr = "GFN_SHELL_472_" & ccDBname
		End If
	ElseIf appName = "DEV" Then
		If appEnvName = "SPRINTQA" Then
			dbrepStr = "GFN_SHELL_SPRINTQA_" & ccDBname
		ElseIf appEnvName = "RELEASEQA" Then
			dbrepStr = "GFN_SHELL_" & ccDBname
		End If
	End If
'	msgbox envdbServerName
'	msgbox dbrepStr
	connect.Open "provider=sqloledb;Server="& envdbServerName &";User Id=ASIA-PAC\Venkata.SrinivasaRao;Password=;Database=" & dbrepStr & ";Trusted_Connection=Yes"
	
	'connect.Open "provider=sqloledb;Server=aewnw00235jm01.europe.shell.com;User Id=ASIA-PAC\Venkata.SrinivasaRao;Password=;Database=" & dbrepStr & ";Trusted_Connection=Yes"
	'connect.Open "Provider=sqloledb; Server=aewnw00235jm01.europe.shell.com; User Id=ASIA-PAC\Niharika.Vinjamuri; Password=; Database=SFN_SHELL_SPRINTQA_ID_REPORTS;Trusted_Connection=Yes;"
	
	objRecordSet.open query, connect
	Set execute_db_query = AddDBValuesToDict(objRecordSet, rows)

	objRecordSet.Close
	connect.Close

End Function


Public Function WebTbl_ClickCell_Dynamic(sBrowser,sPage,sWebTable,sObject,PropertyName,PropertyValue)
 	On error resume next  
  
	Set bBrowser = CreateDescriptionObject(sBrowser)
	Set bPage = CreateDescriptionObject(sPage)
	Set bWebTable = CreateDescriptionObject(sWebTable)
	'  Set bObject = CreateDescriptionObjectV2 (sObject,PropertyName & PropertyValue)  
	Set b = description.Create
	b("micclass").value = sObject
	b(PropertyName).value = PropertyValue
	Set sochild = Browser(bBrowser).Page(bPage).WebTable(bWebTable).ChildObjects(b)   
	
	If sochild.count<>0 Then 
		 sochild(0).Highlight
		 sochild(0).click
	End If
	
	ErrNumber = Err.Number
	ErrMessage = Err.Description
	 
	If ErrNumber<>0 Then	 	
		writelog "","","","","VerifyObjectExistInWebTbl(): VerifyObjectExistInWebTbl is unsucessfull! Object Name : "&PropertyValue & " Error Number : "&err.number & " Error Description : "&Err.description
	else
		writelog "","","","","VerifyObjectExistInWebTbl(): VerifyObjectExistInWebTbl is Successful. Object Name : "&PropertyValue
	End If
 
 End Function
 
 
' 
' 
' Public Function Read_Environment()
'
'	Set objEnvMas = CreateObject("Scripting.FileSystemObject")
'	Set bEnvDetail = objEnvMas.OpenTextFile (QA_Env)
'
'	Do While bEnvDetail.AtEndofStream <> True
'	    sEnvdata = bEnvDetail.ReadLine
'
'		arEnvtestdata = Split(sEnvdata,",")
'		If ucase(arEnvtestdata(6)) = ucase("Yes") Then
'			appEnvName = arEnvtestdata(0)
'			url = arEnvtestdata(1)
'			appSName = arEnvtestdata(2)
'			appDataExchange = appSName & arEnvtestdata(3)
'			appinbPath = appSName & arEnvtestdata(4)
'			appoutbPath = appSName & arEnvtestdata(5)
'			Exit Do
'		End If
'	Loop
' End Function
 
 'New Function with Reflex URL
'  Public Function Read_Environment()
'
'	Set objEnvMas = CreateObject("Scripting.FileSystemObject")
'	Set bEnvDetail = objEnvMas.OpenTextFile (QA_Env)
'
'	Do While bEnvDetail.AtEndofStream <> True
'	    sEnvdata = bEnvDetail.ReadLine
'
'		arEnvtestdata = Split(sEnvdata,",")
'		If ucase(arEnvtestdata(8)) = ucase("Yes") Then
'			appEnvName = arEnvtestdata(0)
'			url = arEnvtestdata(1)
'			appSName = arEnvtestdata(2)
'			appDataExchange = appSName & arEnvtestdata(3)
'			appinbPath = appSName & arEnvtestdata(4)
'			appoutbPath = appSName & arEnvtestdata(5)
'			reflex_url =  arEnvtestdata(6)
'			countryName =  arEnvtestdata(7)
'			Exit Do
'		End If
'	Loop
' End Function


Public Function Read_Environment()

	Set objEnvMas = CreateObject("Scripting.FileSystemObject")
	Set bEnvDetail = objEnvMas.OpenTextFile (QA_Env)

	Do While bEnvDetail.AtEndofStream <> True
	    sEnvdata = bEnvDetail.ReadLine

		arEnvtestdata = Split(sEnvdata,",")
		If ucase(arEnvtestdata(0)) = ucase("Yes") Then
			appName = arEnvtestdata(1)
			appEnvName = arEnvtestdata(2)
			countryName = arEnvtestdata(3)
			countryCode = arEnvtestdata(4)
			url = arEnvtestdata(5)
			appSName = arEnvtestdata(6)
			appDataExchange = appSName & arEnvtestdata(7)
			appinbPath = appSName & arEnvtestdata(8)
			appoutbPath = appSName & arEnvtestdata(9)
			reflex_url =  arEnvtestdata(10)
			api_url =  arEnvtestdata(11)
			client_ID =  arEnvtestdata(12)
			client_SecretKey =  arEnvtestdata(13)
			'Call gfnSystemdetails()
			Exit Do
		End If
	Loop
 End Function
 
 

'-----------New Function-----------------------------
Public Function update_db_query(query, rows, database)

	On error resume next
	
	Dim connect, sql, resultSet, path, text
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set connect = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.RecordSet")
	'Set execute_db_query = empty
	arrDb = split(database,"_")
	arrLen = ubound(arrDb)
	db_Name = arrDb(arrLen)

	database = ""
	database = appName & "_SHELL_"
	envdbServerName = "aewnw00235jm01.europe.shell.com"
	Select Case appEnvName
		Case "SPRINTQA"
			database = database & "SPRINTQA_"
		Case "RELEASEQA"
			database = database
			envdbServerName = "aewnw00235jm02.europe.shell.com"
		Case "RD"
		database = database & "RD_"
	End Select	
	If db_Name <> "IMPORT" or db_Name <> "EXPORT" or db_Name <> "MASTER" Then
		database = database & countryCode & "_" & db_Name
	else
		database = database & db_Name		
	End If
			
	
	connect.Open "provider=sqloledb;Server=" & envdbServerName &";User Id=ASIA-PAC\Venkata.SrinivasaRao;Password=;Database=" & database & ";Trusted_Connection=Yes"
	'connect.Open "Provider=sqloledb; Server=aewnw00235jm01.europe.shell.com; User Id=ASIA-PAC\Niharika.Vinjamuri; Password=; Database=SFN_SHELL_SPRINTQA_ID_REPORTS;Trusted_Connection=Yes;"
'	msgbox query
'	msgbox "provider=sqloledb;Server=" & envdbServerName &";User Id=ASIA-PAC\Venkata.SrinivasaRao;Password=;Database=" & database & ";Trusted_Connection=Yes"
	objRecordSet.open query, connect
'	Set execute_db_query = AddDBValuesToDict(objRecordSet, rows)

	objRecordSet.Close
	connect.Close

End Function

Function FFMXMLFileDataReader(filePath, searchText, searchtagName, searchendtagName)
	On error resume Next
	bFlag = True
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(filePath) Then
		pcheckFlag = False
		checkFlag = False
		'linecount = 0
		sectiondata = ""
		Set fileRead = fileSysObj.OpenTextFile(filePath,1,False,-1)
				Do until fileRead.AtEndOfStream
				content = fileRead.ReadLine	
					
				If instr(trim(content),trim(searchtagName)) >0 Then
					pcheckFlag = True
					tagstartpos = fileRead.Line
				End If
				
				If pcheckFlag and trim(content) = trim(searchText)  Then
					checkFlag = True
				End If
				If checkFlag and trim(content) = trim(searchendtagName)  Then
						tagEndpos = fileRead.Line
						Exit Do
				End  If
			
				'linecount = linecount + 1
			loop	
fileRead.Close	
		FFMXMLFileDataReader = tagstartpos & "|" & tagEndpos
	End If
End Function


Function validateSectionData(filePath,sposdata)
	Set objDict = CreateObject("Scripting.Dictionary")

lspos = cint(Split(sposdata,"|")(0)) 
uspos = cint(Split(sposdata,"|")(1))-1
On error resume Next
	bFlag = True
	Set fileSysObj = createObject("Scripting.FileSystemObject")
	If fileSysObj.FileExists(filePath) Then
		sectiondata = ""
		Set fileRead = fileSysObj.OpenTextFile(filePath,1,False,-1)
		lc = 0 
		Do until fileRead.AtEndOfStream
		content = trim(fileRead.ReadLine)
		
				If (fileRead.Line >= cint(lspos) or cint(lc) = cint(lspos) )  Then
'		fileRead.SkipLine
					Exit Do
				End If
			lc = lc + 1
		Loop	
		For Iterator = lspos To uspos Step 1
			If sectiondata = "" Then
				sectiondata = trim(content)
			Else
				sectiondata = sectiondata & "|" & trim(content)
			End If
			
			if content <> "" Then
				myKName = findKeyName(trim(content))
'				'msgbox myKName
				If objDict.Exists(myKName) and objDict.item(myKName) <> "" Then
					mystrValue = objDict(myKName) & "|"&trim(content)
					Call Update_Dictionary(objDict,myKName,mystrValue)
				else	
					Call Add_Dictionary(objDict,myKName,trim(content))
				End If
			End  If
			content = trim(fileRead.ReadLine)
		Next
End If
Call Add_Dictionary(objDict,"LineNumber",lspos) 
fileRead.Close
Set validateSectionData = objDict
End Function
