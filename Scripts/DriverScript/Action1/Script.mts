'******************************************** HEADER ******************************************
' Description : The Driver Script for BOS AUTOMATION
' Creator : Syed Shafi
' Date : 10th Jul,2017
' Last Modified On : 05th Sep, 2017 
' Last Modified By : Hemaraja 
' Input Parameter : None
' Output Parameter : None
'******************************************** HEADER ******************************************
Dim dictObject8
Dim objExMas
Dim bExefile
Dim exeList()
Dim sExeMasFile
Dim jCount
Dim tcCounter
Dim sExedata
Dim arrExedata
Dim sTCValue
Dim sTCName
Dim sScriptPath
Dim strindexFile
Dim strExecStartTime, strExeEndtime, sTestStart, sTestEnd
Dim qtpObj

On error resume next

strExecStartTime = now

'Call Read_UserDetails

Call Index_initialize

Call Descriptive_Object_Parser
Call Read_Environment


' To Fetch the TestCase to be Executed
sExeMasFile = EMFilePath
  
Set objExMas = CreateObject("Scripting.FileSystemObject")
Set bExefile = objExMas.OpenTextFile (sExeMasFile)

ReDim Preserve exeList(200)
jCount=0

Call WriteConfig

Do While bExefile.AtEndofStream <> True
    sExedata = bExefile.ReadLine
       
    'If instr(sExedata,"Yes") Then
    'Add into Array only when Execute is YES and it is a POS script
    If instr(sExedata,"Yes") <> 0 Then 'and instr(sExedata,"BOS") <> 0 Then
		arrExedata = Split(sExedata,"Yes") 
		k = 0 
		If arrExedata(k)<>"" Then
			exeList(jCount) = arrExedata(k)
			jCount= jCount+1
		End If	
    End If
Loop

'Create a consolidated receipt dictionary for report validations
'Call CreateConsolidatedReceiptDictionary
	
' To Fetch the Keyword from Business Process
For tcCounter = 0 To jCount-1
	
	sTestStart = Now
	
    sTCValue=split(exeList(tcCounter),",")
    sTCName=sTCValue(0)
    
    'read the test data values from the test data csv	
		'''''''''''''''''''  To handle Single Sign On '''''''''''''''''''''''	
	If sTCName="SignOn" then
	   bSignOff= False
	   bSignOn = True
	   bSingleSignOnFlag = True
	   
   	ElseIf sTCName ="SignOff" then
   
      bSignOff= True
      Call BOS_Logout
      bSignOn = True
      bSingleSignOnFlag = False
   
   	Else	 
		
      Call ReadTestDataCSV (sTCName & ".csv")
	End If
    ''''''''''''  
	
	'read the test data values from the test data csv	
'	Call ReadTestDataCSV (sTCName & ".csv")
	sTCScriptName = sTCName
	sBOSTCName =  sTCName
	'initialize Test specific result file
	Initialize_TestHTML sBOSTCName

	sScriptPath = ScriptFolderPath&"\"&sTCName&".vbs"
	'Test Run Flag set to default False
	bRunFlag = True
	'Call f("","")
	'OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebButton_Refresh", "Click", ""
	
	'OperateOnObject "WebBrowser_PTShell_INDO", "WebPage_PTShell_INDO", "", "WebElement_Drop", "Click", ""
	

	
	ExecuteFile sScriptPath
	wait 5
	 
	'Cleanup the Param before next script in the execution set
	Set ParamValDict = Nothing
	Set ParamValDict = Create_Dictionary
	
	sTestEnd = Now
	duration = DateDiff ("n",sTestStart,sTestEnd)	
	Counter = tcCounter + 1
	
	if bRunFlag=False then
		pStatus="Failed"
		objFile_html.writeline  "<tr><td align=""center"">" & Counter & "</td><td><a href=""" & sBOSTCName & ".html" & """>" & sBOSTCName & "</a></td><td align=""center""><font color=""blue"">"&duration&"</td></font><td BGCOLOR=""#FF0000"">FAILED</td></tr>"
		writelog "IndexError","BOS_"& environment.Value("sCurrentDirectory")&"Action","IndexLevel Status for Script "&sTCName&" is "&pStatus
		bRunFlag = True
	else
		pStatus="Passed"
		objFile_html.writeline  "<tr><td align=""center"">" & Counter & "</td><td><a href=""" & sBOSTCName & ".html" & """>" & sBOSTCName & "</a></td><td align=""center""><font color=""blue"">"&duration&"</td></font><td BGCOLOR=""#3ADF00"">PASSED</td></tr>"
		writelog "IndexInfo","BOS_"& environment.Value("sCurrentDirectory")&"Action","IndexLevel Status for Script "&sTCName&" is "&pStatus
	end if
	
	DeInitialize_TestHTML
	
Next

strExeEndtime = now

'objFile_html.writeline "</table>" 	

iTotalTime = DateDiff ("n",strExecStartTime,strExeEndtime)
'
'objFile_html.writeline "<h4 align='center'><u>Execution Summary</u></h4>"
'objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>StartTime</font></th><th BGCOLOR='879773'><font size=2 >End Time</font></th><th BGCOLOR='879773'><font size=2 align='center'>Total Execution(Mins)</font></th></tr>"
'objFile_html.writeline  "<tr><td align=""center"">" & strExecStartTime & "</td><td>"& strExeEndtime & "</a></td><td align=""center""><font color=""blue"">"&iTotalTime&"</td></tr>"
'objFile_html.writeline "</table>"
'
'objFile_html.writeline "<br><br>"
'objFile_html.writeline "<h4 align='center'><u>POS Software Version</u></h4>"
'objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>POS Software Info</font></th></tr>"
'objFile_html.writeline  "<tr><td align=""center"">"& sPOSSoftVersion & "</td></tr>"
'objFile_html.writeline "</table>"
objFile_html.writeline "</table>"
objFile_html.writeline "<br><br>"
objFile_html.writeline "<h4 align='center'><u> Total Automation Execution Time: "& iTotalTime &" Minutes </u></h4>"
'objFile_html.writeline  "<table border=""1"" align='center' width=500><tr><th BGCOLOR='879773'><font size=2 align='center'>BOS Release Info</font></th></tr>"
'objFile_html.writeline  "<tr><td align=""center"">"& BOS_version & "</td></tr>"'
objFile_html.writeline "</table>"



'Deinitilize the html file object and open the index file in iexplore
Call DeInitialize
'Add_Dictionary objDictBOSExecInstruction,"BOSEXECSTATUS","BOSDOESNOTEXIST"
'Update_Dictionary objDictBOSExecInstruction,"RESULTFILE",""
'Update_Dictionary objDictBOSExecInstruction,"XMLFILE",""
'Update_Dictionary objDictBOSExecInstruction,"ONLYBOS",""
'Call CreateBOSInstructionFile
'''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''   

fsObject.CopyFile TextLogDest,logOverrideFile

Dim oShell
Set oShell=createobject("wscript.shell")
sAgentFolderPath = environment.Value("sCurrentDirectory") & "agent"
oShell.CurrentDirectory = sAgentFolderPath
'oShell.CurrentDirectory = "C:\Users\Autopc\Desktop\CARiNode_32Bit\retail-cari-agent\FinalDashBoard\agent"
wait 2
'oShell.run "cmd /K ""CARiNode.exe"""
wait 2





















