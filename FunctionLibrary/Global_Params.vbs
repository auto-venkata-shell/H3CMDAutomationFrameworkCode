environment.Value("sCurrentDirectory") = pathfinder.Locate(environment.Value("sCurrentDirectory"))
'Global dictionary to store all the descriptive objects and its properties
Public BOS_version
Public bSignOnFlag
bSignOnFlag = False
Public objORDict
'Set objORDict = CreateObject("Scripting.Dictionary")
' H3 Global variables
Public sys_CompanyID,sys_ClientCompanyNumber



Public countryName,countryCode
Public dynamicBrowser
Public dynamicPage
Public browserProp, pageProp
Public QA_Env
Public fileseqNo,filedatetime,DX873FName,cust_CreditLimit
Public gbldatavalidation
Public TextLogDest
Public logOverrideFile
Public fsoConfig
Public url
Public reflex_url
Public appName,appEnvName,appSName,appDataExchange,appinbPath,appoutbPath,api_url,client_ID,client_SecretKey
Public strBrowser
Public strPage
Public siteID
Public siteGroup
Public cmyName
Public fuelNetwork
Public strDesc
Public db_salesItemID
Public db_FeeRuleID
Public strStatus
Public db_PAN
Public db_CardID
Public db_CardIDNew
Public address
Public cust_erp_sub
Public cust_erp_sub1
Public cardGrpName
Public flagValueprop
Public dbName

'url = "https://aewnw00235iis1.europe.shell.com/SFNSprintLaunch/"

Public customerERP
Public cardPANNum
Public billReportDate
Public fileName

Public globalTable
set globalTable=CreateObject("Scripting.Dictionary")

Public strPDFFolder

Public sScreenShotpng
Public sTCScriptName

Public glbNumric_Key
Public ParamValDict
Public ORFilePath
Public BPFilePath
Public EMFilePath
Public ResultFilePath
Public TDFilePath
Public AutoFolderPath
Public ScriptFolderPath
Public strPOSResultsPath
Public TestResults
Public strindexFile
Public strSummaryFile
Public strSummaryFileAdhoc
Public strSummaryFileEOD
Public strSummaryFileAdhocdts
Public strSummaryFileEODdts
Public AdhocEODFlag
Public objFile_html
Public objFile_Testhtml
Public objFile_Parenthtml
Public fsObject
Public StepCounter
Public sFolderName
Public fsoFile
Public FuelConfig
Public FuelGradesConfig
Public SequencerFilePath
Public sRecoveryData
Public ReceiptValidationPath
Public categoryDic

Public DomSimulator
Public blgCompFlag
blgCompFlag = "STRCMPFLAG"

Public bRunFlag
Public sCurrentDirectory

Public sPOSSoftVersion
Public bSIMStartFlag
Public bForeCourtStartFlag
Public bFlagPOSVersion
Public StrChildIndex
Public ENV_Flag
Public ENV_URL
Public Comm_Browser1
Public Comm_Page1
Public Comm_Browser2
Public Comm_Page2
Public Comm_Browser3
Public Comm_Page3
Public Comm_Browser4
Public Comm_WebDialog1
Public Comm_WebDialog2
Public Comm_Browser_Sec
Public Comm_Page_Sec
Public Comm_Window1
Public Comm_Frame
Public objFuelDict
Public sURLTest

Public strTimeoutValue
Public strRole
Public strOrgID
Public sTaxSetVal
Public bSignOn
Public bSignOff
Public bSingleSignOnFlag
Public UserDetailPath
Public BOSInstructionFile

Public objDictReceiptData
Public objDictVoucher
Public objDictBOSExecInstruction
Public strParentIndexFile
Public fuelItemSetupFlag
Public objFuelCategory
Public objRefund
Public cCountInvItem

Set objDictReceiptData = CreateObject("Scripting.Dictionary")
Set objDictVoucher = CreateObject("Scripting.Dictionary")
Set objDictBOSExecInstruction = CreateObject("Scripting.Dictionary")

Set categoryDic = CreateObject("Scripting.Dictionary")
Set cCountInvItem = CreateObject("Scripting.Dictionary")

Public bBOSScriptsExistFlag

bBOSScriptsExistFlag=False

'initially the flag will be false, this is to start the simulator only once in any suite execution
bSIMStartFlag = False 
'initially the flag will be false, this is to start the forecourt only once in any suite execution
bForeCourtStartFlag= False

bFlagPOSVersion=False

if Environment.Value("sCurrentDirectory") <> "" then
	sCurrentDirectory = Environment.Value("sCurrentDirectory")
else
	sCurrentDirectory = ""
End if
''msgbox sCurrentDirectory

DomSimulator = "C:\Program Files (x86)\Doms\PumpSim2\DPUSim2.exe"

StepCounter = 1

Set fsObject= CreateObject("Scripting.FileSystemObject")

Set objORDict = CreateObject("Scripting.Dictionary")
Set ParamValDict = CreateObject("Scripting.Dictionary")

ORFolderPath = sCurrentDirectory & "ObjectRepository\"
''msgbox "ORFolderPath - " & ORFolderPath

ORFilePath = sCurrentDirectory & "ObjectRepository\Desc_OR_Namos.txt"
''msgbox "ORFilePath - " & ORFilePath

' To file path to fetch the BusinessProcess Associated to the test cases to be Executed
'BPFilePath = "Z:\Automation\BusinessProcessManger\BusinessProcess.txt"

' To file path to fetch the Test Cases to be Executed
EMFilePath =  sCurrentDirectory & "ExecutionManager\ExecutionManager.txt"
''msgbox "EMFilePath - " & EMFilePath

' The file path to fetch teh Login Credentials details
UserDetailPath = sCurrentDirectory & "ExecutionManager\UserDetail.txt"
QA_Env = sCurrentDirectory & "ExecutionManager\AppDetails.txt"

' To file path to save the screenshot of the Validation points
ScreenshotPath =  sCurrentDirectory & "Screenshots\"
''msgbox "ResultFilePath - " & ResultFilePath

' To file path to fetch the test data to execute
TDFilePath = sCurrentDirectory & "Test Data\"
''msgbox "TDFilePath - " & TDFilePath

' The automation Folder Path
AutoFolderPath =  sCurrentDirectory 
''msgbox "AutoFolderPath - " & AutoFolderPath

' The automation Folder Path
ScriptFolderPath =  sCurrentDirectory & "Scripts"

''msgbox "ScriptFolderPath - " & ScriptFolderPath

' The automation Test Execution results folder
TestResults =  sCurrentDirectory & "Results\"
''msgbox "TestResults - " & TestResults


' To Store Receipt File
ReceiptValidationPath=sCurrentDirectory & "ReceiptValidation\"

' To fetch the Recovery folder path
RecoveryPath =   sCurrentDirectory & "Recovery"
'This variable is used to save the prefix to the numeric object which can be used in the generic function
glbNumric_Key = "Btn_N_Num"

'This variable is used to pass the instructions from POS to EPS to complete any fuel transaction.
FuelConfig= sCurrentDirectory & "Test Data\FuelConfig\FuelConfig.txt"
FuelGradesConfig= sCurrentDirectory & "Test Data\FuelConfig\Fuel_Grades.xml"

SequencerFilePath = "C:\Temp\BP Pump\NORMAL1.SEQ"

BOSInstructionFile= sCurrentDirectory & "Test Data\BOSInstruction\BOSInstructionFile.txt"

sRecoveryData =""

sCurrencyName= "R"
sCurrency= "South African Rand"
cCancelReceiptFileName=""

sTaxSetVal="15.0000%"
'ENV_Flag="CH-PREP"
strPublicRole="Site Manager"

strTimeoutValue=35

' To set the URL path
'sURLTest = "http://sww-x64test2.zarp.redprairie.shell.com/portal/welcome.htm"

bSignOn = "True"
bSignOff="True"
bSingleSignOnFlag="False"

Set objFuelDict = CreateObject("Scripting.Dictionary")

objFuelDict.Add "V-Power 95 (I)","1"
objFuelDict.Add "Unleaded Extra 93","2"
objFuelDict.Add "Diesel Extra","3"
objFuelDict.Add "V-Power Diesel","4"
objFuelDict.Add "PumpNumber","2"


