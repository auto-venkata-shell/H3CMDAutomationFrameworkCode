Set obj = CreateObject("Quicktest.Application")
path =  "F:\users\Venkata.SrinivasaRao\Desktop\FinalCodes\H3UFTAutomationFrameworkCode\Scripts\DriverScript"
'// Set AOM file and test script path
obj.Visible = True
obj.Launch
obj.Options.Run.ImageCaptureForTestResults = "OnError"
obj.Options.Run.RunMode = "Fast"
obj.Options.Run.ViewResults = False
obj.Open path
	obj.Test.Run qtpResults,True
	VarStatus = obj.Test.LastRunResults.Status
obj.Test.Close