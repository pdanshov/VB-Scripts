Dim objShell: Set objShell = wscript.createObject("wscript.shell")
Dim fsObj: Set fsObj = CreateObject("Scripting.FileSystemObject")
Dim vbsFile: Set vbsFile = fsObj.OpenTextFile("C:\Admin\DealFunctions.vbs", 1, False)
Dim myFunctionsStr
myFunctionsStr = vbsFile.ReadAll
vbsFile.Close
Set vbsFile = Nothing
Set fsObj = Nothing
ExecuteGlobal myFunctionsStr
OpenFileMaker "User","Pass"
