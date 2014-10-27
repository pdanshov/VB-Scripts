
 '---------------------------------------------------------------------------------------
 '              : Terminates a process. First checking to see if it is running or not.
 '              : Uses WMI (Windows Management Instrumentation) to query all running processes
 '              : then terminates ALL instances of the specified process
 '              : held in the variable strTerminateThis.
 '              :
 '              : ***WARNING: This will terminate a specified running process,use with caution!.
 '              : ***Terminating certain processes can effect the running of Windows and/or
 '              : ***running applications.
 '---------------------------------------------------------------------------------------
Dim strTerminateThis' As String 'The variable to hold the process to terminate
Dim objWMIcimv2' As Object
Dim objProcess' As Object
Dim objList' As Object
Dim intError' As Integer
 
if WScript.Arguments.Count = 0 then
    WScript.Echo "Missing parameters"
end if
'strTerminateThis = "wscript.exe" 'Process to terminate,
strTerminateThis = WScript.Arguments(0)
 'change notepad.exe to the process you want to terminate
 
Set objWMIcimv2 = GetObject("winmgmts:" _ 
& "{impersonationLevel=impersonate}!\\.\root\cimv2") 'Connect to CIMV2 Namespace
 
Set objList = objWMIcimv2.ExecQuery _ 
("select * from win32_process where name='" & strTerminateThis & "'") 'Find the process to terminate
 
 
If objList.Count = 0 Then 'If 0 then process isn't running
	'MsgBox strTerminateThis & " is NOT running." & vbCr & vbCr _ 
	'& "Exiting procedure.", vbCritical, "Unable to continue" 
	Set objWMIcimv2 = Nothing 
	Set objList = Nothing 
	Set objProcess = Nothing 
	'Exit Sub 
Else 
	 'Ask if OK to continue
'    Select Case MsgBox("Are you sure you want to terminate this running process?:" _ 
'        & vbCrLf & "" _ 
'        & vbCrLf & "Process name: " & strTerminateThis _ 
'        & vbCrLf & "" _ 
'        & vbCrLf & "Note:" _ 
'        & vbCrLf & "Terminating certain processes can effect the running of Windows" _ 
'        & "and/or running applications. The process will terminate if you OK it, WITHOUT " _ 
'        & "giving you the chance to save any changes in anything that is running in the specified process above." _ 
'        , vbOKCancel Or vbQuestion Or vbSystemModal Or vbDefaultButton1, "WARNING:") 
'         
'    Case vbOK 
		 'OK to continue with terminating the process
	For Each objProcess In objList 
		 
		intError = objProcess.Terminate 'Terminates a process and all of its threads.
		 'Return value is 0 for success. Any other number is an error.
		If intError <> 0 Then 
			MsgBox "ERROR: Unable to terminate that process.", vbCritical, "Aborting" 
			'Exit Sub 
		End If 
	Next 
	 'ALL instances of specified process (strTerminateThis) has been terminated
	Call MsgBox("ALL instances of process " & strTerminateThis & " has been successfully terminated.", _ 
	vbInformation, "Process Terminated") 
	 
	Set objWMIcimv2 = Nothing 
	Set objList = Nothing 
	Set objProcess = Nothing 
	'Exit Sub 
		 
'    Case vbCancel 
'         'NOT OK to continue with the termination, abort
'        Set objWMIcimv2 = Nothing 
'        Set objList = Nothing 
'        Set objProcess = Nothing 
'        Exit Sub 
'    End Select 
End If