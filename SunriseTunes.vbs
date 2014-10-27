
'Set objShell = CreateObject("WScript.Shell")

	'objShell.Run("""T:\FileMaker\Production\Deal Trading.fp7"""), 1, 0

'	Do Until objShell.AppActivate("iTunes")
'		Wscript.sleep 500
'	Loop
	
'	If objShell.AppActivate("iTunes") Then
'		objShell.AppActivate "iTunes"
		
'Option Explicit

'Dim Shell, WMI, wql, process

'Set Shell = CreateObject("WScript.Shell")
'Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'iTunes.exe"

'For Each process In WMI.ExecQuery(wql)
'    Shell.AppActivate process.ProcessId
    'Shell.SendKeys "% r"
    'Shell.SendKeys "% x"
'	Wscript.sleep 100
'	Shell.SendKeys "{ALT}"
'	Wscript.sleep 100
'	Shell.SendKeys "{RIGHT}"
'	Wscript.sleep 100
'	Shell.SendKeys "{RIGHT}"
'	Wscript.sleep 100
'	Shell.SendKeys "{RIGHT}"
'	Wscript.sleep 100
'	Shell.SendKeys "{DOWN}"
'	Wscript.sleep 100
'	Shell.SendKeys "{UP}"
'	Wscript.sleep 100
'	Shell.SendKeys "{UP}"
'	Wscript.sleep 100
'	Shell.SendKeys "{UP}"
'	Wscript.sleep 100
'	Shell.SendKeys "{RIGHT}"
'	Wscript.sleep 100
'	Shell.SendKeys "{ENTER}"
'	Wscript.sleep 100
'	Shell.SendKeys "{TAB}"
'	Wscript.sleep 100
'	Shell.SendKeys "{TAB}"
'	Wscript.sleep 100
'	Shell.SendKeys "{ENTER}"
'	Wscript.sleep 100
'Next

Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

intProcessId = ""
For Each Process In Processes
    If StrComp(Process.Name, "iTunes.exe", vbTextCompare) = 0 Then
        intProcessId = Process.ProcessId
        Exit For
    End If
Next

If Len(intProcessId) > 0 Then
    With CreateObject("WScript.Shell")
        .AppActivate intProcessId
        '.SendKeys "%{F4}"
		Wscript.sleep 100
		.SendKeys "%" 'Alt
		Wscript.sleep 100
		.SendKeys "{UP}" 'Activate Menu
		Wscript.sleep 100
		.SendKeys "{UP}" 'IndieVolume
		Wscript.sleep 100
		.SendKeys "{UP}" 'Close
		Wscript.sleep 100
		.SendKeys "{UP}" 'Maximize
		Wscript.sleep 100
		.SendKeys "{ENTER}"
		Wscript.sleep 100
		.SendKeys "%" 'Alt
		Wscript.sleep 100
		.SendKeys "{RIGHT}" 'Edit
		Wscript.sleep 100
		.SendKeys "{RIGHT}" 'View
		Wscript.sleep 100
		.SendKeys "{RIGHT}" 'Controls
		Wscript.sleep 100
		.SendKeys "{DOWN}" 'Activate Menu
		Wscript.sleep 100
		.SendKeys "{UP}" '
		Wscript.sleep 100
		.SendKeys "{UP}" '
		Wscript.sleep 100
		.SendKeys "{UP}" '
		Wscript.sleep 100
		.SendKeys "{RIGHT}"
		Wscript.sleep 100
		.SendKeys "{ENTER}"
		Wscript.sleep 100
		.SendKeys "{TAB}"
		Wscript.sleep 100
		.SendKeys "{TAB}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{UP}"
		Wscript.sleep 100
		.SendKeys "{ENTER}"
		Wscript.sleep 100
    End With
End If

'Wscript.sleep 100
'objShell.SendKeys "{ALT}"
'Wscript.sleep 100
'objShell.SendKeys "{RIGHT}"
'Wscript.sleep 100
'objShell.SendKeys "{RIGHT}"
'Wscript.sleep 100
'objShell.SendKeys "{RIGHT}"
'Wscript.sleep 100
'objShell.SendKeys "{DOWN}"
'Wscript.sleep 100
'objShell.SendKeys "{UP}"
'Wscript.sleep 100
'objShell.SendKeys "{UP}"
'Wscript.sleep 100
'objShell.SendKeys "{UP}"
'Wscript.sleep 100
'objShell.SendKeys "{RIGHT}"
'Wscript.sleep 100
'objShell.SendKeys "{ENTER}"
'Wscript.sleep 100
'objShell.SendKeys "{TAB}"
'Wscript.sleep 100
'objShell.SendKeys "{TAB}"
'Wscript.sleep 100
'objShell.SendKeys "{ENTER}"
'Wscript.sleep 100

'	End If

'Set objShell = Nothing
Set Shell = Nothing
WScript.Quit
