' Run in debug mode from cmd
' 		cscript //D HardwareAudit.vbs
' 		wscript //D HardwareAudit.vbs
' I recommend scheduling your script as follows:
' 	C:\Windows\System32\CScript.exe //Nologo //B X:\PathToYourScript\YourScript.vbs
' To resolve this issue, define the system environment variable using the Process environment. For an example about how to use the Process environment variable, view the following Visual Basic Script samples. 
' This script works when an administrator runs it but does not work when a user who does not have administrator permissions runs it.
'		Set WSHShell = WScript.CreateObject("WScript.Shell")
'		Set WSHEnv = WSHShell.Environment
'		WScript.Echo WSHEnv("OS")
' NOTE: Notice that since there is no type specified after Environment that it will default to the "System" type. 
' This script works when either an administrator or non-administrator runs it.
'		Set WSHShell = WScript.CreateObject("WScript.Shell")
'		Set WSHEnv = WSHShell.Environment("Process")
'		WScript.Echo WSHEnv("OS")
' http://stackoverflow.com/questions/13620748/msgbox-vs-msgbox-in-vbscript
' 1 Click the Start menu and type "regedit" on the search box. This will launch the Windows Registry Editor program. 
' 2 Navigate to the following Registry entry: 
'		HKEY_CURRENT_USER\Software\Microsoft\Windows Script Host\Settings 
' 3 Select the "Enabled" entry in the right window pane. If this entry does not exist, right-click anywhere in the right window pane and select "New" followed by "DWORD Value." Name the value "Enabled." 
' 4 Right-click the "Enabled" entry and click "Modify." 
' 5 Change the number in the "Value" box to "1." This will re-enable WSH. 
' Note: If WSH has been disabled for all users on your computer, use this same process to restore it except instead of using 
'		"HKEY_CURRENT_USER\Software\Microsoft\Windows Script Host\Settings" go to the
'		"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows Script Host\Settings" key.
Dim wshshell
Set wshshell = WScript.CreateObject("WScript.Shell")
Set WSHEnv = WSHShell.Environment("Process")
wshshell.run """C:\Program Files\Mozilla Firefox\firefox.exe"" https://prod.citytime.nycnet/",1,False

MsgBox "Hello World!", 65, "MsgBox Example"

Set wshshell = Nothing
wscript.quit