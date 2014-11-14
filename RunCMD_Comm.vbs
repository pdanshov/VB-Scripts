
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' '							Peter Danshov
' '					pdanshv@gmail.com - 11.12.14
' '		This program opens a cmd command line interface
' '		and runs a command.
' '
' '
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Dim oShell
On Error Resume Next
Set oShell = WScript.CreateObject("WScript.Shell")
If Err.Number <> 0 Then
	MsgBox "VBS: " & Err.Number, 65, "VB"
	Wscript.Quit
End If
'oShell.run "cmd.exe /C copy ""S:Claims\Sound.wav"" ""C:\WINDOWS\Media\Sound.wav""",1,true
oShell.run "c:\windows\system32\cmd.exe /c wasset unattend receive 0003",1,true
On Error Goto 0
Set oShell = Nothing

