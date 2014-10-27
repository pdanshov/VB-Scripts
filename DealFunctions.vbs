

Function LoadStringFromFile(filename)
	Const fsoForReading = 1
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, fsoForReading)
    LoadStringFromFile = f.ReadAll
    f.Close
End Function


Sub SaveStringToFile(filename, text)
	Const fsoForWriting = 2
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, fsoForWriting)
    f.Write text
    f.Close
End Sub


'user is used twice, in the login logic and in the firstuseropen logic
Sub OpenFileMaker(user, password)
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run("""\\172.28.1.231\DealFreeNAS\FileMaker\Production\Deal Trading.fp7"""), 1, 0
	'Wait for login window to appear
	Do Until objShell.AppActivate("FileMaker Pro Advanced")
		Wscript.sleep 500
	Loop
	'ActiveWindow.FindByName("sbar", "GuiStatusbar").text
	'FileMaker Pro
	'"Deal Trading.fp7" is currently in use and could not be opened. If the file is shared, you can use the Open Remote command to open the file on the network. (If you've opened the file before, check the Open Recent menu.)
	'OK

	'If first user start window appears Do:
	If objShell.AppActivate("Open ""Deal Trading""") Then
		
		Login user, password
		
		objShell.AppActivate("FileMaker Pro Advanced")
		
		'Wait for second (first in login subroutine above) shared file window prompt
		Do Until objShell.AppActivate("FileMaker Pro")
			Wscript.sleep 500
		Loop
		
		'select it and enter correct sequence of keys to proceed
		objShell.AppActivate("FileMaker Pro")
		Wscript.sleep 200
		objShell.SendKeys "{TAB}"
		Wscript.sleep 500
		objShell.SendKeys "{ENTER}"
		
		SaveStringToFile "\\172.28.1.231\DealFreeNAS\FileMaker\vbopener.txt", user
		Dim lastOpen
		'lastOpen = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
		lastOpen = user & " " & Date & " " & Time
		SaveStringToFile "\\172.28.1.231\DealFreeNAS\FileMaker\vblastopen.txt", lastOpen
			
	'else if multi user window appears Do:
	ElseIf objShell.AppActivate("FileMaker Pro") Then
		
		Dim firstuser
		firstuser = LoadStringFromFile("\\172.28.1.231\DealFreeNAS\FileMaker\vbopener.txt")
		
		'Wscript.echo "Multi User"
		objShell.AppActivate("FileMaker Pro")
		Wscript.sleep 500
		objShell.SendKeys "{TAB}"
		Wscript.sleep 500
		objShell.SendKeys "{ENTER}"
		Wscript.sleep 100
		objShell.SendKeys "^+o" 'send simultaneous ctrl+shift+o shortcut for open remote
		
		Do Until objShell.AppActivate("Open Remote File")
			Wscript.sleep 500
		Loop
		
		objShell.AppActivate("Open Remote File")
		Wscript.sleep 100
		
		MultiUserOne firstuser
		Wscript.sleep 50
		
		objShell.SendKeys "{TAB}"
		Wscript.sleep 130
	    objShell.SendKeys " "
	    Wscript.sleep 100
	    objShell.SendKeys "{ENTER}"
		
		Do Until objShell.AppActivate("Open ""Deal Trading""")
			Wscript.sleep 500
		Loop
		
		login user, password
		
		Do Until objShell.AppActivate("Open File")
			Wscript.sleep 500
		Loop
		'select it and enter correct sequence of keys to proceed
		objShell.AppActivate("Open File")
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Right Nav Bar
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Left Side Bar
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Center File Box
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'File name Field
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Files Of Type Field
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Open
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Cancel
		Wscript.sleep 100
		objShell.SendKeys "{TAB}" 'Remote
		Wscript.sleep 100
		objShell.SendKeys " "
		Wscript.sleep 50
		
		Do Until objShell.AppActivate("Open Remote File")
			Wscript.sleep 500
		Loop
		
		objShell.AppActivate("Open Remote File")
		Wscript.sleep 100
		
		MultiUserTwo firstuser
		Wscript.sleep 50
		
		objShell.SendKeys "{TAB}"
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}"
		Wscript.sleep 100
		objShell.SendKeys "{ENTER}"
		
		'This is unnecessary, if a second login box pops up, it's because the user account doesn't exist in both databases
		'LoginTwo user, password
		
		End If

		'Dim proc As Process = Process.Start("\\172.28.1.231\DealFreeNAS\FileMaker\Production\Deal Trading.fp7")
		'proc.WaitForInputIdle()
		'While objShell.Busy
		'	Wscript.sleep 500
		'Wend
		'If FindWindow("AutoCad.Application.18", "acad.exe") Then
		'FileMaker Pro Advanced
		'FileMaker Pro Advanced.exe
		'FileMaker Pro
		'Wscript.echo "Outside"
		'If FindWindow("FileMaker Pro") Then
		'	Wscript.echo "Inside FileMaker Pro"
		'	'System.Threading.Thread.Sleep(1000)
		'	objShell.AppActivate "Internet Explorer"
		'	objShell.SendKeys "{TAB}"
		'	objShell.SendKeys "%{TAB}"
		'	'objShell.SendKeys "{ENTER}"
		'End If

	Set objShell = Nothing
	WScript.Quit
End Sub


Sub Login(user, password)
	'When it appears, go to it and enter password & correct sequence of
	'keys to proceed
	objShell.AppActivate("Open ""Deal Trading""")
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'Account Name
	Wscript.sleep 100
	objShell.SendKeys user
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'Password
	Wscript.sleep 100
	objShell.SendKeys password
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'Change Password
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'OK
	Wscript.sleep 100
	objShell.SendKeys "{ENTER}"
	'Wait for first shared file window prompt
	Do Until objShell.AppActivate("FileMaker Pro")
		Wscript.sleep 500
	Loop
	'select it and enter correct sequence of keys to proceed
	objShell.AppActivate("FileMaker Pro")
	Wscript.sleep 200
	objShell.SendKeys "{TAB}"
	Wscript.sleep 500
	objShell.SendKeys "{ENTER}"
End Sub


Sub LoginTwo(user, password)
	'When it appears, go to it and enter password & correct sequence of
	'keys to proceed
	objShell.AppActivate("Open ""Deal Trading""")
	Wscript.sleep 100
	objShell.SendKeys "+{TAB}" 'Account Name
	Wscript.sleep 100
	objShell.SendKeys user
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'Password
	Wscript.sleep 100
	objShell.SendKeys password
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'Change Password
	Wscript.sleep 100
	objShell.SendKeys "{TAB}" 'OK
	Wscript.sleep 100
	objShell.SendKeys "{ENTER}"
	'Wait for first shared file window prompt
	Do Until objShell.AppActivate("FileMaker Pro")
		Wscript.sleep 500
	Loop
	'select it and enter correct sequence of keys to proceed
	objShell.AppActivate("FileMaker Pro")
	Wscript.sleep 200
	objShell.SendKeys "{TAB}"
	Wscript.sleep 500
	objShell.SendKeys "{ENTER}"
End Sub


Sub MultiUserOne(firstuser)
	Select Case firstuser
	  Case "John"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		
	  Case "Kay"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Kay
		Wscript.sleep 100
		
	  Case "Liana"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Kay
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Liana
		Wscript.sleep 100
		
	  Case "Peter"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Kay
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Liana
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Peter
		Wscript.sleep 100
		
	  Case else
		MsgBox "Unknown User: " + firstuser
		
	End Select
End Sub


Sub MultiUserTwo(firstuser)
	Select Case firstuser		
	  Case "John"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		
	  Case "Kay"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Kay
		
	  Case "Liana"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Kay
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Liana
		
	  Case "Peter"
		objShell.SendKeys "{TAB}" 'John
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Kay
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Liana
		Wscript.sleep 100
		objShell.SendKeys "{DOWN}" 'Peter
		
	  Case else
		MsgBox "Unknown User: " + firstuser
		
	End Select
End Sub

