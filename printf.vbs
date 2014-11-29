


REM function Subst(str, args)
    REM dim res
    REM dim i
    REM res = str

    REM for i = 0 to UBound(args)
        REM res = Replace(res, "%" & CStr(i+1) & "%" , args(i) )
    REM next

    REM res = Replace(res, "\n", vbCrLf)
    REM res = Replace(res, "\t", vbTab)

    REM Subst = res
REM end function

REM 'And an example of use:

REM theSelector = "SELECT * FROM %1% WHERE [aKey] >= (DATESERIAL(%2%, %3%, %4% ) + TIMESERIAL( %5%, 0, 0 )) AND [aKey] <= (DATESERIAL(%2%, %3%, %4% ) + TIMESERIAL( %5%, 59, 59 ))"
REM theSelect = Subst( theSelector, Array( theTable, theYear, theMonth, theDay, theHour ) )
REM call Subst( theSelector, Array( theTable, theYear, theMonth, theDay, theHour ) )

REM theSe = Subst( "One Two Three", Array( 1, 2, 3, 4) )

REM wscript.echo theSelect
REM wscript.echo theSe


Option Explicit
 
Dim arrTemp : arrTemp = Array("something", "something else", "another thing", "this thing", "is it time to go home yet")
Dim strTemp


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Each strTemp In arrTemp 
  WScript.Echo strTemp & " " & Now
Next
WScript.Echo String(40, "=")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' try to get the longest string length
Dim intSpace : intSpace = 0
For Each strTemp In arrTemp
  If Len(strTemp) > intSpace Then intSpace = Len(strTemp)
Next
' add some extra spaces
intSpace = intSpace + 5
' loop through strings putting in the necessary spaces
For Each strTemp In arrTemp
  WScript.Echo strTemp & Space(intSpace - Len(strTemp)) & Now
Next
WScript.Echo String(40, "=")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'or if you know the longest string length will be less than a certain number
intSpace = 32
For Each strTemp In arrTemp
  WScript.Echo strTemp & Space(intSpace - Len(strTemp)) & Now
Next




