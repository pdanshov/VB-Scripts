
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''			Peter Danshov
'''''			11.06.14
'''''			pdanshv@gmail.com
'''''			This script reads .mdb
'''''			files and saves queried
'''''			rows to a txt file
'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim cn 
Dim rs
Dim strResult
Dim headLine

strFile = "C:\Program Files\Trading Partner\System\Recon.mdb"

strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFile & ";"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Open strCon

strSQL = "SELECT * FROM Tracking WHERE AckCode Is Null AND Direction='Out' ORDER BY DateSent DESC"
'strSQL = "SELECT * FROM Tracking WHERE AckCode='' OR AckCode=NULL ORDER BY DateSent" _
'& "WHERE CrDate Between Now() And Date()-1 " _
'& "WHERE AckCode='' " _
'& "AND OtherField='abc' " _
'& "OR AckCode=NULL " _
'& "AND PartNumber=1 " _
'& "ORDER BY CrDate, PartNumber"
'& "ORDER BY DateSent"

rs.Open strSQL, cn

'Set xl = CreateObject("Excel.Application")
'Set xlBk = xl.Workbooks.Add
'With xlbk.Worksheets(1)
'    For i = 0 To rs.Fields.Count - 1
'        .Cells(1, i + 1) = rs.Fields(i).Name
'    Next
'    .Cells(2, 1).CopyFromRecordset rs
'    .Columns("B:B").NumberFormat = "m/d/yy h:mm"
'End With
'xl.Visible=True

'objField = objRecordset.Fields.Item("ProductID")
'objField = objRecordset.Fields("ProductID")
'objField = objRecordset.Fields.Item(0)
'objField = objRecordset.Fields(0)

'0		1		2		3			4			5	6		7		8		9		10		11		12	13		14		15		16	17		18	19		20		21		22
'Direction	Sender		Receiver	DateSent		DateAcked		AckCode	EnvControl	GroupControl	GroupVersion	GroupError	TSControl	TSIDCode	TSError	SegIDCode	SegError	DEIDCode	DEError	Partner		Key	GroupSender	GroupReceiver	AckEnvControl	AckGroupControl
'Out		2906196701	6112391050	11/6/2014 4:34:00 AM	11/6/2014 4:03:00 AM	A	000000017	17		004010VICS	000170001			856										DILLARDS	6923	2906196701	6112391050	000000021	21
'Partner 17	TSIDCode 11	Key 18 		DateSent 3

Dim OrderedArray, Counter
OrderedArray=Array(17,11,18,3) 'the order in which the Headers and Fields should be written
Counter=0

Dim fso, MyFile, HdrFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile("C:\MinGW\msys\1.0\TPPC_FA.txt", True)
Set HdrFile = fso.CreateTextFile("C:\MinGW\msys\1.0\TPPC_FA_HDR.txt", True)

For i = 0 To rs.Fields.Count - 1
	If ((i = 3) Or (i = 11) Or (i = 17) Or (i = 18)) Then
		headLine = headLine & rs.Fields(i).Name & " "
	End If
	''MyFile.WriteLine(rs.Fields(i).Name)
	'If (i = OrderedArray(Counter)) Then
	'	headLine = headLine & rs.Fields(i).Name & "     "
	'End If
Next

Dim headSplit
headLine = headLine
'MsgBox(headLine)
headSplit = Split(headLine, " ")
'MsgBox(headSplit(2))
Dim strSplit
headLine = headSplit(2) & " " & headSplit(1) & " " & headSplit(3) & " " & headSplit(0)
HdrFile.WriteLine(headLine)

Do Until rs.EOF
	For i = 0 To rs.Fields.Count - 1
		If ((i = 3) Or (i = 11) Or (i = 17) Or (i = 18)) Then
			strResult = strResult & rs.Fields.Item(i) & " "
		End If
	Next
	strSplit = Split(strResult, " ")
	strResult = strSplit(2) & " " & strSplit(1) & " " & strSplit(3) & " " & strSplit(0)
	'Print MyFile, strResult
	MyFile.WriteLine(strResult)
	strResult = ""
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
'MyFile.WriteLine("This is a test.")
MyFile.Close
HdrFile.Close

