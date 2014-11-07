
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''			Peter Danshov
'''''			11.06.14
'''''			pdanshv@gmail.com
'''''			This script reads .mdb
'''''			files and saves queried
'''''			rows to a txt file
'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

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

Dim fso, MyFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile("C:\MinGW\msys\1.0\TPPC_FA.txt", True)
For i = 0 To rs.Fields.Count - 1
	If ((i = 0) Or (i = 3) Or (i = 11) Or (i = 17)) Then
		headLine = headLine & rs.Fields(i).Name & "     "
	End If
	'MyFile.WriteLine(rs.Fields(i).Name)
Next
MyFile.WriteLine(headLine)
Do Until rs.EOF
	For i = 0 To rs.Fields.Count - 1
		If ((i = 0) Or (i = 3) Or (i = 11) Or (i = 17)) Then
			strResult = strResult & rs.Fields.Item(i) & "     "
		End If
	Next
	'Print MyFile, strResult
	MyFile.WriteLine(strResult)
	strResult = ""
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
'MyFile.WriteLine("This is a test.")
MyFile.Close

