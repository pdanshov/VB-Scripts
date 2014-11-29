
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''							Peter Danshov
''					11.28.14			pdanshv@gmail.com
''			This script Queries a SQL 2008 database and saves
''			the data in table-format in a txt file.
''
''		Logs:
''				C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\QUERY.LOG
''				C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\STATS.LOG
''
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

option explicit 'enforces variable names and declarations

Dim Connection
Dim Recordset
Dim SQL
Dim Server
'Dim field
Dim ImportId
Dim ImportBatchId
Dim DocType
Dim DocKey
Dim SenderID
Dim ReceiverID
Dim ISAControlNum
Dim GSControlNum
Dim STControlNum
Dim GSVersion
Dim StatusFlag
Dim SendDate
Dim SendTime
Dim IsNew
Dim DeletedYn
Dim DeleteUserId
Dim DeleteWrkStnId
Dim DeleteDate
Dim FAISAControlNum
Dim FAGSControlNum
Dim FASTControlNum
Dim FAReturnCode
Dim FADate
Dim FATime
Dim KeyLevelSeq1
Dim PartnerId
'Dim arrTemp : arrTemp = Array("something", "something else", "another thing", "this thing", "is it time to go home yet")
Dim arrTemp
Dim strTemp
Dim intSpace : intSpace = 0
Dim arrTemp2
Dim strTemp2
Dim intSpace2 : intSpace2 = 0

Dim objFSO
Dim outFile
Dim objFile

Set objFSO=CreateObject("Scripting.FileSystemObject")
' How to write file
outFile="C:\MinGW\msys\1.0\FA_Ack_Qry.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
arrTemp = Array("PartnerId","ISAControl","DocDesc","DocKey","FA_ISACntrl","	Send_Date","Send_Time","FA_Rtrn_Code")
'objFile.Write "PartnerId	ISAControl		DocDesc		DocKey		FA_ISACntrl		Send_Date		Send_Time		FA_Rtrn_Code" & vbCrLf
'objFile.Write "PartnerId" & vbTab & "ISAControl" & vbTab & "DocDesc" & vbTab & "DocKey" & vbTab & "FA_ISACntrl" & vbTab & "Send_Dat" & vbTab & "Send_Time" & vbTab & "FA_Rtrn_Code" & vbCrLf
For Each strTemp In arrTemp
  If Len(strTemp) > intSpace Then intSpace = Len(strTemp)
Next
' add some extra spaces
intSpace = intSpace + 5
' loop through strings putting in the necessary spaces
For Each strTemp In arrTemp
  'WScript.Echo strTemp & Space(intSpace - Len(strTemp)) & Now
  objFile.Write strTemp & Space(intSpace - Len(strTemp))
Next
objFile.Write vbCrLf
'objFile.Write "test string" & vbCrLf
'objFile.Close
'How to read a file
REM strFile = "c:\test\file"
REM Set objFile = objFS.OpenTextFile(strFile)
REM Do Until objFile.AtEndOfStream
    REM strLine= objFile.ReadLine
    REM Wscript.Echo strLine
REM Loop
REM objFile.Close
'to get file path without drive letter, assuming drive letters are c:, d:, etc
REM strFile="c:\test\file"
REM s = Split(strFile,":")
REM WScript.Echo s(1)



'declare the SQL statement that will query the database
'SQL = "SELECT ClusterAssignment FROM dbo.Studies WHERE Study_ID = '18054'"
SQL = "select * from [RAS].dbo.tblEdMailControl where (DocType='810' or DocType='856') and FAISAControlNum is null"

'create an instance of the ADO connection and recordset objects
Set Connection = CreateObject("ADODB.Connection")
Set Recordset = CreateObject("ADODB.Recordset")

'open the connection to the database
'Connection.Open "DSN=x;UID=u;PWD=p;Database=db"
'Setup DSN ODBC Datasource in Control Panel -> Administrative Tools
'	-> Data Sources (ODBC)
Connection.Open "DSN=Freelance;UID=Administrator;PWD=Freelance;Database=RAS"

'Open the recordset object executing the SQL statement and return records 
Recordset.Open SQL,Connection

'first of all determine whether there are any records 
If Recordset.EOF Then 
	wscript.echo "There are no records to retrieve; Check that you have the correct job number."
Else 
	'if there are records then loop through the fields 
	Do While NOT Recordset.Eof   
		'field = Recordset("ClusterAssignment")
		ImportId = RecordSet("ImportId")
		ImportBatchId = RecordSet("ImportBatchId")
		DocType = RecordSet("DocType")
		DocKey = RecordSet("DocKey")
		SenderID = RecordSet("SenderID")
		ReceiverID = RecordSet("ReceiverID")
		ISAControlNum = RecordSet("ISAControlNum")
		GSControlNum = RecordSet("GSControlNum")
		STControlNum = RecordSet("STControlNum")
		GSVersion = RecordSet("GSVersion")
		StatusFlag = RecordSet("StatusFlag")
		SendDate = RecordSet("SendDate")
		SendTime = RecordSet("SendTime")
		IsNew = RecordSet("IsNew")
		DeletedYn = RecordSet("DeletedYn")
		DeleteUserId = RecordSet("DeleteUserId")
		DeleteWrkStnId = RecordSet("DeleteWrkStnId")
		DeleteDate = RecordSet("DeleteDate")
		FAISAControlNum = RecordSet("FAISAControlNum")
		FAGSControlNum = RecordSet("FAGSControlNum")
		FASTControlNum = RecordSet("FASTControlNum")
		FAReturnCode = RecordSet("FAReturnCode")
		FADate = RecordSet("FADate")
		FATime = RecordSet("FATime")
		KeyLevelSeq1 = RecordSet("KeyLevelSeq1")
		PartnerId = RecordSet("PartnerId")
		'if field <> "" then
		if ImportId <> "" then
			'wscript.echo field
		'	wscript.echo "ImportId = " & ImportId
		'	wscript.echo "ImportBatchId = " & ImportBatchId
		'	wscript.echo "DocType = " & DocType
		'	wscript.echo DocKey
		'	wscript.echo SenderID
		'	wscript.echo ReceiverID
		'	wscript.echo ISAControlNum
		'	wscript.echo GSControlNum
		'	wscript.echo STControlNum
		'	wscript.echo GSVersion
		'	wscript.echo StatusFlag
		'	wscript.echo SendDate
		'	wscript.echo SendTime
		'	wscript.echo IsNew
		'	wscript.echo DeletedYn
		'	wscript.echo DeleteUserId
		'	wscript.echo DeleteWrkStnId
		'	wscript.echo DeleteDate
		'	wscript.echo FAISAControlNum
		'	wscript.echo FAGSControlNum
		'	wscript.echo FASTControlNum
		'	wscript.echo FAReturnCode
		'	wscript.echo FADate
		'	wscript.echo FATime
		'	wscript.echo KeyLevelSeq1
		'	wscript.echo PartnerId
			arrTemp2 = Array(PartnerId,ISAControlNum,DocType,DocKey,"null","null","null","null") 'FAISAControlNum,FADate,FATime,FAReturnCode = was null and crashed script
	'		objFile.Write PartnerId & vbTab & ISAControlNum & vbTab & DocType & vbTab & DocKey & vbTab & FAISAControlNum & vbTab & FADate & vbTab & FATime & vbTab & FAReturnCode & vbCrLf
			' loop through strings putting in the necessary spaces
			For Each strTemp2 In arrTemp2
			  'WScript.Echo strTemp & Space(intSpace - Len(strTemp)) & Now
			  objFile.Write strTemp2 & Space(intSpace - Len(strTemp2))
			Next
			'objFile.Write vbCrLf
		end if
		Recordset.MoveNext
	Loop
End If
objFile.Close
'close the connection and recordset objects to free up resources
Recordset.Close
Set Recordset=nothing
Connection.Close
Set Connection=nothing


