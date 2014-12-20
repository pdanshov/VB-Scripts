
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''		Peter Danshov - Consolidate All Worksheets to first worksheet
''		12.19.14 - 1207
''
''		Can use built-in functon: Data -> Consolidate
''		Or this macro.
''
''		http://excel.tips.net/T003005_Condensing_Multiple_Worksheets_Into_One.html
''
''		http://excelribbon.tips.net/T008884_Condensing_Multiple_Worksheets_Into_One.htmls
''
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Microsoft Excel versions: 97, 2000, 2002, and 2003
Sub Combine()
    Dim J As Integer

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    ' work through sheets
    For J = 2 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(J).Activate ' make the sheet active
        Range("A1").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets

        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select

        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
End Sub

' Microsoft Excel versions: 2007, 2010, and 2013
Sub Combine()
    Dim J As Integer
    Dim s As Worksheet

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    For Each s In ActiveWorkbook.Sheets
        If s.Name <> "Combined" Then
            Application.GoTo Sheets(s.Name).[a1]
            Selection.CurrentRegion.Select
            ' Don't copy the headings
            Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
            Selection.Copy Destination:=Sheets("Combined"). _
              Cells(Rows.Count, 1).End(xlUp)(2)
        End If
    Next
End Sub

' http://stackoverflow.com/questions/20903181/excel-vba-copying-multiple-sheets-into-new-workbook
REM Try do something like this (the problem was that you trying to use MyBook.Worksheets, but MyBook is not a Workbook object, but string, containing workbook name. I've added new varible Set WB = ActiveWorkbook, so you can use WB.Worksheets instead MyBook.Worksheets):
REM Sub NewWBandPasteSpecialALLSheets()
   REM MyBook = ActiveWorkbook.Name ' Get name of this book
   REM Workbooks.Add ' Open a new workbook
   REM NewBook = ActiveWorkbook.Name ' Save name of new book
   REM Workbooks(MyBook).Activate ' Back to original book
   REM Set WB = ActiveWorkbook
   REM Dim SH As Worksheet
   REM For Each SH In WB.Worksheets
       REM SH.Range("WholePrintArea").Copy
       REM Workbooks(NewBook).Activate
       REM With SH.Range("A1")
        REM .PasteSpecial (xlPasteColumnWidths)
        REM .PasteSpecial (xlFormats)
        REM .PasteSpecial (xlValues)
       REM End With
     REM Next
REM End Sub
REM But your code doesn't do what you want: it doesen't copy something to a new WB. So, the code below do it for you:
REM Sub NewWBandPasteSpecialALLSheets()
   REM Dim wb As Workbook
   REM Dim wbNew As Workbook
   REM Dim sh As Worksheet
   REM Dim shNew As Worksheet
   REM Set wb = ThisWorkbook
   REM Workbooks.Add ' Open a new workbook
   REM Set wbNew = ActiveWorkbook
   REM On Error Resume Next
   REM For Each sh In wb.Worksheets
      REM sh.Range("WholePrintArea").Copy
      REM 'add new sheet into new workbook with the same name
      REM With wbNew.Worksheets
          REM Set shNew = Nothing
          REM Set shNew = .Item(sh.Name)
          REM If shNew Is Nothing Then
              REM .Add After:=.Item(.Count)
              REM .Item(.Count).Name = sh.Name
              REM Set shNew = .Item(.Count)
          REM End If
      REM End With
      REM With shNew.Range("A1")
          REM .PasteSpecial (xlPasteColumnWidths)
          REM .PasteSpecial (xlFormats)
          REM .PasteSpecial (xlValues)
      REM End With
   REM Next
REM End Sub
REM share|improve this answer
REM edited Jan 3 at 16:30
REM answered Jan 3 at 12:15
REM simoco
REM 23.9k92339
REM Many thanks - but I have a problem that this only works for the 1st sheet and then produces "Subscrpt out of range" error message. Also there is definitely a problem with my range called "WholePrintArea" as some sheets do have Print_Area which differ, so I have tried inserting here:. –  user3157086 Jan 3 at 16:15
REM So, I guess, an error occurs because the line sh.Range("WholePrintArea").Copy. First sheet has range WholePrintArea, but the second sheet doesn't. Tell me please, what is main idea of your code? I mean, what are you expecting your code should do for you? –  simoco Jan 3 at 16:21
REM Many thanks - but I have a problem that this works only if I use "sh.copy" ie without a range, and even then only for the 1st sheet producing a "Subscrpt out of range" error message. Also there is definitely a problem with my range called "WholePrintArea" as some sheets do have Print Areas which differ, so I have tried inserting "sh.Range(Print_Area).Copy" instead but this produces a 400 code implying the range name does not exist even though it does. –  user3157086 Jan 3 at 16:22
REM I've updated my answer, so it fixes the issue with "Subscrpt out of range" error message –  simoco Jan 3 at 16:30
REM Code is to create a new WB, copying varying print areas from each of the source sheets into the new WB and using the same sheet names. The source WB has data outside of the print areas (set using File > Print Area > Set print area) which is not to be copied. If required I could name each sheet and each print area, but there are 19 of them. –  user3157086 Jan 3 at 16:34
REM In the loop For Each sh In wb.Worksheets you can add something like this sh.PageSetup.PrintArea = "$A$1:$X100", so it will add "Print_Area" range to your sheet. And then you can use sh.Range("Print_Area").Copy –  simoco Jan 3 at 16:39


' http://stackoverflow.com/questions/18854432/copy-all-worksheets-into-one-sheet
REM Sub tgr()
    REM Dim ws As Worksheet
    REM Dim wsDest As Worksheet
    REM Set wsDest = Sheets("Results")
    REM For Each ws In ActiveWorkbook.Sheets
        REM If ws.Name <> wsDest.Name Then
            REM ws.Range("A2", ws.Range("A2").End(xlToRight).End(xlDown)).Copy
            REM wsDest.Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValues
        REM End If
    REM Next ws
REM End Sub

' http://www.rondebruin.nl/win/s3/win002.htm


