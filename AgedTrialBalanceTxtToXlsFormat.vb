''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
'''
''' Peter Danshov 09.12.14
''' pdanshv@gmail.com
''' This macro formats Mainetti Aged
''' Trial Balance txt files.
'''
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''

Sub MainettiAgedTrialBalance_CSI()
    Dim LastRow As Integer
    Dim Row As Integer
    Dim BlankRow As Integer
    Dim LineItems As Integer
    Dim CustID As String
    Dim Co As String
    Dim Type_ As String
    Dim Num As String
    Dim Reg As String
    Dim Date_ As Date
    Dim Open_ As Integer
    Dim RowText As String
    Dim BlockRow As Integer
    Dim Block As Integer
    Dim CurrRow As Integer
    
    TrimExcssSpaces
    
    Row = 1
    
    Range("A1").Select
    
    Rows("1:9").Select
    Selection.Delete Shift:=xlUp
    
    Cells.Select
    Selection.NumberFormat = "@"
    Columns("F:F").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    Columns("G:G").Select
    Selection.NumberFormat = "$#,##0.00"
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    
    Cells(Row, 1) = "CustID"
    Cells(Row, 2) = "Co"
    Cells(Row, 3) = "Type"
    Cells(Row, 4) = "Number"
    Cells(Row, 5) = "Reference"
    Cells(Row, 6) = "Date"
    Cells(Row, 7) = "Open"
    
    Columns("G:G").ColumnWidth = 12.14
    Columns("F:F").ColumnWidth = 10.43
    Columns("E:E").ColumnWidth = 11.57
    Columns("D:D").ColumnWidth = 10.71
    Columns("A:A").ColumnWidth = 9.29
    
    LastRow = LastRowInOneColumn
    Range("A2").Select
    
    Do While Row <> LastRow + 1
        Row = ActiveCell.Row
        'Sub Sample()
        '    '~~> Check if Cell is empty and has spaces
        '    If Len(Trim(Range("A1").value)) = 0 And Len(Range("A1").value) <> 0 Then
        '        MsgBox "Cell A1 has spaces"
        '    '~~> Check is cell is empty and doesn't have anything
        '    ElseIf Len(Trim(Range("A1").value)) = 0 Then
        '        MsgBox "Cell A1 is empty"
        '    End If
        'End Sub
        If BlankRow = 5 Then
            Rows(Row & ":" & (Row + 8)).Delete
            LastRow = LastRow - 8
            BlankRow = 0
        '                Cells are empty                                        or      have only spaces
        'ElseIf Len(Trim(Range("A1").value)) = 0 And Len(Range("A1").value) <> 0 Or Len(Trim(Range("A1").value)) = 0 Then
        ElseIf ActiveCell.value = "" Then
            Cells(Row, "A").EntireRow.Delete
            LastRow = LastRow - 1
            BlankRow = BlankRow + 1
        ElseIf Left(Get_Word(Trim(ActiveCell.value), "First"), 1) = "-" Then
            Cells(Row, "A").EntireRow.Delete
            Cells(Row, "A").EntireRow.Delete
            LastRow = LastRow - 2
            BlankRow = 0
        Else
            CustID = Get_Word(Trim(ActiveCell.value), "First")
            Cells(Row, "A").EntireRow.Delete
            LastRow = LastRow - 1
            'ActiveCell.Offset.Select
            'Co = Get_Word(Trim(ActiveCell.value), "First")
            'Type_ = Get_Word(Trim(ActiveCell.value), 2)
            'Num = Get_Word(Trim(ActiveCell.value), 3)
            'Reg = Get_Word(Trim(ActiveCell.value), 4)
            'Date_ = Get_Word(Trim(ActiveCell.value), 5)
            'Open_ = Get_Word(Trim(ActiveCell.value), 7)
            '''''''''''''''''''''''''''''''''Count rows until next blank space on left
            CurrRow = ActiveCell.Row
            Block = NextBlank
            ActiveCell.Offset(-Block, 0).Select
            Do While BlockRow <> Block
                BlockRow = BlockRow + 1
                'Co = Get_Word(Trim(ActiveCell.value), "First")
                'Type_ = Get_Word(Trim(ActiveCell.value), 2)
                'Num = Get_Word(Trim(ActiveCell.value), 3)
                'Reg = Get_Word(Trim(ActiveCell.value), 4)
                'Date_ = Get_Word(Trim(ActiveCell.value), 5)
                'Open_ = Get_Word(Trim(ActiveCell.value), 7)
                RowText = ActiveCell.value
                ActiveCell.value = CustID
                Cells(Row, 2) = Get_Word(Trim(RowText), "First")
                Cells(Row, 3) = Get_Word(Trim(RowText), 2)
                Cells(Row, 4) = Get_Word(Trim(RowText), 3)
                Cells(Row, 5) = Get_Word(Trim(RowText), 4)
                Cells(Row, 6) = Get_Word(Trim(RowText), 5)
                Cells(Row, 7) = Get_Word(Trim(RowText), 7)
                ActiveCell.Offset(1, 0).Select
                Row = ActiveCell.Row
            Loop
            BlockRow = 0
            BlankRow = 0
        End If
        Row = ActiveCell.Row
    Loop
    Columns("G:G").Select
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("G5124").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-5122]C:R[-2]C)"
    Range("G5125").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(R[-5123]C:R[-3]C, "">0"")"
    Range("G5126").Select
End Sub

Public Function TrimExcssSpaces()
    Dim Row As Long
    Dim LastRow As Long
    Row = ActiveCell.Row
    LastRow = Cells.Find("*", [A1], , , xlByRows, xlPrevious).Row
    Do While Row <> (LastRow + 1)
        'CellVal = ActiveCell.value
        'TrimCell = Trim(CStr(CellVal))
        'ActiveCell.value = TrimCell
        ActiveCell.value = Application.WorksheetFunction.Substitute(Trim(ActiveCell.value), "     ", " ")
        ActiveCell.value = Application.WorksheetFunction.Substitute(Trim(ActiveCell.value), "    ", " ")
        ActiveCell.value = Application.WorksheetFunction.Substitute(Trim(ActiveCell.value), "   ", " ")
        ActiveCell.value = Application.WorksheetFunction.Substitute(Trim(ActiveCell.value), "  ", " ")
        ActiveCell.Offset(1, 0).Select
        Row = ActiveCell.Row
    Loop
End Function

Public Function NextBlank() As Integer
    Dim Next_ As Integer
    If ActiveCell.value = "" Then
        Exit Function
    End If
    Do While Left(Get_Word(Trim(ActiveCell.value), "First"), 1) <> "-"
        Next_ = Next_ + 1
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.value = "" Then
            Exit Do
        End If
    Loop
    NextBlank = Next_
End Function

Public Function LastRowInOneColumn() As Integer
'Find the last used row in a Column: column A in this example
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
'    MsgBox LastRow
    LastRowInOneColumn = LastRow
End Function

Public Function Get_Word(text_string As String, nth_word) As String
    'Option Compare Text
    Dim lWordCount As Long
    With Application.WorksheetFunction
        lWordCount = Len(text_string) - Len(.Substitute(text_string, " ", "")) + 1
        If IsNumeric(nth_word) Then
           nth_word = nth_word - 1
           Get_Word = Mid(Mid(Mid(.Substitute(text_string, " ", "^", nth_word), 1, 256), _
                .Find("^", .Substitute(text_string, " ", "^", nth_word)), 256), 2, _
                .Find(" ", Mid(Mid(.Substitute(text_string, " ", "^", nth_word), 1, 256), _
                .Find("^", .Substitute(text_string, " ", "^", nth_word)), 256)) - 2)
        ElseIf nth_word = "First" Then
            Get_Word = Left(text_string, .Find(" ", text_string) - 1)
        ElseIf nth_word = "Last" Then
            Get_Word = Mid(.Substitute(text_string, " ", "^", Len(text_string) - _
            Len(.Substitute(text_string, " ", ""))), .Find("^", .Substitute(text_string, " ", "^", _
            Len(text_string) - Len(.Substitute(text_string, " ", "")))) + 1, 256)
        End If
    End With
End Function

'Sample:

'  034201                                                 Mainetti Group LP                                   Page      -           1
' Mainetti USA INC                                 A/R Current Detail Aging - HSBC                            Date      -     7/22/14
'                                                                                                             As of     -    05/31/14
'
' Customer Number/Name                   Phone Number
'         . . Document Reference . .          . . . Balance . . .                    . . . . . . . . . Aging . . . . . . . .
'       Co    Ty  Number      Inv Date     Original             Open       Current          1 -  30       31 -  60       61 -  90
'------------------------------------- ----------------- ----------------- -------------- -------------- -------------- -------------
'
'  300012 A & H SPORTSWEAR               610    759-9550
'       00001 RI  1818519 001 03/10/14           272.45            272.45                                                     272.45
'       00001 RI  1818927 001 03/13/14           100.28            100.28                                                     100.28
'       00001 RI  1819450 001 03/20/14        28,760.00         28,760.00                                                   28760.00
'       00001 RI  1819633 001 03/21/14            40.00             40.00                                                      40.00
'       00001 RI  1820218 001 03/28/14            40.00             40.00                                                      40.00
'       00001 RI  1820474 001 04/01/14        26,240.00         26,240.00                                     26240.00
'       00001 RI  1823356 001 04/29/14           389.92            389.92                                       389.92
'       00001 RI  1823760 001 05/02/14            72.00             72.00                         72.00
'                                      ----------------- ----------------- -------------- -------------- -------------- -------------
'      300012 A & H SPORTSWEAR                55,914.65         55,914.65                         72.00       26629.92      29212.73
'
'
'    7456 A.H. SCHREIBER CO., INC.       212    564-2700
'       00001 RI  1817541 001 02/27/14         1,918.08          1,918.08                                                    1918.08
'       00001 RI  1817541 002 02/27/14            80.00             80.00                                                      80.00
'       00001 RI  1817831 001 03/03/14            62.44             62.44                                                      62.44
'                                      ----------------- ----------------- -------------- -------------- -------------- -------------
'        7456 A.H. SCHREIBER CO., INC          2,060.52          2,060.52                                                    2060.52
'
'

