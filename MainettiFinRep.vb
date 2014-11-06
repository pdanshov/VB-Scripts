''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
'''
''' Peter Danshov 09.01.14
''' pdanshv@gmail.com
''' This macro formats Mainetti reports
''' generated in Traverse.
'''
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''

Sub Mainetti_Financial_Reports()

    Dim LastCol As Integer
    Dim Col As Integer
    Dim ColLtr As String
    Col = 1
    
    Columns("A:A").ColumnWidth = 28.43
    Columns("B:B").ColumnWidth = 36.43
    
    LastCol = LastColumnInOneRow
    
    Do While Col <> LastCol + 1
        ActiveSheet.Cells(6, Col).Select
        ColLtr = Chr(Col + 64)
        If IsNumeric(ActiveCell.value) Then
            Columns(ColLtr & ":" & ColLtr).Select
            Selection.TextToColumns Destination:=Range(ColLtr & "1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        End If
        Col = Col + 1
    Loop
    
End Sub

    Sub LastRowInOneColumn()
    'Find the last used row in a Column: column A in this example
        Dim LastRow As Long
        With ActiveSheet
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
    '    MsgBox LastRow
    End Sub
    
    Public Function LastColumnInOneRow() As Integer
    'Find the last used column in a Row: row 6 in this example
        Dim LastCol As Integer
        With ActiveSheet
            LastCol = .Cells(6, .Columns.Count).End(xlToLeft).Column
        End With
    '    MsgBox LastCol
    'Proper way to return value,
    'instead of setting return = LastCol
    'set nameoffuntion = LastCol
        LastColumnInOneRow = LastCol
    End Function
