Attribute VB_Name = "RawQNRTrim"
Sub RawQNRTrim()
'
' RawQNRTrim Macro
'
On Error Resume Next
Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    Cells.ColumnWidth = 10
    Rows("1:1").Borders.LineStyle = xlContinuous
    Rows("1:1").Interior.ColorIndex = 37

    Range("A:A").Replace What:="Client ", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("C:C").Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Range("C:C").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    
    Cells(1, 3).Value = "File Number"
    Cells(1, 3).ColumnWidth = 16

    Columns("U").EntireColumn.Delete
    Columns("O:P").EntireColumn.Delete
    Columns("M").EntireColumn.Delete
    Columns("F:I").EntireColumn.Delete
    Columns("B").EntireColumn.Delete
        
    Range("A1:Q1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
    
    Selection.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
Range("A2:A" & Rows.Count).Rows.RowHeight = 14.4
    
Columns("C").ColumnWidth = 6

Columns("D").ColumnWidth = 5
Columns("E").ColumnWidth = 5

Cells(1, 16).Value = "Days from Date Open"
Range("P2").Formula = "=J2-G2"
Cells(1, 17).Value = "Days from Date Create"
Range("Q2").Formula = "=J2-O2"

Range("P2:Q2").Select
    
End Sub


