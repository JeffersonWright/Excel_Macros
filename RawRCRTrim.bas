Attribute VB_Name = "RawRCRTrim"
Sub RawRCRTrim()
'
' RawRCRTrim Macro
'
On Error Resume Next
Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    Cells.ColumnWidth = 10
    Rows("1:1").Borders.LineStyle = xlContinuous
    Rows("1:1").Interior.ColorIndex = 37

            
    Range("F:F").Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Range("F:F").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    
    Cells(1, 6).Value = "File Number"
    Cells(1, 6).ColumnWidth = 16
    
Columns("N:O").EntireColumn.Delete

Columns("L").Cut
Columns("B").Insert Shift:=xlToRight
Columns("M").Cut
Columns("C").Insert Shift:=xlToRight
Columns("F").Cut
Columns("E").Insert Shift:=xlToRight
Columns("H").Cut
Columns("F").Insert Shift:=xlToRight
Columns("L").Cut
Columns("I").Insert Shift:=xlToRight
    
    Range("A1:M1").Select
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

Range("A1").Select
    
End Sub




