Attribute VB_Name = "BlanksPendingsFormat"
Sub BlanksPendingsFormat()

    Cells.Select
    Selection.Replace What:="NULL", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'Rows("1:1").Select
'Selection.AutoFilter
        
Rows(1).RowHeight = 45
Rows(1).WrapText = True

Rows("2:2").Select
ActiveWindow.FreezePanes = True


'Range("A1:P1").Borders.LineStyle = xlContinuous
'Range("A1:P1").Interior.ColorIndex = 37

Columns("A").ColumnWidth = 15
Columns("B").ColumnWidth = 10
Columns("C").ColumnWidth = 25
Columns("D").ColumnWidth = 15
Columns("E").ColumnWidth = 15
Columns("F").NumberFormat = "0"
Columns("F").ColumnWidth = 10
Columns("G").HorizontalAlignment = xlLeft
Columns("G").ColumnWidth = 15
Columns("H").NumberFormat = "m/d/yyyy"
Columns("H").ColumnWidth = 10
Columns("I").NumberFormat = "m/d/yyyy"
Columns("I").ColumnWidth = 10
Columns("J").ColumnWidth = 10
Columns("K").ColumnWidth = 10
Columns("L").ColumnWidth = 10
Columns("M").NumberFormat = "m/d/yyyy"
Columns("M").ColumnWidth = 10
Columns("N").ColumnWidth = 10
Columns("O").ColumnWidth = 10
Columns("P").NumberFormat = "m/d/yyyy"
Columns("P").ColumnWidth = 10

Application.DisplayAlerts = False
On Error Resume Next
Sheets("Sheet2").Delete
On Error Resume Next
Sheets("Sheet3").Delete
Application.DisplayAlerts = True

Range("A1").Select

End Sub




