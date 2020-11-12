Attribute VB_Name = "QCforIOFormat"
Sub QCforIOFormat()

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
Columns("C").NumberFormat = "0"
Columns("C").HorizontalAlignment = xlLeft
Columns("K").NumberFormat = "m/d/yyyy"
Columns("L").NumberFormat = "m/d/yyyy"

'Range("A1:AD1").Borders.LineStyle = xlContinuous
'Range("A1:AD1").Interior.ColorIndex = 37


Columns("A").ColumnWidth = 10
Columns("B").ColumnWidth = 25
Columns("C").ColumnWidth = 18
Columns("D").ColumnWidth = 10
Columns("E").ColumnWidth = 10
Columns("F").ColumnWidth = 10
Columns("G").ColumnWidth = 10
Columns("H").ColumnWidth = 10
Columns("I").ColumnWidth = 10
Columns("J").ColumnWidth = 10
Columns("K").ColumnWidth = 12.22
Columns("L").ColumnWidth = 12.22
Columns("M").ColumnWidth = 10
Columns("N").ColumnWidth = 10
Columns("O").ColumnWidth = 10
Columns("P").ColumnWidth = 10
Columns("Q").ColumnWidth = 10
Columns("R").ColumnWidth = 10
Columns("S").ColumnWidth = 10
Columns("T").ColumnWidth = 10
Columns("U").ColumnWidth = 10
Columns("V").ColumnWidth = 10
Columns("W").ColumnWidth = 10
Columns("X").ColumnWidth = 10
Columns("Y").ColumnWidth = 10
Columns("Z").ColumnWidth = 10
Columns("AA").ColumnWidth = 10
Columns("AB").ColumnWidth = 10
Columns("AC").ColumnWidth = 10
Columns("AD").ColumnWidth = 10

Application.DisplayAlerts = False
On Error Resume Next
Sheets("Sheet2").Delete
On Error Resume Next
Sheets("Sheet3").Delete
Application.DisplayAlerts = True

Range("A1").Select

End Sub


