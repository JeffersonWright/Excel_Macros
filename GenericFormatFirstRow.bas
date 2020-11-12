Attribute VB_Name = "GenericFormatFirstRow"
Sub GenericFormatFirstRow()

Application.StatusBar = "Running, please wait..."
Application.ScreenUpdating = False

Dim FileNumberColumn As Integer

'On Error Resume Next
'Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

On Error Resume Next
FileNumberColumn = Application.WorksheetFunction.Match("File Number", Rows(1), 0)
On Error Resume Next
FileNumberColumn = Application.WorksheetFunction.Match("File Number (unformatted)", Rows(1), 0)
        
Columns(FileNumberColumn).Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder _
:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

Columns(FileNumberColumn).Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder _
:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    Columns(FileNumberColumn).Select
    Selection.NumberFormat = "0"
        
    Cells(1, FileNumberColumn).Value = "File Number"
        
Rows(1).RowHeight = 45
Range("A2:A" & Rows.Count).Rows.RowHeight = 14.4
Rows(1).WrapText = True
Cells.ColumnWidth = 12

Cells(1, FileNumberColumn).ColumnWidth = 16


Rows("2:2").Select
ActiveWindow.FreezePanes = True

Rows("1:1").AutoFilter

Rows("1:1").Borders.LineStyle = xlContinuous
Rows("1:1").Interior.ColorIndex = 37

Range("A1").Select

Application.ScreenUpdating = True
Application.StatusBar = False

End Sub
