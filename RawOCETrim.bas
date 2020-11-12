Attribute VB_Name = "RawOCETrim"
Sub RawOCETrim()
Attribute RawOCETrim.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RawOCETrim Macro
'
On Error Resume Next
Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

ActiveSheet.Name = "Sheet1"
    Cells.ColumnWidth = 16
    Rows("1:1").Borders.LineStyle = xlContinuous
    Rows("1:1").Interior.ColorIndex = 37
    
    Columns("EE:EF").EntireColumn.Delete 'Del Location Effective Date through Location Expiry Date
    Columns("DS").EntireColumn.Delete 'Del Denied/Term Reason Description
    Columns("BW:DD").EntireColumn.Delete 'Del Completed By through Level 5 Value
    Columns("BR:BS").EntireColumn.Delete 'Del Location Zip through Location County
    Columns("BM:BP").EntireColumn.Delete 'Del Location Name through Location City
    Columns("BH:BJ").EntireColumn.Delete 'Del Supress reporting through Number of Employees
    Columns("BC:BD").EntireColumn.Delete 'Del Occurrence Year and Case Source
    Columns("AY").EntireColumn.Delete 'Del What was injury or illness
    Columns("AN").EntireColumn.Delete 'Del Employer Premises
    Columns("AJ:AK").EntireColumn.Delete 'Del ER Visit Flag and Hospital Flag
    Columns("Z:AF").EntireColumn.Delete 'Del Work Status Effective Date to EE Zip
    Columns("M").EntireColumn.Delete 'Del Death Injury Related
    Columns("K").EntireColumn.Delete 'Del Case Description
    Columns("G").EntireColumn.Delete 'Del Middle Initial
    Columns("C").EntireColumn.Delete 'Del Vendor ID
    Columns("A").EntireColumn.Delete 'Del Case Number
    
    
    Range("A:A").Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Range("A:A").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    
    Cells(1, 1).Value = "File Number"
    Cells(1, 1).ColumnWidth = 16
    
    Range("A1:BT1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
    
    Range("A2:A" & Rows.Count).Rows.RowHeight = 14.4
    
    Selection.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
Application.DisplayAlerts = False
On Error Resume Next
Sheets("Sheet2").Delete
On Error Resume Next
Sheets("Sheet3").Delete
Application.DisplayAlerts = True

Range("A1").Select
    
End Sub
