Attribute VB_Name = "ExaminerQCFormat"
Sub ExaminerQCFormat()

Application.StatusBar = "Running, please wait..."
Application.ScreenUpdating = False
    
Range("I1").Value = "On Previous QC 9/4/2019"
Rows(1).RowHeight = 45
Rows(1).WrapText = True
Rows("2:2").Select
ActiveWindow.FreezePanes = True
Columns("C").NumberFormat = "0"
Columns("W").NumberFormat = "m/d/yyyy"
Columns("AI").NumberFormat = "m/d/yyyy"
Columns("AK").NumberFormat = "m/d/yyyy"
Columns("AM").NumberFormat = "m/d/yyyy"
Columns("BA").NumberFormat = "m/d/yyyy"
Columns("BG").NumberFormat = "m/d/yyyy"
Columns("BK").NumberFormat = "m/d/yyyy"

Range("=$A$1:$BK$1").FormatConditions.Add(xlUniqueValues).Interior.Color = RGB(79, 129, 189)


Range("=$Z:$Z").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($J1))").Interior.Color = vbYellow
Range("=$AI:$AJ").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($K1))").Interior.Color = vbYellow
Range("=$AB:$AB").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($L1))").Interior.Color = vbYellow
Range("=$AC:$AC").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($M1))").Interior.Color = vbYellow
Range("=$AB:$AB,$AE:$AE").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($N1))").Interior.Color = vbYellow
Range("=$AC:$AD,$AF:$AF").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($O1))").Interior.Color = vbYellow
Range("=$AP:$AQ").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($P1))").Interior.Color = vbYellow
Range("=$AN:$AN").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($Q1))").Interior.Color = vbYellow
Range("=$AO:$AO").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($R1))").Interior.Color = vbYellow
Range("=$AN:$AO").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($S1))").Interior.Color = vbYellow
Range("=$AL:$AL").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($T1))").Interior.Color = vbYellow
Range("=$AM:$AM").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($U1))").Interior.Color = vbYellow
Range("=$AK:$AK").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($V1))").Interior.Color = vbYellow

Columns("A").ColumnWidth = 10
Columns("B").ColumnWidth = 25
Columns("C").ColumnWidth = 18
Columns("D").ColumnWidth = 10
Columns("E").ColumnWidth = 11
Columns("F").ColumnWidth = 11
Columns("G").ColumnWidth = 12
Columns("H").ColumnWidth = 8
Columns("I").ColumnWidth = 8
Columns("J").ColumnWidth = 9.67
Columns("K").ColumnWidth = 8
Columns("L").ColumnWidth = 8
Columns("M").ColumnWidth = 8.33
Columns("N").ColumnWidth = 8.33
Columns("O").ColumnWidth = 8.33
Columns("P").ColumnWidth = 10.44
Columns("Q").ColumnWidth = 9.22
Columns("R").ColumnWidth = 8
Columns("S").ColumnWidth = 9.33
Columns("T").ColumnWidth = 8
Columns("U").ColumnWidth = 8
Columns("V").ColumnWidth = 8
Columns("W").ColumnWidth = 11
Columns("X").ColumnWidth = 12.56
Columns("Y").ColumnWidth = 10
Columns("Z").ColumnWidth = 14
Columns("AA").ColumnWidth = 10
Columns("AB").ColumnWidth = 9
Columns("AC").ColumnWidth = 8.56
Columns("AD").ColumnWidth = 9
Columns("AE").ColumnWidth = 8
Columns("AF").ColumnWidth = 8.33
Columns("AG").ColumnWidth = 9.44
Columns("AH").ColumnWidth = 9.56
Columns("AI").ColumnWidth = 11
Columns("AJ").ColumnWidth = 14
Columns("AK").ColumnWidth = 11
Columns("AL").ColumnWidth = 8
Columns("AM").ColumnWidth = 10
Columns("AN").ColumnWidth = 9.11
Columns("AO").ColumnWidth = 8
Columns("AP").ColumnWidth = 12
Columns("AQ").ColumnWidth = 10
Columns("AR").ColumnWidth = 10
Columns("AS").ColumnWidth = 10
Columns("AT").ColumnWidth = 10
Columns("AU").ColumnWidth = 8
Columns("AV").ColumnWidth = 9
Columns("AW").ColumnWidth = 15
Columns("AX").ColumnWidth = 10
Columns("AY").ColumnWidth = 10
Columns("AZ").ColumnWidth = 10
Columns("BA").ColumnWidth = 10
Columns("BB").ColumnWidth = 8
Columns("BC").ColumnWidth = 9.33
Columns("BD").ColumnWidth = 10
Columns("BE").ColumnWidth = 8
Columns("BF").ColumnWidth = 10
Columns("BG").ColumnWidth = 10
Columns("BH").ColumnWidth = 10
Columns("BI").ColumnWidth = 8
Columns("BJ").ColumnWidth = 9.22
Columns("BK").ColumnWidth = 10

Application.DisplayAlerts = False
On Error Resume Next
Sheets("Sheet2").Delete
On Error Resume Next
Sheets("Sheet3").Delete
Application.DisplayAlerts = True

    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table_Query_from_LTR1LEVSQL01" _
        ).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table_Query_from_LTR1LEVSQL01" _
        ).Sort.SortFields.Add Key:=Range( _
        "Table_Query_from_LTR1LEVSQL01[Examiner E-Mail]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects( _
        "Table_Query_from_LTR1LEVSQL01").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Application.ScreenUpdating = True
Application.StatusBar = False

Range("A1").Select

End Sub


