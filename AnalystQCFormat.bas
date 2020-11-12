Attribute VB_Name = "AnalystQCFormat"
Sub AnalystQCFormat()

Application.StatusBar = "Running, please wait..."
Application.ScreenUpdating = False
    
ActiveSheet.Name = "Sheet1"
Rows(1).RowHeight = 45
Rows(1).WrapText = True
Range("D2").Select
ActiveWindow.FreezePanes = True
Columns("C").NumberFormat = "0"
Range("F1").Value = "On Previous QC 9/16/2019"
Columns("X").NumberFormat = "m/d/yyyy"
Columns("AP").NumberFormat = "m/d/yyyy"
Columns("AR").NumberFormat = "m/d/yyyy"
Columns("BF").NumberFormat = "m/d/yyyy"
Columns("BJ").NumberFormat = "m/d/yyyy"
Columns("BQ").NumberFormat = "m/d/yyyy"
Columns("BX").NumberFormat = "m/d/yyyy"
Columns("CB").NumberFormat = "m/d/yyyy"
Columns("CI").NumberFormat = "m/d/yyyy"

Range("=$A$1:$CI$1").FormatConditions.Add(xlUniqueValues).Interior.Color = RGB(79, 129, 189)
Range("=$AB:$AB,$BI:$BI").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($G1))").Interior.Color = vbYellow
Range("=$CA:$CA,$BI:$BI").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($H1))").Interior.Color = vbYellow
Range("=$BZ:$BZ,$BI:$BI").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($I1))").Interior.Color = vbYellow
Range("=$AC:$AC").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($P1))").Interior.Color = vbYellow
Range("=$AE:$AE").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($R1))").Interior.Color = vbYellow
Range("=$AG:$AG,$AI:$AI").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($S1))").Interior.Color = vbYellow
Range("=$BF:$BF").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($U1))").Interior.Color = vbYellow
Range("=$BJ:$BK").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($V1))").Interior.Color = vbYellow
Range("=$BI:$BI").FormatConditions.Add(xlExpression, xlFormula, "=NOT(ISBLANK($W1))").Interior.Color = vbYellow


Columns("A").ColumnWidth = 10
Columns("B").ColumnWidth = 25
Columns("C").ColumnWidth = 18
Columns("D").ColumnWidth = 15
Columns("E").ColumnWidth = 15
Columns("F").ColumnWidth = 11
Columns("G").ColumnWidth = 11
Columns("H").ColumnWidth = 11.56
Columns("I").ColumnWidth = 8
Columns("J").ColumnWidth = 11
Columns("K").ColumnWidth = 9
Columns("L").ColumnWidth = 12.22
Columns("M").ColumnWidth = 8
Columns("N").ColumnWidth = 10
Columns("O").ColumnWidth = 11.67
Columns("P").ColumnWidth = 8
Columns("Q").ColumnWidth = 10
Columns("R").ColumnWidth = 10
Columns("S").ColumnWidth = 9
Columns("T").ColumnWidth = 8
Columns("U").ColumnWidth = 9
Columns("V").ColumnWidth = 11
Columns("W").ColumnWidth = 12.56
Columns("X").ColumnWidth = 10
Columns("Y").ColumnWidth = 14
Columns("Z").ColumnWidth = 14
Columns("AA").ColumnWidth = 18
Columns("AB").ColumnWidth = 13.22
Columns("AC").ColumnWidth = 9
Columns("AD").ColumnWidth = 8
Columns("AE").ColumnWidth = 10
Columns("AF").ColumnWidth = 18
Columns("AG").ColumnWidth = 10
Columns("AH").ColumnWidth = 14
Columns("AI").ColumnWidth = 12
Columns("AJ").ColumnWidth = 18
Columns("AK").ColumnWidth = 10
Columns("AL").ColumnWidth = 10
Columns("AM").ColumnWidth = 10
Columns("AN").ColumnWidth = 13
Columns("AO").ColumnWidth = 14
Columns("AP").ColumnWidth = 10
Columns("AQ").ColumnWidth = 10
Columns("AR").ColumnWidth = 10
Columns("AS").ColumnWidth = 10
Columns("AT").ColumnWidth = 8
Columns("AU").ColumnWidth = 10
Columns("AV").ColumnWidth = 10
Columns("AW").ColumnWidth = 10
Columns("AX").ColumnWidth = 10
Columns("AY").ColumnWidth = 10
Columns("AZ").ColumnWidth = 8
Columns("BA").ColumnWidth = 8
Columns("BB").ColumnWidth = 14
Columns("BC").ColumnWidth = 14
Columns("BD").ColumnWidth = 14
Columns("BE").ColumnWidth = 10
Columns("BF").ColumnWidth = 10
Columns("BG").ColumnWidth = 10
Columns("BH").ColumnWidth = 10
Columns("BI").ColumnWidth = 10
Columns("BJ").ColumnWidth = 10
Columns("BK").ColumnWidth = 10
Columns("BL").ColumnWidth = 10
Columns("BM").ColumnWidth = 10
Columns("BN").ColumnWidth = 9
Columns("BO").ColumnWidth = 9
Columns("BP").ColumnWidth = 9
Columns("BQ").ColumnWidth = 10
Columns("BR").ColumnWidth = 10
Columns("BS").ColumnWidth = 10
Columns("BT").ColumnWidth = 10
Columns("BU").ColumnWidth = 10
Columns("BV").ColumnWidth = 8
Columns("BW").ColumnWidth = 10
Columns("BX").ColumnWidth = 10
Columns("BY").ColumnWidth = 8
Columns("BZ").ColumnWidth = 8
Columns("CA").ColumnWidth = 10
Columns("CB").ColumnWidth = 11
Columns("CC").ColumnWidth = 10
Columns("CD").ColumnWidth = 13.22
Columns("CE").ColumnWidth = 12
Columns("CF").ColumnWidth = 12
Columns("CG").ColumnWidth = 10
Columns("CH").ColumnWidth = 10
Columns("CI").ColumnWidth = 10


Application.DisplayAlerts = False
On Error Resume Next
Sheets("Sheet2").Delete
On Error Resume Next
Sheets("Sheet3").Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.StatusBar = False

Range("A1").Select

End Sub

