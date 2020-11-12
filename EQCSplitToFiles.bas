Attribute VB_Name = "EQCSplitToFiles"
Public iMonth As String
Public iYear As String

Public Sub EQCSplitToFiles()

' MACRO SplitToFiles
' Last update: 2012-03-04
' Author: mtone
' Version 1.1
' Description:
' Loops through a specified column, and split each distinct values into a separate file by making a copy and deleting rows below and above
'
' Note: Values in the column should be unique or sorted.
'
' The following cells are ignored when delimiting sections:
' - blank cells, or containing spaces only
' - same value repeated
' - cells containing "total"
'
' Files are saved in a "Split" subfolder from the location of the source workbook, and named after the section name.

ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table_Query_from_LTR1LEVSQL01" _
        ).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table_Query_from_LTR1LEVSQL01" _
        ).Sort.SortFields.Add Key:=Range( _
        "Table_Query_from_LTR1LEVSQL01[[#All],[Examiner E-Mail]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects( _
        "Table_Query_from_LTR1LEVSQL01").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Dim osh As Worksheet ' Original sheet
Dim iRow As Long ' Cursors
Dim iCol As Long
Dim iFirstRow As Long ' Constant
Dim iTotalRows As Long ' Constant
Dim iStartRow As Long ' Section delimiters
Dim iStopRow As Long
Dim sSectionName As String ' Section name (and filename)
Dim rCell As Range ' current cell
Dim owb As Workbook ' Original workbook
Dim sFilePath As String ' Constant
Dim iCount As Integer ' # of documents created

    iMonth = MonthName(Month(Date))
    iYear = Year(Date)
    
iCol = 5 ' Application.InputBox("Enter the column number used for splitting", "Select column", 2, , , , , 1)
iRow = 2 ' Application.InputBox("Enter the starting row number (to skip header)", "Select row", 5, , , , , 1)
iFirstRow = iRow

Set osh = Application.ActiveSheet
Set owb = Application.ActiveWorkbook
iTotalRows = osh.UsedRange.Rows.Count
sFilePath = "J:\My Documents\QC\Examiner QCs" ' Application.ActiveWorkbook.Path


If Dir(sFilePath + "\Split", vbDirectory) = "" Then
    MkDir sFilePath + "\Split"
End If

'Turn Off Screen Updating  Events
Application.EnableEvents = False
Application.ScreenUpdating = False

Do
    ' Get cell at cursor
    Set rCell = osh.Cells(iRow, iCol)
    sCell = Replace(rCell.Text, " ", "")

    If sCell = "" Or (rCell.Text = sSectionName And iStartRow <> 0) Or InStr(1, rCell.Text, "total", vbTextCompare) <> 0 Then
        ' Skip condition met
    Else
        ' Found new section
        If iStartRow = 0 Then
            ' StartRow delimiter not set, meaning beginning a new section
            sSectionName = rCell.Text
            iStartRow = iRow
        Else
            ' StartRow delimiter set, meaning we reached the end of a section
            iStopRow = iRow - 1

            ' Pass variables to a separate sub to create and save the new worksheet
            CopySheet osh, iFirstRow, iStartRow, iStopRow, iTotalRows, sFilePath, sSectionName, owb.fileFormat
            iCount = iCount + 1

            ' Reset section delimiters
            iStartRow = 0
            iStopRow = 0

            ' Ready to continue loop
            iRow = iRow - 1
        End If
    End If

    ' Continue until last row is reached
    If iRow < iTotalRows Then
            iRow = iRow + 1
    Else
        ' Finished. Save the last section
        iStopRow = iRow
        CopySheet osh, iFirstRow, iStartRow, iStopRow, iTotalRows, sFilePath, sSectionName, owb.fileFormat
        iCount = iCount + 1

        ' Exit
        Exit Do
    End If
Loop

'Turn On Screen Updating  Events
Application.ScreenUpdating = True
Application.EnableEvents = True

MsgBox Str(iCount) + " documents saved in " + sFilePath


End Sub

Public Sub DeleteRows(targetSheet As Worksheet, RowFrom As Long, RowTo As Long)

Dim rngRange As Range
Set rngRange = Range(targetSheet.Cells(RowFrom, 1), targetSheet.Cells(RowTo, 1)).EntireRow
rngRange.Select
rngRange.Delete

End Sub


Public Sub CopySheet(osh As Worksheet, iFirstRow As Long, iStartRow As Long, iStopRow As Long, iTotalRows As Long, sFilePath As String, sSectionName As String, fileFormat As XlFileFormat)
     Dim ash As Worksheet ' Copied sheet
     Dim awb As Workbook ' New workbook
    
     ' Copy book
     osh.Copy
     Set ash = Application.ActiveSheet

     ' Delete Rows after section
     If iTotalRows > iStopRow Then
         DeleteRows ash, iStopRow + 1, iTotalRows
     End If

     ' Delete Rows before section
     If iStartRow > iFirstRow Then
         DeleteRows ash, iFirstRow, iStartRow - 1
     End If

     ' Select left-topmost cell
     ash.Cells(1, 1).Select

     ' Clean up a few characters to prevent invalid filename
     sSectionName = Replace(Left(sSectionName, InStr(1, sSectionName, "@")), "@", ".xlsx")

     ' Save in same format as original workbook
     ash.SaveAs "J:\My Documents\QC\Examiner QCs\Split\OSHA_QC_" + iMonth + iYear + sSectionName, fileFormat

     ' Close
     Set awb = ash.Parent
     awb.Close SaveChanges:=False
End Sub
