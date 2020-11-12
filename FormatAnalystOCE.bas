Attribute VB_Name = "FormatAnalystOCE"
Sub FormatAnalystOCE()
Attribute FormatAnalystOCE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' OCE Trim for Analyst weekly report
'

Application.ScreenUpdating = False

'Delete all blank rows
On Error Resume Next
Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    
'Delete selected columns
    Columns("DV").EntireColumn.Delete 'Del Claim Paid Medical
    Columns("DR:DS").EntireColumn.Delete 'Del Denied Reason Code and Denied/Term Reason Description
    Columns("DK:DL").EntireColumn.Delete 'Del Examiner's Office and Supervisor
    Columns("DG:DI").EntireColumn.Delete 'Del Calendar Lost Time and Restricted
    Columns("CU:DE").EntireColumn.Delete 'Del Examiner Login through Level 1 Label
    Columns("CP:CS").EntireColumn.Delete 'Del Sharps Preventable to OSHA Level
    Columns("CB:CC").EntireColumn.Delete 'Del Sharps Procedure
    Columns("BW:BZ").EntireColumn.Delete 'Del Completed By through Level 5 Value
    Columns("BR:BS").EntireColumn.Delete 'Del Location Zip through Location County
    Columns("BN:BP").EntireColumn.Delete 'Del Location Address 1 through Location City
    Columns("BH:BL").EntireColumn.Delete 'Del Suppress reporting through SIC
    Columns("BC:BD").EntireColumn.Delete 'Del Occurrence Year and Case Source
    Columns("AY:BA").EntireColumn.Delete 'Del What was injury or illness through Last Updated By
    Columns("Z:AN").EntireColumn.Delete 'Del Current Work Status Effective Date through Employer Premises
    Columns("J:M").EntireColumn.Delete 'Del Privacy Case through Death Injury Related
    Columns("G").EntireColumn.Delete 'Del Middle Initial
    Columns("C").EntireColumn.Delete 'Del Vendor ID
    Columns("A").EntireColumn.Delete 'Del Case Number
    
    
'Format File Number in Column A before rearranging columns
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
    

'Rearrange Columns
    Dim arrColOrder As Variant, ndx As Integer
    Dim Found As Range, counter As Integer
    counter = 1
    
    'Place the column headers in the end result order you want.
    arrColOrder = Array("Client Number", "Client Name", "File Number", "Occurrence Date", "Employee Last Name", _
                        "Employee First Name", "Occupation", "Note Created Date", "Last Note Text", "Note Created By", _
                        "Create Timestamp", "Date Medical Note Created", "Medical Note", "Date MC Progress Report Note Created", "Managed Care Progress Report Note", _
                        "Claim Type", "Claim Status", "Claim Open Date", "OSHA Recordable Flag", "OSHA Work Related", _
                        "Classification of Case", "What happened?", "JURIS Cause Description", "JURIS Nature/Result Description", _
                        "JURIS Body Part/Target Description", "Cause Description", "Nature Description", "Body Part Desc", "Body Side", _
                        "Accident/Illness Type", "Accident Location", "Object", "What was emp doing before the incident?", "Recordable Days Away", _
                        "Recordable Days Restricted", "Actual Days Away", "Actual Days Restricted", "Work Status", "Benefit End Date", "Location Code", _
                        "Location Name", "Location State Code", "Physician Name", "Hospital Name", "Hospital Address1", "Hospital Address2", _
                        "Hospital City", "Hospital State", "Hospital Zip", "SHARPS Injury Case", "SHARPS How Incident Occurred", _
                        "SHARPS Body Part", "SHARPS Body Side", "Sharps Type", "SHARPS Type Other", _
                        "SHARPS Brand", "SHARPS Model", "SHARPS Work Area", "SHARPS Have Protection", _
                        "SHARPS Protection Involved", "SHARPS Exposure Occurred", "SHARPS Protection Activated", _
                        "SHARPS Level", "Examiner E-mail", "Supervisor Name", "Claim Substatus", "Denial Status Date", _
                        "Accepted Status Date", "Denied/Term Reason Description", "Location Effective Date", "Location Expiry Date")
    
    For ndx = LBound(arrColOrder) To UBound(arrColOrder)
    
        Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        
        If Not Found Is Nothing Then
            If Found.Column <> counter Then
                Found.EntireColumn.Cut
                Columns(counter).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
            counter = counter + 1
        End If
    Next ndx
    
 
'Format Alignment and Wrap
    
    'Format first row
    Rows("1:1").Borders.LineStyle = xlContinuous
    Rows("1:1").Interior.ColorIndex = 37
    
    'Format columns
    Range("A1:BR1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
    
    'Set row height
    Rows(1).RowHeight = 45
    Range("A2:A" & Rows.Count).Rows.RowHeight = 14.4
    
    'Apply filter
    Selection.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True


'Resize columns
    Cells.ColumnWidth = 16
    Columns("A").ColumnWidth = 9
    Columns("B").ColumnWidth = 24
    Columns("D").ColumnWidth = 11
    Columns("E").ColumnWidth = 13
    Columns("F").ColumnWidth = 13
    Columns("H").ColumnWidth = 11
    Columns("I").ColumnWidth = 35
    Columns("P").ColumnWidth = 11
    Columns("Q").ColumnWidth = 12
    Columns("R").ColumnWidth = 12
    Columns("S").ColumnWidth = 12
    Columns("T").ColumnWidth = 11
    Columns("U").ColumnWidth = 20
    Columns("AC").ColumnWidth = 11
    Columns("AD").ColumnWidth = 8
    Columns("AH:AK").ColumnWidth = 11
    Columns("AL").ColumnWidth = 12
    Columns("AP").ColumnWidth = 12
    Columns("AU:AV").ColumnWidth = 9
    Columns("AW").ColumnWidth = 12
    

'Reset screen for use
    Set Found = Rows("1:1").Find("", LookIn:=xlValues, LookAt:=xlPart)
    Application.ScreenUpdating = True
    Range("A1").Select

End Sub



