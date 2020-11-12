Attribute VB_Name = "RemoveRecommendations"
Sub RemoveRecommendation()

Dim NoteColumn As Integer
NoteColumn = Application.WorksheetFunction.Match("Last Note Text", Rows(1), 0)


'Remove all rows with a note that begins with Recommendation
        Dim DeletedRowCount As Long
        Dim i As Long
        LRow = Cells(Rows.Count, 1).End(xlUp).Row
    With Columns(NoteColumn)
        For i = LRow To 1 Step -1
            With .Cells(i, 1)
                If Not CStr(.Value) Like "*auto recorded*" Then If CStr(.Value) Like "Recommendation:*" Then .EntireRow.Delete: DeletedRowCount = DeletedRowCount + 1
            End With
        Next i
    End With


MsgBox "Number of cases trimmed with recommendation note: " & DeletedRowCount
End Sub
