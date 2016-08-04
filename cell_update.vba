Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim rCell As Range
    Dim rChange As Range
    
    On Error GoTo ErrHandler
    Set rChange = Intersect(Target, Range("B:C"))
    If Not rChange Is Nothing Then
        Application.EnableEvents = False
        For Each rCell In rChange
            If rCell > "" Then
                With Cells(rCell.Row, 8)
                    .Value = Now
                    .NumberFormat = "dddd, dd/mm/yy h:mm AM/PM"
                End With
         'Else
          'rCell.Offset(0, 1).Clear
            End If
        Next
    End If

ExitHandler:
    Set rCell = Nothing
    Set rChange = Nothing
    Application.EnableEvents = True
    Exit Sub
ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
End Sub
