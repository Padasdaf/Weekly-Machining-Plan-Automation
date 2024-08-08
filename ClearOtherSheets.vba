Sub ClearOtherSheets()
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Built plan" Then
            ws.UsedRange.Offset(1).ClearContents
        End If
    Next ws
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

