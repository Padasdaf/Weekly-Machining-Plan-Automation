Sub SortEntries()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Built plan" Then
            lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            If lastRow > 1 Then
                ws.Range("A1:L" & lastRow).Sort Key1:=ws.Range("K1"), Order1:=xlAscending, Header:=xlYes
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
