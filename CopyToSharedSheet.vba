Sub CopyToSharedSheet()
    Dim builtPlanSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lastRow As Long
    Dim sourceLastRow As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set builtPlanSheet = ThisWorkbook.Sheets("Built plan")
    
    lastRow = builtPlanSheet.Cells(builtPlanSheet.Rows.Count, 1).End(xlUp).Row + 1
    
    For Each sourceSheet In ThisWorkbook.Worksheets
        If sourceSheet.Name <> "Built plan" Then
            sourceLastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
            
            If sourceLastRow > 1 Then
                For i = 2 To sourceLastRow
                    sourceSheet.Rows(i).Copy Destination:=builtPlanSheet.Rows(lastRow)
                    lastRow = lastRow + 1
                Next i
            End If
        End If
    Next sourceSheet
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
