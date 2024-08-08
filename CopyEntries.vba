Sub CopyEntries()
    Dim builtPlanSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim entryName As String
    Dim lastRow As Long
    Dim i As Long
    Dim targetLastRow As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set builtPlanSheet = ThisWorkbook.Sheets("Built plan")

    lastRow = builtPlanSheet.Cells(builtPlanSheet.Rows.Count, "H").End(xlUp).Row
    
    For i = 2 To lastRow
        Set targetSheet = ThisWorkbook.Sheets(builtPlanSheet.Cells(i, "H").Value)
        targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
        builtPlanSheet.Range("A" & i & ":L" & i).Copy Destination:=targetSheet.Range("A" & targetLastRow)
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
