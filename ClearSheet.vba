Sub ClearSheet()
    Dim builtPlanSheet As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set builtPlanSheet = ThisWorkbook.Sheets("Built plan")
    
    lastRow = builtPlanSheet.Cells(builtPlanSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        builtPlanSheet.Rows("2:" & lastRow).ClearContents
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
