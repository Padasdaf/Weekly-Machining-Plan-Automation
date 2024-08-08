Sub SortSharedSheet()
    Dim builtPlanSheet As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set builtPlanSheet = ThisWorkbook.Sheets("Built plan")
    lastRow = builtPlanSheet.Cells(builtPlanSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        'Sort by Column B (A to Z) and then by Column F (Smallest to Largest)
        builtPlanSheet.Sort.SortFields.Clear
        builtPlanSheet.Sort.SortFields.Add Key:=builtPlanSheet.Range("B2:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        builtPlanSheet.Sort.SortFields.Add Key:=builtPlanSheet.Range("F2:F" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        With builtPlanSheet.Sort
            .SetRange builtPlanSheet.Range("A1:L" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

