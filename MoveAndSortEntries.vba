Sub MoveAndSortEntries()
    Dim ws As Worksheet
    Dim builtPlanSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim entryName As String
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range
    Dim targetLastRow As Long
    Dim wsName As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set builtPlanSheet = ThisWorkbook.Sheets("Built plan")
    lastRow = builtPlanSheet.Cells(builtPlanSheet.Rows.Count, "H").End(xlUp).Row
    
    For i = 2 To lastRow
        entryName = builtPlanSheet.Cells(i, "H").Value
        
        If entryName <> "" Then
            On Error Resume Next
            Set targetSheet = ThisWorkbook.Sheets(entryName)
            On Error GoTo 0
            
            If targetSheet Is Nothing Then
                Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                targetSheet.Name = entryName
            End If
            
            targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, "H").End(xlUp).Row + 1
            builtPlanSheet.Rows(i).Copy Destination:=targetSheet.Rows(targetLastRow)
            builtPlanSheet.Rows(i).ClearContents
        End If
    Next i
    
    builtPlanSheet.UsedRange.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Built plan" Then
            lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            Set rng = ws.Range("A1:K" & lastRow)
            rng.Sort Key1:=ws.Range("K2"), Order1:=xlAscending, Header:=xlYes
        End If
    Next ws
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
