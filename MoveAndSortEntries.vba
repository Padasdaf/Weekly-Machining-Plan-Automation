Sub MoveAndSortEntries()
    Dim ws As Worksheet
    Dim builtPlanSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim entryName As String
    Dim lastRow As Long
    Dim i As Long
    Dim targetLastRow As Long
    Dim rowCount As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set builtPlanSheet = ThisWorkbook.Sheets("Built plan")
    lastRow = builtPlanSheet.Cells(builtPlanSheet.Rows.Count, "H").End(xlUp).Row
    
    For i = 2 To lastRow
        entryName = builtPlanSheet.Cells(i, "H").Value
        
        If entryName <> "" Then
            ' Check if the sheet exists, if not, create it
            On Error Resume Next
            Set targetSheet = ThisWorkbook.Sheets(entryName)
            On Error GoTo 0
            
            If targetSheet Is Nothing Then
                Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                targetSheet.Name = entryName
            End If
            
            targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
            builtPlanSheet.Rows(i).Copy Destination:=targetSheet.Rows(targetLastRow)
        End If
    Next i
    
    For i = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(builtPlanSheet.Rows(i)) = 0 Then
            builtPlanSheet.Rows(i).Delete
        End If
    Next i
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Built plan" Then
            lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            If lastRow > 1 Then
                ws.Range("A1:K" & lastRow).Sort Key1:=ws.Range("K1"), Order1:=xlAscending, Header:=xlYes
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
