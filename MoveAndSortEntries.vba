Sub CopyEntries()
    Dim ws As Worksheet
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
    
    ' Change format of column H to General
    builtPlanSheet.Columns("H").NumberFormat = "General"
    
    ' Iterate through column H starting from row 2
    For i = 2 To lastRow
        entryName = builtPlanSheet.Cells(i, "H").Value
        
        ' Check if the entryName is not empty
        If entryName <> "" Then
            ' Check if the sheet exists, if not, create it
            On Error Resume Next
            Set targetSheet = ThisWorkbook.Sheets(entryName)
            On Error GoTo 0
            
            If targetSheet Is Nothing Then
                Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                targetSheet.Name = entryName
            End If
            
            ' Find the next empty row in the target sheet
            targetLastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Copy the entire row from the Built plan sheet to the target sheet
            builtPlanSheet.Rows(i).Copy Destination:=targetSheet.Rows(targetLastRow)
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
