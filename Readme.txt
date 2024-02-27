Sub ApplyFormulaToCol58()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Find the last non-empty row in COL56
    lastRow = ws.Cells(ws.Rows.Count, 56).End(xlUp).Row
    
    ' Find the current row for inserting COL58 value
    currentRow = lastRow + 1
    
    ' Loop through the rows to find the matching values for COL44 and COL42 based on COL3
    For i = lastRow To 1 Step -1
        If ws.Cells(i, 3).Value = ws.Cells(currentRow, 3).Value And ws.Cells(i, 44).Value <> "" And ws.Cells(i, 42).Value <> "" Then
            ' Apply the formula to COL58 based on the found values
            ws.Cells(currentRow, 58).Formula = "=" & "R" & currentRow & "C56" & " - (R" & i & "C44 - R" & i & "C42)"
            Exit For
        End If
    Next i
    
End Sub
