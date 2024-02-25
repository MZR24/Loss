Dim ws As Worksheet
Dim lastRow As Long
Dim col As Integer
Dim rng As Range

Set ws = ThisWorkbook.Sheets("Sheet1") ' Update with your sheet name
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Assuming column A has data

' Loop through each column starting from column B (assuming data starts from column B)
For col = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
    
    ' Convert text values in the range to numeric values
    rng.Value = Evaluate("IF(ISNUMBER(--" & rng.Address & "), --" & rng.Address & ", " & rng.Address & ")")
    
    ' Format the cells to display numeric values with two decimal places
    rng.NumberFormat = "0.00"
    
    ' Replace commas with periods for decimal values
    For Each cell In rng
        If InStr(cell.Value, ",") > 0 Then
            cell.Value = Replace(cell.Value, ",", ".")
        End If
    Next cell
Next col
