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


Subject: Inquiry Regarding Current Research Projects at eCUBATOR Innovation Unit

Dear [Recipient's Name],

I hope this message finds you well. I am writing to express my keen interest in the cutting-edge work being carried out by the eCUBATOR innovation unit at Knorr-Bremse. As a Master's student specializing in Vehicle Technology with a strong inclination towards electric propulsion systems, I am fascinated by the advancements in e-mobility and the innovative solutions being developed in this field.

Having had the privilege of interning within the department at Knorr-Bremse for the past three months, I have gained valuable insights into the company's commitment to excellence and innovation. The experience has further fueled my passion for electric propulsion technologies and their application in commercial vehicles.

As I approach the completion of my internship next month, I am eager to explore potential research topics related to the advanced innovations currently underway at the eCUBATOR unit. I believe that aligning my academic interests with the groundbreaking projects at eCUBATOR would not only enhance my learning experience but also contribute meaningfully to the ongoing developments in electric mobility.

I would greatly appreciate the opportunity to discuss any potential research collaborations or projects that I could contribute to within the eCUBATOR team. Your guidance and insights would be invaluable as I seek to deepen my knowledge and expertise in this dynamic field.

Thank you for considering my inquiry. I look forward to the possibility of exploring exciting research opportunities with the esteemed team at eCUBATOR.

Warm regards,

[Your Name]
Master's Student in Vehicle Technology
Knorr-Bremse Intern
