Sub ClearAll()
    ' Reset the workbook to a blank template
    Sheets("Demand").Range("D2:R100").ClearContents
    Sheets("Week").Range("A2:AZ3000").ClearContents
    Sheets("Week").Cells(1, 25).ClearContents
    Dim iChar As Integer
    For iChar = 7 To 34
        Sheets((iChar)).Range("A3:EB4000").ClearContents
        Sheets((iChar)).Range("A1:EB1").ClearContents
        Sheets((iChar)).Range("F2:EB2").ClearContents
        For i = 1 To 5
            Sheets((iChar)).Cells(2, i).Value = "Part ID"
        Next i
    Next iChar
End Sub
