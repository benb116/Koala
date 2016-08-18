Sub Efficiency()
    Call TimeCollection
    Dim sum As Long
    sum = 0
    For i = 5 To 24
        sum = sum + Application.sum(ThisWorkbook.Sheets(i).Range("A1:H1"))
    Next i
    MsgBox sum
End Sub
