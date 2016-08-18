Public descmap As Collection

Sub AddDescriptions()

    Set descmap = New Collection
    
    Dim index As Long
    index = 2
    Dim pid As String
    Dim thedesc As String
    Do Until IsEmpty(ThisWorkbook.Sheets("Master").Cells(index, 1))
        pid = ThisWorkbook.Sheets("Master").Cells(index, 1).Value
        thedesc = ThisWorkbook.Sheets("Master").Cells(index, 2).Value
        descmap.Add thedesc, CStr(Val(pid))
        index = index + 1
    Loop
End Sub
