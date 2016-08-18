Public timemap As Collection
Public blankmap As Collection

Sub TimeCollection()
    ' Create a map with key: operation letter as string and value: time in seconds as integer
    Set timemap = New Collection
    
    Dim index As Integer
    index = 14
    Dim opID As String
    Dim duration As Long
    Do Until IsEmpty(Sheets("Operations").Cells(index, 1))
        opID = Sheets("Operations").Cells(index, 1).Value
        duration = Sheets("Operations").Cells(index, 5).Value
        timemap.Add duration, opID
        index = index + 1
    Loop
    
    Set blankmap = New Collection
    index = 2
    Dim neckid As String
    Dim blankid As String
    Do Until IsEmpty(Sheets("Neck Blanks").Cells(index, 1))
        neckid = Sheets("Neck Blanks").Cells(index, 1).Value
        blankid = Sheets("Neck Blanks").Cells(index, 2).Value
        blankmap.Add blankid, neckid
        index = index + 1
    Loop
End Sub
