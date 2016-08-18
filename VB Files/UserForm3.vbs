Public TypesList As Collection
Private num As String
Private desc As String

Private Sub AddButton_Click()
    Dim booItemExistsInCollection As Boolean
    Dim thetype As Long
    On Error Resume Next
    thetype = TypesList.Item(TypeList.Value)
    booItemExistsInCollection = (Err.Number = 0)
    If booItemExistsInCollection Then
        Dim nextrow As Long
        nextrow = Sheets("Master").Cells(Sheets("Master").Rows.Count, 1).End(xlUp).row + 1
        Sheets("Master").Cells(nextrow, 1).Value = NeckNum.Caption
        Sheets("Master").Cells(nextrow, 2).Value = NeckDesc.Caption
        Sheets("Master").Cells(nextrow, 3).Value = thetype
        Sheets("Demand").Cells(passIndex.Caption, 4).Value = thetype
        MsgBox NeckNum.Caption & " " & NeckDesc.Caption & " has been added to the master list."
        Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    'NeckNum.Caption = num
    'NeckDesc.Caption = desc
    Dim index As Long
    index = 1
    Set TypesList = New Collection
    
    Do Until IsEmpty(Sheets("Operations").Cells(index, 2))
        TypeList.AddItem Sheets("Operations").Cells(index, 2).Value
        TypesList.Add index, Sheets("Operations").Cells(index, 2).Value
        index = index + 1
    Loop
    
End Sub
