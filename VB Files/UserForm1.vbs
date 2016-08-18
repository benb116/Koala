Public labelcounter As Integer
Public opcounter As Integer

Private Sub AddButton_Click()
    Call AddLine(labelcounter)
    labelcounter = labelcounter + 1
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub ClearButton_Click()
    Sheets("Master").Range(Sheets("Master").Cells(2, opcounter + 3), Sheets("Master").Cells(400, opcounter + 3)).ClearContents
    Call reset
End Sub

Private Sub Image1_Click()

End Sub

Private Sub NextOpButton_Click()
    Dim cCont As Control
    Dim contCount As Long
    contCount = 1
    For Each cCont In Me.Controls
        If (TypeName(cCont) = "ComboBox" And cCont.Name <> "DayChoice") Then
            
            If cCont.Value <> vbNullString Then
                If IsNumeric(Me.Controls("qty" & contCount).Value) Then
                    Call AddToMaster(opcounter, cCont.Value, Me.Controls("qty" & contCount).Value)
                Else:
                    MsgBox "WTF, m8!!! That's supposed to be a number."
                    Exit Sub
                End If
            End If
            contCount = contCount + 1
        End If
    Next cCont
    
    opcounter = opcounter + 1
    Call reset
End Sub

Private Sub AddToMaster(opnum As Integer, pid As String, qty As Long)
    Dim row As Long
    Dim splitid As String
    splitid = Split(pid, " ")(0)
    row = Application.WorksheetFunction.Match(splitid, Sheets("Master").Range("A1:A400"), 0)
    Sheets("Master").Cells(row, opnum + 3).Value = qty

End Sub

Private Sub reset()
    GetCurrentWIP (opcounter)
    OpTitle.Caption = Worksheets("Operations").Cells((13 + opcounter), 2).Value
    For j = 1 To labelcounter - 1
        Me.Controls.Remove ("combo" & j)
        Me.Controls.Remove ("qty" & j)
    Next j
    Call AddLine(1)
    Call AddLine(2)
    Call AddLine(3)
    labelcounter = 4
End Sub

Private Sub PrevOpButton_Click()
    opcounter = opcounter - 1
    If opcounter < 2 Then opcounter = 2
    Call reset
End Sub

Private Sub UserForm_Initialize()
    Call Determine_Type
    OpTitle.Caption = Worksheets("Operations").Cells(15, 2).Value
    Dim theLabel As MSForms.Label
    labelcounter = 1
    
    opcounter = 2
    GetCurrentWIP (opcounter)
    Call AddLine(1)
    Call AddLine(2)
    Call AddLine(3)
    labelcounter = 4
End Sub
Private Sub AddLine(index As Integer)
    Dim cb As MSForms.ComboBox
    Dim qbox As MSForms.TextBox
    
    Set cb = Me.Controls.Add("Forms.ComboBox.1", "combo" & index)
    With cb
        .Left = 20
        .Width = 200
        .Top = 75 + 30 * index
    End With
    For Each Rng In Worksheets("Demand").Range("A2:A100")
        If Not IsEmpty(Rng) Then
            Dim thetype As Long
            thetype = Rng.Offset(0, 3).Value
            Dim thestring As String
            thestring = "D" & thetype & ":Z" & thetype
            Set therange = ThisWorkbook.Sheets("Operations").Range(thestring)

            For Each op In therange
                If (op.Value = opcounter) Then
                   cb.AddItem Rng.Value & " " & Rng.Offset(0, 1).Value
                   Exit For
                End If
            Next op
        End If
    Next Rng
    Set qbox = Me.Controls.Add("Forms.TextBox.1", "qty" & index)
    With qbox
        .Left = 225
        .Width = 50
        .Top = 75 + 30 * index
    End With
End Sub

Private Sub GetCurrentWIP(opindex As Long)
    Dim rownum As Long
    rownum = 2
    Dim wiptext As String
    wiptext = "Current WIP Totals"
    Do Until IsEmpty(Sheets("Master").Cells(rownum, 1))
        If Not IsEmpty(Sheets("Master").Cells(rownum, opindex + 3)) Then
            wiptext = wiptext & Chr(10) & Sheets("Master").Cells(rownum, 1).Value & " " & Sheets("Master").Cells(rownum, 2).Value & " - " & Sheets("Master").Cells(rownum, opindex + 3)
        End If
        rownum = rownum + 1
    Loop
    CurrentWIP.Caption = wiptext
End Sub
