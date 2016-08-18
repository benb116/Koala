Sub Determine_Type()
    ' Uses the master type list to determine which set of operations applies to a part number
    ' Should probably have used VLOOKUP. Oh well
    Dim map As Collection
    Set map = New Collection
    
    Dim index As Integer
    index = 2
    Dim pid As String
    Dim thetype As Integer
    ' Build a collection with key:P/N as string and value:type as integer
    Do Until IsEmpty(Sheets("Master").Cells(index, 1))
        pid = Sheets("Master").Cells(index, 1).Value
        thetype = Sheets("Master").Cells(index, 3).Value
        map.Add thetype, pid
        index = index + 1
    Loop
    
    index = 2
    Dim qty As Long
    ' Assign types to the demand list
    Do Until IsEmpty(Sheets("Demand").Cells(index, 1))
        pid = Sheets("Demand").Cells(index, 1).Value
        qty = Sheets("Demand").Cells(index, 3).Value
        
        Dim booItemExistsInCollection As Boolean
        On Error Resume Next
        thetype = map.Item(pid)
        booItemExistsInCollection = (Err.Number = 0)
        If booItemExistsInCollection Then
            Sheets("Demand").Cells(index, 4).Value = thetype
            
        Else:
            Dim desc As String
            desc = Sheets("Demand").Cells(index, 2).Value
            Dim newadd As UserForm3
            Set newadd = New UserForm3
            
            newadd.NeckNum.Caption = pid
            newadd.NeckDesc.Caption = desc
            newadd.passIndex.Caption = index
            newadd.Show
        End If
        index = index + 1
    Loop
    ' Sort based on a custom list. List is based on which types reach certain operations first (and therefore should be scheduled earlier)
    'Sheets("Demand").Columns("A:D").Sort key1:=Range("D2"), order1:=xlAscending, Header:=xlYes, OrderCustom:=6, _
            MatchCase:=False, Orientation:=xlTopToBottom
End Sub
