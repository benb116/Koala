Public hourspershift As Long
Sub Fill()
    Application.Calculation = xlCalculationAutomatic
    ' This is the main scheduling program
    ' Loop through each part number
    ' Repeat for each individual neck
    ' Go through each necessary operation

    'Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim index As Long
    index = 2
    Dim pid As String
    Dim qty As Long
    Dim thetype As Long
    Dim partindex As Long
    Dim theop As Long ' The operation number
    partindex = 2
    Dim partTime As Long ' How long has a part been worked on on a given day
    Dim dayofweek As Long
    dayofweek = 1
    
    Dim wipqty As Integer
    hourspershift = 10

    Call TimeCollection
            
    For f = 5 To 26
        Dim h As Long
        
        h = 31 - f
        'If Not (WorksheetFunction.CountA(Sheets("Master").Range(Cells(1, h), Cells(100, h)) = 0)) Then
            For g = 2 To 300
                If (Sheets("Master").Cells(g, h).Value <> vbNullString) Then
                    wipqty = Sheets("Master").Cells(g, h).Value
                    pid = Sheets("Master").Cells(g, 1).Value
                    thetype = Sheets("Master").Cells(g, 3).Value
                    
                    For neck = 1 To wipqty
                        opcounter = Sheets("Operations").Range(Sheets("Operations").Cells(thetype, 4), Sheets("Operations").Cells(thetype, 26)).Find((h - 3), LookIn:=xlValues).Column
                        dayofweek = 1
                        partTime = 0
                        Do Until IsEmpty(Sheets("Operations").Cells(thetype, opcounter)) ' For each operation
                            theop = Sheets("Operations").Cells(thetype, opcounter).Value
                            Call OpSched(pid, partindex, partTime, theop, dayofweek)
                            opcounter = opcounter + 1
                        Loop
                        partindex = partindex + 1
                    
                    Next neck
                    ' Subtract from the demand list
                    Dim demrow As Long
                    'demrow = Sheets("Demand").Range("A2:A100").Find(pid, LookIn:=xlValues).row
                    demrow = Application.Match(pid, Sheets("Demand").Range("A1:A100"), 0)
                    Sheets("Demand").Cells(demrow, 3).Value = Sheets("Demand").Cells(demrow, 3).Value - wipqty
                End If
            Next g
        'End If
    Next f
    index = 2
    Do Until IsEmpty(Sheets("Demand").Cells(index, 1)) ' For each part number
        If Not Sheets("Demand").Cells(index, 4).Value = 0 Then
            pid = Sheets("Demand").Cells(index, 1).Value
            qty = Sheets("Demand").Cells(index, 3).Value
            thetype = Sheets("Demand").Cells(index, 4).Value
            For i = 1 To qty ' For each individual neck
                dayofweek = 1
                opcounter = 4 ' Column of the first operations on the Operations list
                Dim optime As Long ' How long does the operation take per neck
                Dim cycletime As Long
                Dim macTime As Long ' How long has a machine been used on a given day
                Dim nextrow As Long
                partTime = 0
                Do Until IsEmpty(Sheets("Operations").Cells(thetype, opcounter)) ' For each operation
                    theop = Sheets("Operations").Cells(thetype, opcounter).Value
                    Call OpSched(pid, partindex, partTime, theop, dayofweek)
                    opcounter = opcounter + 1
                Loop
                partindex = partindex + 1
            Next i
        End If
        index = index + 1
        Application.ScreenUpdating = True
        Sheets("Demand").Cells(1, 5).Value = index
        Application.ScreenUpdating = False
    Loop
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    Sheets("Demand").Cells(2, 5) = dayofweek ' Total number of days that work is done. Used in the summarize call
End Sub
