Sub Summarize()
    ' Creates a new workbook with summaries and nice formatting
    Application.ScreenUpdating = False
    'Application.PrintCommunication = False
    Call AddDescriptions
    Application.SheetsInNewWorkbook = 1
    Set NewBook = Workbooks.Add
    With NewBook
        .Title = "Neck Mill Schedule"
    End With
    ' Note: ActiveSheet and ActiveWorkbook are the currently worked on sheet and workbook, respectively
    ' Note: ThisWorkbook is the workbok from which the macro is running (Master Neck Table.xlsm)
    ' Set up the week summary
    ActiveSheet.Name = "Week Summary"
    ActiveSheet.Cells(1, 1).Value = "Operation"
    ActiveSheet.Range("A:A").Font.Bold = True
    ActiveSheet.Range("A1:AZ1").Font.Bold = True
    ActiveSheet.Range("A2:AZ2").Font.Bold = True
    ActiveSheet.Range("A1:AZ1").HorizontalAlignment = xlHAlignCenter
    
    ' In the main workbook's Operations sheet, what is the first operation's row index
    Dim opcounter As Long
    Dim ws As Worksheet
    opcounter = 15 ' If we add more types, change this number!
    Dim opname As String
    
    Dim numofdays As Long
    numofdays = ThisWorkbook.Sheets("Demand").Cells(2, 5).Value ' Get the number of days in the workweek from the main workbook
    
    ' Add day headers to the week summary sheet
    For h = 1 To numofdays
        ActiveSheet.Range(ActiveSheet.Cells(1, (3 * h - 1)), ActiveSheet.Cells(1, (3 * h + 1))).Merge
    Next h
    
    Dim daymap As Collection
    Set daymap = New Collection
    daymap.Add ("Monday")
    daymap.Add ("Tuesday")
    daymap.Add ("Wednesday")
    daymap.Add ("Thursday")
    daymap.Add ("Friday")
    
    Dim weekloops As Long
    weekloops = numofdays / 5 ' How many full weeks does it take?
    
    ' Add day labels
    Dim daylabel As Long
    daylabel = 1
    For daylabel = 1 To numofdays
        Dim thisday As Integer
        thisday = daylabel
        If thisday > 5 Then
            thisday = thisday - 5
        End If
        ActiveSheet.Cells(1, (3 * daylabel - 1)).Value = daymap.Item(thisday)
        ActiveSheet.Cells(2, (3 * daylabel - 1)).Value = "Part ID"
        ActiveSheet.Cells(2, (3 * daylabel)).Value = "Description"
        ActiveSheet.Cells(2, (3 * daylabel + 1)).Value = "Qty"
    Next daylabel
    
    Dim daycounter As Long
    Dim opletter As String
    Dim neckcounter As Long
    Dim thisneck As String
    Dim oldneck As String
    Dim machinerow As Long
    Dim partrow As Long
    
    'For each operation, summarize the week's work
    Do Until IsEmpty(ThisWorkbook.Sheets("Operations").Cells(opcounter, 2))
        
        opname = ThisWorkbook.Sheets("Operations").Cells(opcounter, 2).Value
        ' Find the next totally clear row to start a new machine summary on the week page
        machinerow = 3
        Do Until WorksheetFunction.CountA(Sheets(1).Range("A" & machinerow & ":AZ" & machinerow)) = 0
            machinerow = machinerow + 1
        Loop
        Sheets(1).Cells(machinerow, 1).Value = opname
        
        ' Make an operation-specific sheet
        Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        ws.Name = opname
        ' Add titles
        ws.Range("A1:C1").Merge
        ws.Cells(1, 1).Value = opname
        opletter = ThisWorkbook.Sheets("Operations").Cells(opcounter, 1)
        daycounter = 1
        For daycounter = 1 To numofdays
            thisday = daycounter
            If thisday > 5 Then
                thisday = thisday - 5
            End If
            ws.Cells(2, (3 * daycounter - 2)).Value = daymap.Item(thisday)
            ws.Cells(3, (3 * daycounter - 2)).Value = "Part ID"
            ws.Cells(3, (3 * daycounter - 1)).Value = "Description"
            ws.Cells(3, (3 * daycounter)).Value = "Qty"
        Next daycounter
        
        
        ws.Range("A1:AZ3").Font.Bold = True
        
        ' Find the next clear row for this operation on the week page
        partrow = machinerow
        Do Until WorksheetFunction.CountA(Sheets(1).Range("B" & partrow & ":W" & partrow)) = 0
            partrow = partrow + 1
        Loop
        
        Dim therange As Range
        Dim newrange As Range
        ' Summarize the week
        For l = 1 To numofdays
            Dim startchar As String
            Dim endchar As String
            startchar = Chr(64 + l)
            endchar = Chr(64 + (3 * l - 2))
            Set therange = ThisWorkbook.Sheets(opletter).Range(startchar & "3:" & startchar & "3000") ' The part numbers worked on in one day (in main workbook)
            ' If the operation is not empty
            If Not (WorksheetFunction.CountA(therange) = 0) Then
                therange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range(ConvertToLetter(3 * l - 2) & "4"), Unique:=True ' Copy unique part numbers to the new op sheet

                Dim uniqcount As Long
                uniqcount = 4
                Dim qty As Long
                Dim thedesc As String
                ' Summarize each part number
                Do Until IsEmpty(ws.Cells(uniqcount, (3 * l - 2)))
                    qty = Application.WorksheetFunction.CountIf(therange, ws.Cells(uniqcount, (3 * l - 2))) ' Find quantity from the main workbook
                    pid = CStr(ws.Cells(uniqcount, (3 * l - 2)).Value)
                    thedesc = descmap.Item(pid) ' Get the P/N's description
                    ws.Cells(uniqcount, (3 * l - 1)).Value = thedesc
                    ws.Cells(uniqcount, (3 * l)).Value = qty
                    uniqcount = uniqcount + 1
                Loop
            End If
            ' Formatting
            ws.Columns(3 * l).Cells.HorizontalAlignment = xlHAlignCenter
            'MsgBox ws.UsedRange.Rows.Count
            
            endchar = Chr(65 + 3 * numofdays)
            With ws.Range(ws.Cells(1, 1), ws.Cells(3, (3 * numofdays))).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
        
        Next l
        For k = 1 To numofdays
            With ws.Range(Cells(1, (3 * k - 2)), Cells(ws.UsedRange.Rows.Count, (3 * k - 2))).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
        Next k
        endchar = Chr(65 + 3 * numofdays)
        ws.Rows(4).Delete ' Not sure why there is an extra row created. Delete it
        ' More formatting
        With ws.Range(ws.Cells(4, 1), ws.Cells(ws.UsedRange.Rows.Count, (3 * numofdays))).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        ws.Range(ws.Cells(4, 1), ws.Cells(ws.UsedRange.Rows.Count, (3 * numofdays))).Copy Sheets(1).Range("B" & partrow) ' Copy this info to the week summary page
        
        With Sheets(1).Range(Sheets(1).Cells(partrow, 1), Sheets(1).Cells(partrow, (3 * numofdays))).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        ws.Range("A:AZ").EntireColumn.AutoFit
        With ws.PageSetup
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
        opcounter = opcounter + 1
    Loop
    ActiveWorkbook.Sheets(1).Range("A:AZ").EntireColumn.AutoFit
    'Application.PrintCommunication = True
    Application.ScreenUpdating = True
    ActiveWorkbook.Sheets(1).Activate
End Sub

Function ConvertToLetter(iCol As Long) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function
