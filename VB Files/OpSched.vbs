Sub OpSched(pid As String, partindex As Long, partTime As Long, theop As Long, dayofweek As Long)
    optime = timemap.Item(CStr(theop))
    
    If (theop <> 24 And theop <> 28) Then ' Treat wait periods differently (no machine time)
        Dim realOp As Long
        realOp = theop
        macTime = 0
        If (theop = 26) Then
            realOp = 2
        ElseIf (theop = 27) Then
            realOp = 3
        End If
        macTime = Sheets(CStr(realOp)).Cells(1, dayofweek).Value ' How much time has this machine used already today
        If macTime < partTime Then ' If the machine is ready but the part is not, the machine waits for the part
            macTime = partTime
        Else: ' If the part is ready but the machine is not, the part waits for the machine
            partTime = macTime
        End If
        cycletime = optime
        Do While ((3600 * hourspershift - partTime) < cycletime) ' If the machine time is full or if the part has been worked on for too long
            dayofweek = dayofweek + 1 ' Go to the next day
            macTime = Sheets(CStr(realOp)).Cells(1, dayofweek).Value ' Get the next day's machine time
            partTime = macTime ' Reset the neck's daily part time
        Loop
        
        nextrow = 3
        nextrow = Sheets(CStr(realOp)).Cells(Sheets(CStr(realOp)).Rows.Count, dayofweek).End(xlUp).row + 1 ' Find the next empty cell on the op sheet in that day
        
        If (theop = 1) Then
            'MsgBox blankmap.Item(pID)
            Sheets(CStr(realOp)).Cells(nextrow, dayofweek).Value = blankmap.Item(pid) & "-" & pid
        Else:
            Sheets(CStr(realOp)).Cells(nextrow, dayofweek).Value = pid ' Record that the part is worked on
        End If
        partTime = partTime + optime
        macTime = macTime + optime ' Add the optime to the machine usage time and the part work time
        Sheets(CStr(realOp)).Cells(1, dayofweek).Value = macTime
        Sheets("Week").Cells(partindex, 1) = pid ' Add to the week summary
        Sheets("Week").Cells(partindex, (realOp + 1)).Value = dayofweek

    Else:
        ' Wait times don't have total machine usage times, also a part can wait for longer than the optime
        nextrow = 3
        nextrow = Sheets(CStr(theop)).Cells(Sheets(CStr(theop)).Rows.Count, dayofweek).End(xlUp).row + 1
        Sheets(CStr(theop)).Cells(nextrow, dayofweek).Value = pid
        partTime = partTime + optime
        Sheets("Week").Cells(partindex, (theop + 1)).Value = Sheets("Week").Cells(partindex, (theop + 1)).Value & ", " & dayofweek
        
        If (3600 * hourspershift < partTime) Then ' If the neck is still waiting after the shift ends, start working the next day
               dayofweek = dayofweek + 1
               partTime = 0
        End If
    End If
End Sub
