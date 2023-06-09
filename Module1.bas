Option Explicit

Dim dataSheet As Worksheet
Dim todaySheet As Worksheet
Dim weeklySheet As Worksheet

Function GetRowWithValueInColumn(ws As Worksheet, col As Integer, val As String)
    AssignWorksheetGlobals
    
    Dim totalRows As Long
    totalRows = NumOfRows(ws)
    
    Dim i As Long
    For i = 1 To totalRows
        If ws.Cells(i, col).Value = val Then
            GetRowWithValueInColumn = i
            Exit Function
        End If
    Next i
    GetRowWithValueInColumn = CLng(-1)
End Function

Function GetTodayRowWithID(ID As String) As Long
    AssignWorksheetGlobals
    
    GetTodayRowWithID = GetRowWithValueInColumn(todaySheet, 5, ID)
End Function

Function GetDataRowWithID(ID As String) As Long
    AssignWorksheetGlobals
    
    GetDataRowWithID = GetRowWithValueInColumn(dataSheet, 5, ID)
End Function

Sub guaranteeSheet(sheetName As String)
    If sheetExists(sheetName) = False Then
        ThisWorkbook.Sheets.Add.Name = sheetName
    End If
End Sub

Function sheetExists(sheetToFind As String) As Boolean
    Dim Sheet As Object
    For Each Sheet In ThisWorkbook.Sheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
    sheetExists = False
End Function

Sub AssignWorksheetGlobals()
    Set dataSheet = Worksheets("data")
    Set todaySheet = Worksheets("Today")
    Set weeklySheet = Worksheets("Weekly")
End Sub

Function NumOfRows(ws As Worksheet) As Long
    Dim totalRow As Long
    
    With ws
        totalRow = .Range("B" & .Rows.Count).End(xlUp).Row
    End With
    
    NumOfRows = totalRow
End Function

Sub ClearTodaySheet()
    todaySheet.Cells.Clear
    
    Dim cb As CheckBox
    For Each cb In todaySheet.CheckBoxes
        cb.Delete
    Next
End Sub

Sub TransferTasksDueToday()
    Application.EnableEvents = False
    On Error GoTo ERR_HANDLE
    
    AssignWorksheetGlobals
    
    ClearTodaySheet
    
    Dim totalTasks As Long
    totalTasks = NumOfRows(dataSheet)
    
    Dim i As Long
    Dim targetRow As Long
    targetRow = 2
    
    For i = 1 To totalTasks
        
        With dataSheet
            If .Range("A" & i).Value = Date Then
                CopyRowToToday i, CLng(targetRow)
                
                AddCheckbox todaySheet.Range("A" & targetRow), "Called", "calledCheckbox"
                AddCheckbox todaySheet.Range("F" & targetRow), "Done", "doneCheckbox"
                AddCheckbox todaySheet.Range("H" & targetRow), "Close", "closeCheckbox"
                
                targetRow = targetRow + 2
            End If
        End With
    Next i
ERR_HANDLE:
    Application.EnableEvents = True
End Sub

Sub calledCheckbox()
    Dim cb As CheckBox
    Dim targetRow As Integer
    Dim nameRow As String
    
    Set cb = todaySheet.CheckBoxes(Application.Caller)
    targetRow = cb.TopLeftCell.Row + 1

    If cb.Value = 1 Then
        todaySheet.Range("A" & targetRow, "H" & targetRow).Interior.ColorIndex = 15
    Else
        todaySheet.Range("A" & targetRow, "H" & targetRow).Interior.ColorIndex = 0
    End If
End Sub

Sub closeCheckbox()
    AssignWorksheetGlobals
    Dim cb As CheckBox
    Dim targetRow As Integer
    Dim ID As String
    Dim dataRow As Integer
    
    Set cb = todaySheet.CheckBoxes(Application.Caller)
    targetRow = cb.TopLeftCell.Row + 1
    
    ID = todaySheet.Range("E" & targetRow).Value
    
    dataRow = GetDataRowWithID(ID)
    dataSheet.Rows(dataRow).Delete
    
    doneCheckbox
    
End Sub

Function GetNameRow(cbName As String) As Integer
    Dim lastDollarSign As Integer
    lastDollarSign = InStrRev(cbName, "$", -1, vbTextCompare)
    
    GetNameRow = CInt(Right(cbName, Len(cbName) - lastDollarSign))
End Function

Function ChangeNameRow(cbName As String, newRow As Integer) As String
    Dim lastDollarSign As Integer
    lastDollarSign = InStrRev(cbName, "$", -1, vbTextCompare)
    
    ChangeNameRow = Left(cbName, lastDollarSign) & newRow
End Function

Sub doneCheckbox()
    AssignWorksheetGlobals
    
    Dim cb As CheckBox
    Dim targetRow As Integer
    Dim nameRow As String
    
    Set cb = todaySheet.CheckBoxes(Application.Caller)
    targetRow = cb.TopLeftCell.Row + 1
    
    nameRow = GetNameRow(cb.Name)
    
    With cb
        If .Value = 1 Then
            todaySheet.Shapes("cb$A$" & nameRow).Delete
            todaySheet.Shapes("cb$H$" & nameRow).Delete
            todaySheet.Shapes("cb$F$" & nameRow).Delete
            
            Rows(targetRow).Delete
            Rows(targetRow).Delete
            
            Dim box As CheckBox
            For Each box In todaySheet.CheckBoxes
                If box.TopLeftCell.Row >= targetRow Then
                    If GetNameRow(box.Name) <> (box.TopLeftCell.Row + 1) Then
                        box.Name = ChangeNameRow(box.Name, GetNameRow(box.Name) - 2)
                    End If
                    
                    If GetNameRow(box.Name) <> (box.TopLeftCell.Row + 1) Then
                        box.Top = box.TopLeftCell.Top + box.TopLeftCell.Height / 2 - box.Height / 2
                    End If
                End If
            Next
            
            Dim btn As Button
            For Each btn In todaySheet.Buttons
                btn.Top = btn.TopLeftCell.Offset(2, 0).Top
            Next btn
        End If
    End With
End Sub

Sub AddCheckbox(c As Range, caption As String, Optional action As String = "")
    AssignWorksheetGlobals
    
    Dim cb As CheckBox
    Set cb = todaySheet.CheckBoxes.Add(0, 1, 100, 20)
    
    With cb
        .caption = caption
        .Top = c.Top + c.Height / 2 - cb.Height / 2
        .Left = c.Left + c.Width / 2 - cb.Width / 2
        .Name = "cb" & c.Address
        .Value = xlOff
        .Display3DShading = False
        .OnAction = action
        .Placement = xlMoveAndSize
    End With
End Sub

Sub CopyRowToToday(rowNum As Long, targetRow As Long)
    'dataSheet.Range("A" & rowNum, "D" & rowNum).Copy todaySheet.Range("B" & targetRow)
    Dim i As Integer
    For i = 1 To 5
        If i < 3 Then
            todaySheet.Cells(targetRow, i + 1).Value = dataSheet.Cells(rowNum, i).Value
        End If
        If i > 3 Then
            todaySheet.Cells(targetRow, i).Value = dataSheet.Cells(rowNum, i).Value
        End If
    Next i
End Sub

Sub CopyRowToData(rowNum As Long, targetRow As Long)
    'todaySheet.Range("B" & rowNum, "E" & rowNum).Copy dataSheet.Range("A" & targetRow)
    Dim i As Integer
    For i = 2 To 5
        If i < 4 Then
            dataSheet.Cells(targetRow, i - 1).Value = todaySheet.Cells(rowNum, i).Value
        End If
        If i >= 4 Then
            dataSheet.Cells(targetRow, i).Value = todaySheet.Cells(rowNum, i).Value
        End If
    Next i
End Sub

Sub CreateWeekly()
    AssignWorksheetGlobals
    
    ClearWeekly
    
    Dim actualDay As Integer
    actualDay = Weekday(Date)
    
    Dim sundayDelta As Integer
    sundayDelta = 1 - actualDay
    
    Dim saturdayDelta As Integer
    saturdayDelta = 7 - actualDay
    
    Dim i As Integer
    Dim currCol As Integer
    currCol = 1
    Dim currDate As Date
    Dim nextWeekDate As Date
    
    For i = sundayDelta To saturdayDelta
        currDate = DateAdd("d", i, Date)
        AddToWeekly currDate, currCol
        
        nextWeekDate = DateAdd("ww", 1, currDate)
        AddToWeekly nextWeekDate, currCol + 16
        
        currCol = currCol + 2
    Next i
End Sub

Sub ClearWeekly()
    Dim numRows As Integer
    Dim currCol As Integer
    currCol = 1
    
    Do While currCol <= 35
        weeklySheet.Range(Cells(3, currCol), Cells(50, currCol)).Clear
        
        currCol = currCol + 2
    Loop

End Sub

Sub AddToWeekly(targetDate As Date, targetCol As Integer)
    Application.EnableEvents = False
    On Error GoTo ERR_HANDLE
    
    AssignWorksheetGlobals
    
    Dim totalTasks As Long
    totalTasks = NumOfRows(dataSheet)
    
    Dim i As Long
    Dim targetRow As Long
    targetRow = 3
    
    For i = 1 To totalTasks
        With dataSheet
            If .Range("A" & i).Value = targetDate Then
                CopyNameToWeekly i, CLng(targetRow), targetCol
                
                targetRow = targetRow + 1
            End If
        End With
    Next i
ERR_HANDLE:
    Application.EnableEvents = True
End Sub

Sub CopyNameToWeekly(dataRow As Long, targetRow As Integer, targetCol As Integer)
    weeklySheet.Cells(targetRow, targetCol).Value = dataSheet.Range("D" & dataRow)
End Sub
