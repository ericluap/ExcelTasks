Option Explicit

Dim dataSheet As Worksheet
Dim todaySheet As Worksheet
Dim weeklySheet As Worksheet

' which column in the data sheet stores the info
Public DATA_NAME_COLUMN As Long
Public DATA_DATE_COLUMN As Long
Public DATA_NOTE_COLUMN As Long
Public DATA_ID_COLUMN As Long

' which column in the today sheet stores the info
Public TODAY_NAME_COLUMN As Long
Public TODAY_DATE_COLUMN As Long
Public TODAY_NOTE_COLUMN As Long
Public TODAY_ID_COLUMN As Long
Public TODAY_CALLED_COLUMN As Long
Public TODAY_DONE_COLUMN As Long
Public TODAY_CLOSE_COLUMN As Long

' first row to start tasks on
Public TODAY_INITIAL_ROW As Long
' how many blank rows in between each task
Public TODAY_ROW_SPACING As Long

Sub AssignWorksheetGlobals()
    Set dataSheet = Worksheets("data")
    Set todaySheet = Worksheets("Today")
    Set weeklySheet = Worksheets("Weekly")
    
    DATA_NAME_COLUMN = 3
    DATA_DATE_COLUMN = 1
    DATA_NOTE_COLUMN = 2
    DATA_ID_COLUMN = DATA_NAME_COLUMN
    
    TODAY_NAME_COLUMN = 4
    TODAY_DATE_COLUMN = 2
    TODAY_NOTE_COLUMN = 3
    TODAY_ID_COLUMN = TODAY_NAME_COLUMN
    
    TODAY_INITIAL_ROW = 2
    TODAY_ROW_SPACING = 1
End Sub

' for a given column c and value v, find the first row r such that (r,c) = v
Function GetRowWithValueInColumn(ws As Worksheet, col As Long, val As String)
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
    
    GetTodayRowWithID = GetRowWithValueInColumn(todaySheet, TODAY_ID_COLUMN, ID)
End Function

Function GetDataRowWithID(ID As String) As Long
    AssignWorksheetGlobals
    
    GetDataRowWithID = GetRowWithValueInColumn(dataSheet, DATA_ID_COLUMN, ID)
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

Function NumOfRows(ws As Worksheet) As Long
    AssignWorksheetGlobals
    
    Dim totalRow As Long
    
    totalRow = 0
    
    With ws
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, DATA_NAME_COLUMN).End(xlUp).row, totalRow)
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, DATA_DATE_COLUMN).End(xlUp).row, totalRow)
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, DATA_NOTE_COLUMN).End(xlUp).row, totalRow)
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, DATA_ID_COLUMN).End(xlUp).row, totalRow)
        
        totalRow = WorksheetFunction.Max(Cells(.Rows.Count, TODAY_NAME_COLUMN).End(xlUp).row, totalRow)
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, TODAY_DATE_COLUMN).End(xlUp).row, totalRow)
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, TODAY_NOTE_COLUMN).End(xlUp).row, totalRow)
        totalRow = WorksheetFunction.Max(.Cells(.Rows.Count, TODAY_ID_COLUMN).End(xlUp).row, totalRow)
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
    
    targetRow = TODAY_INITIAL_ROW
    
    For i = 1 To totalTasks
        
        With dataSheet
            If .Cells(i, DATA_DATE_COLUMN).Value = Date Then
                CopyRowToToday i, CLng(targetRow)
                
                AddCheckbox todaySheet.Range("A" & targetRow), "Called", "calledCheckbox"
                AddCheckbox todaySheet.Range("F" & targetRow), "Done", "doneCheckbox"
                AddCheckbox todaySheet.Range("H" & targetRow), "Close", "closeCheckbox"
                
                targetRow = targetRow + 1 + TODAY_ROW_SPACING
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
    targetRow = cb.TopLeftCell.row + 1

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
    targetRow = cb.TopLeftCell.row + 1
    
    ID = todaySheet.Range("E" & targetRow).Value
    
    dataRow = GetDataRowWithID(ID)
    dataSheet.Rows(dataRow).Delete
    
    doneCheckbox
    
End Sub

Sub doneCheckbox()
    Dim cb As CheckBox
    Dim targetRow As Integer
    Dim nameRow As String
    
    Set cb = todaySheet.CheckBoxes(Application.Caller)
    targetRow = cb.TopLeftCell.row + 1
    
    nameRow = Right(cb.Name, 1)
    
    With cb
        If .Value = 1 Then
            todaySheet.Shapes("cb$A$" & nameRow).Delete
            todaySheet.Shapes("cb$H$" & nameRow).Delete
            todaySheet.Shapes("cb$F$" & nameRow).Delete
            
            Rows(targetRow).Delete
            Rows(targetRow).Delete
            
            Dim box As CheckBox
            For Each box In todaySheet.CheckBoxes
                If box.TopLeftCell.row >= targetRow Then
                    box.Top = box.TopLeftCell.Top + box.TopLeftCell.Height / 2 - box.Height / 2
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
    Dim cb As CheckBox
    Set cb = todaySheet.CheckBoxes.Add(0, 1, 100, 0)
    
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
    For i = 1 To 4
        todaySheet.Cells(targetRow, i + 1).Value = dataSheet.Cells(rowNum, i).Value
    Next i
End Sub

Sub CopyRowToData(rowNum As Long, targetRow As Long)
    'todaySheet.Range("B" & rowNum, "E" & rowNum).Copy dataSheet.Range("A" & targetRow)
    Dim i As Integer
    For i = 2 To 5
        dataSheet.Cells(targetRow, i - 1).Value = todaySheet.Cells(rowNum, i).Value
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
            If .Cells(i, DATA_DATE_COLUMN).Value = targetDate Then
                CopyNameToWeekly i, CLng(targetRow), targetCol
                
                targetRow = targetRow + 1
            End If
        End With
    Next i
ERR_HANDLE:
    Application.EnableEvents = True
End Sub

Sub CopyNameToWeekly(dataRow As Long, targetRow As Integer, targetCol As Integer)
    weeklySheet.Cells(targetRow, targetCol).Value = dataSheet.Cells(dataRow, DATA_NAME_COLUMN)
End Sub

