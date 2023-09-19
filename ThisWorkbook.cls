Option Explicit
Private Sub Workbook_Open()
    guaranteeSheet "data"
    guaranteeSheet "Today"
    guaranteeSheet "Weekly"
    
    AssignWorksheetGlobals
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.Name = "Weekly" Then
        CreateWeekly
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.EnableEvents = False
    On Error GoTo ERR_HANDLE

    AssignWorksheetGlobals

    Dim currRow As Range
    Dim targetRow As Long
    Dim ID As String

    If Target.Rows.Count < 100 And Sh.Name = "data" Then
        For Each currRow In Target.Rows
            ID = Sh.Cells(currRow.row, DATA_ID_COLUMN).Value
            targetRow = GetTodayRowWithID(ID)

            If targetRow >= 0 Then
                CopyRowToToday currRow.row, targetRow
            End If
        Next
    End If
    
    If Target.Rows.Count < 100 And Sh.Name = "Today" Then
        For Each currRow In Target.Rows
            ID = Sh.Cells(currRow.row, TODAY_ID_COLUMN).Value
            
            Debug.Print ID
            
            targetRow = GetDataRowWithID(ID)

            If targetRow >= 0 Then
                CopyRowToData currRow.row, targetRow
            End If
        Next
    End If

ERR_HANDLE:
    Application.EnableEvents = True
End Sub

