VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()
    guaranteeSheet "data"
    guaranteeSheet "Today"
    guaranteeSheet "Weekly"
    
    AssignWorksheetGlobals
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.EnableEvents = False
    On Error GoTo ERR_HANDLE

    AssignWorksheetGlobals

    Dim currRow As Range
    Dim targetRow As Long
    Dim ID As String

    If Target.Rows.Count < 100 And Sh.Name = "data" Then
        Debug.Print ("here!")
        For Each currRow In Target.Rows
            ID = Sh.Cells(currRow.row, "D").Value
            targetRow = GetTodayRowWithID(ID)
            
            Debug.Print (targetRow)

            If targetRow >= 0 Then
                CopyRowToToday currRow.row, targetRow
            End If
        Next
    End If
    
    If Target.Rows.Count < 100 And Sh.Name = "Today" Then
        For Each currRow In Target.Rows
            ID = Sh.Cells(currRow.row, "E").Value
            targetRow = GetDataRowWithID(ID)

            If targetRow >= 0 Then
                CopyRowToData currRow.row, targetRow
            End If
        Next
    End If

ERR_HANDLE:
    Application.EnableEvents = True
End Sub
