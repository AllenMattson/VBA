Attribute VB_Name = "Module1"
Option Explicit

Sub DeleteEmptyRows()
    Dim LastRow As Long
    Dim r As Long
    Dim Counter As Long
    Application.ScreenUpdating = False
    LastRow = ActiveSheet.UsedRange.Rows.Count + _
      ActiveSheet.UsedRange.Rows(1).Row - 1
    For r = LastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(Rows(r)) = 0 Then
            Rows(r).Delete
            Counter = Counter + 1
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox Counter & " Empty rows were deleted."
End Sub


