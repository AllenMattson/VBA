Attribute VB_Name = "Module1"
Option Explicit

Sub ShowShortcutMenuNames()
    Dim Row As Long
    Dim cbar As CommandBar
    Row = 1
    For Each cbar In Application.CommandBars
        If cbar.Type = 2 Then
            Cells(Row, 1) = cbar.Index
            Cells(Row, 2) = cbar.Name
            Row = Row + 1
        End If
    Next cbar
End Sub

