Attribute VB_Name = "Module1"
Option Explicit

Sub ShowShortcutMenuItems()
  Dim Row As Long
  Dim Cbar As CommandBar
  Dim ctl As CommandBarControl
  Range("A1:G1") = Array("Index", "Name", "ID", "Caption", _
    "Type", "Enabled", "Visible")
  Row = 2
  Application.ScreenUpdating = False
  For Each Cbar In Application.CommandBars
    If Cbar.Type = 2 Then
      For Each ctl In Cbar.Controls
          Cells(Row, 1) = Cbar.Index
          Cells(Row, 2) = Cbar.Name
          Cells(Row, 3) = ctl.ID
          Cells(Row, 4) = ctl.Caption
          If ctl.Type = 1 Then
              Cells(Row, 5) = "Button"
          Else
              Cells(Row, 5) = "Submenu"
          End If
          Cells(Row, 6) = ctl.Enabled
          Cells(Row, 7) = ctl.Visible
         Row = Row + 1
     Next ctl
   End If
  Next Cbar
End Sub

Sub ShowCaptions()
    Dim txt As String
    Dim ctl As CommandBarControl
    For Each ctl In CommandBars("Ply").Controls
        txt = txt & ctl.Caption & vbNewLine
    Next ctl
    MsgBox txt
End Sub


