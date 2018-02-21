VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim shtName As String

Private Sub Worksheet_Activate()
  shtName = ActiveSheet.Name
  Range("B2").Select
End Sub


Private Sub Worksheet_Deactivate()
  MsgBox "You deactivated " & _
    shtName & "." & vbCrLf & _
    "You switched to " & _
    ActiveSheet.Name & "."
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal _
     Target As Range, Cancel As Boolean)
  If Target.Address = "$C$9" Then
    MsgBox "No double-clicking, please."
    Cancel = True
  Else
    MsgBox "You may edit this cell."
  End If
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal _
    Target As Range, Cancel As Boolean)

  With Application.CommandBars("Cell")
    .Reset
    If Target.Rows.Count > 1 Or _
        Target.Columns.Count > 1 Then
        With .Controls.Add(Type:=msoControlButton, _
            before:=1, temporary:=True)
          .Caption = "Print..."
          .OnAction = "PrintMe"
        End With
    End If
  End With
End Sub


