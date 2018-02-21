VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    'simulated procedure
    MsgBox "Shortcut menu items have been added."
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim Msg As String
    If Me.Saved = False Then
        Msg = "Do you want to save the changes you made to "
        Msg = Msg & Me.Name & "?"
        Ans = MsgBox(Msg, vbQuestion + vbYesNoCancel)
        Select Case Ans
            Case vbYes
                Me.Save
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    End If
    Call DeleteShortcutMenuItems
    Me.Saved = True
End Sub

Sub DeleteShortcutMenuItems()
    'simulated procedure
    MsgBox "Shortcut menu items have been deleted."
End Sub

