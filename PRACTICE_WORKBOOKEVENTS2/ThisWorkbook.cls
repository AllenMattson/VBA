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

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    If MsgBox("Do you want to place  " & vbCrLf _
        & "the new sheet at the beginning " & vbCrLf _
        & "of the workbook?", vbYesNo) = vbYes Then
                Sh.Move before:=ThisWorkbook.Sheets(1)
    Else
        Sh.Move After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count)
        MsgBox Sh.Name & _
        " is now the last sheet in the workbook."
        
    End If
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    Wn.GridlineColor = vbYellow
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    MsgBox "You have just deactivated " & Wn.Caption
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    If Wn.WindowState <> xlMaximized Then
        Wn.Left = 0
        Wn.Top = 0
    End If
End Sub

