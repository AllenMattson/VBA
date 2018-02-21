VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Chart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Chart_Activate()
    MsgBox "You've activated the chart sheet."
End Sub

Private Sub Chart_Deactivate()
    MsgBox "It looks like you want to leave the " _
     & "chart sheet."
End Sub

Private Sub Chart_Select(ByVal ElementID As Long, _
            ByVal Arg1 As Long, ByVal Arg2 As Long)
    If Arg1 <> 0 And Arg2 <> 0 Then
        MsgBox ElementID & ", " & Arg1 & ", " & Arg2
    End If
    If ElementID = 4 Then
        MsgBox "You've selected the chart title."
    ElseIf ElementID = 24 Then
        MsgBox "You've selected the chart legend."
    ElseIf ElementID = 12 Then
        MsgBox "You've selected the legend key."
    ElseIf ElementID = 13 Then
        MsgBox "You've selected the legend entry."
    End If
End Sub

Private Sub Chart_Calculate()
    MsgBox "The data in your spreadsheet has " & vbCrLf _
        & "changed. Your chart has been updated."
End Sub

Private Sub Chart_BeforeRightClick(Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Chart_MouseDown(ByVal Button As Long, _
    ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    If Button = 1 Then
         MsgBox "You pressed the left mouse button. "
    ElseIf Button = 2 Then
        MsgBox "You pressed the right mouse button. "
    Else
        MsgBox "You pressed the middle mouse button. "
    End If
End Sub

