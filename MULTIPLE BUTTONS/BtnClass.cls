VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BtnClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ButtonGroup As CommandButton
Attribute ButtonGroup.VB_VarHelpID = -1

Private Sub ButtonGroup_Click()
    Dim Msg As String
    
    Msg = "You clicked " & ButtonGroup.Name & vbCrLf & vbCrLf
    Msg = Msg & "Caption: " & ButtonGroup.Caption & vbCrLf
    Msg = Msg & "Left Position: " & ButtonGroup.Left & vbCrLf
    Msg = Msg & "Top Position: " & ButtonGroup.Top
    MsgBox Msg, vbInformation, ButtonGroup.Name
End Sub
