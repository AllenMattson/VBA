VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCmdBarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmdBtnEvents As CommandBarButton
Attribute cmdBtnEvents.VB_VarHelpID = -1

Private Sub cmdBtnEvents_Click(ByVal Ctrl As Office.CommandBarButton, _
        CancelDefault As Boolean)
    On Error Resume Next
    MsgBox "inside class module"
    ' run the procedure specified in the onAction property
    Application.Run Ctrl.OnAction

    ' specify that we already handled this event
    CancelDefault = True
End Sub



