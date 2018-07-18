VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinActivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents AppEvents As Application
Attribute AppEvents.VB_VarHelpID = -1

Private Declare PtrSafe Function SetParent Lib "user32" _
   (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr


Private Sub AppEvents_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    'Change the parent of the UserForm to be the newly activated window
    If Val(Application.Version) >= 15 Then SetParent UserFormHandle, Wn.hWnd
    UserForm1.Label1 = ActiveWorkbook.Name
End Sub


