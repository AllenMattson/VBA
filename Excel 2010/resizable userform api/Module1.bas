Attribute VB_Name = "Module1"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Const WS_THICKFRAME As Long = &H40000

Public Const GWL_STYLE As Long = (-16)
Public Const SW_SHOW As Long = 5

Sub ShowForm()
   UserForm1.Show vbModeless
End Sub


