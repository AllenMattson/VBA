VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
   Dim lStyle As Long
   Dim hWnd As Long

'  Make it resizable
   hWnd = FindWindow("ThunderDFrame", Me.Caption)
   lStyle = GetWindowLong(hWnd, GWL_STYLE)
   lStyle = lStyle Or WS_THICKFRAME
   SetWindowLong hWnd, GWL_STYLE, lStyle
   ShowWindow hWnd, SW_SHOW
   DrawMenuBar hWnd
   SetFocus hWnd

End Sub


