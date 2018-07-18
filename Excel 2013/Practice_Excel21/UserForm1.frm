VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_MouseDown(ByVal Button _
        As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)
    If Button = 2 Then
        Call Show_ShortMenu
    Else
        MsgBox "You must right-click this button."
    End If
End Sub


