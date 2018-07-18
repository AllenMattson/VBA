VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Move The Images"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7545
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'UserForm position
Dim OldX As Double, OldY As Double

Private Sub Image1_MouseDown(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Starting position when button is pressed
    OldX = X
    OldY = Y
    Image1.ZOrder 0
End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Move the image
    If Button = 1 Then
        Image1.Left = Image1.Left + (X - OldX)
        Image1.Top = Image1.Top + (Y - OldY)
    End If
End Sub

Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Starting position when button is pressed
    OldX = X
    OldY = Y
    Image2.ZOrder 0
End Sub

Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Move the image
    If Button = 1 Then
        Image2.Left = Image2.Left + (X - OldX)
        Image2.Top = Image2.Top + (Y - OldY)
    End If
End Sub

Private Sub Image3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Starting position when button is pressed
    OldX = X
    OldY = Y
    Image3.ZOrder 0
End Sub

Private Sub Image3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Move the image
    If Button = 1 Then
        Image3.Left = Image3.Left + (X - OldX)
        Image3.Top = Image3.Top + (Y - OldY)
    End If
End Sub



Private Sub CloseButton_Click()
    Unload Me
End Sub



