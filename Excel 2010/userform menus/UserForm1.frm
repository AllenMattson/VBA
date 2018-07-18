VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Button Menu Demo"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    Call Macro1
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
    Call Macro2
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    Me.Hide
    Call Macro3
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    Me.Hide
    Call Macro4
    Unload Me
End Sub

Private Sub CommandButton5_Click()
    Me.Hide
    Call Macro5
    Unload Me
End Sub

Private Sub CommandButton6_Click()
    Me.Hide
    Call Macro6
    Unload Me
End Sub
