VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UnloadButton_Click()
    Unload Me
End Sub

Private Sub HideButton_Click()
    Me.Hide
End Sub

Private Sub ShowAnotherButton_Click()
    UserForm2.Show
End Sub

Private Sub UserForm_Activate()
    MsgBox "The Activate Event has occured.", vbInformation
End Sub

Private Sub UserForm_Deactivate()
    MsgBox "The Deactivate Event has occured.", vbInformation
End Sub

Private Sub UserForm_Initialize()
    MsgBox "The Initialize Event has occured.", vbInformation
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    MsgBox "The QueryClose Event has occured.", vbInformation
End Sub

Private Sub UserForm_Terminate()
    MsgBox "The Terminate Event has occured.", vbInformation
End Sub
