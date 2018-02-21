VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Export Charts Add-In"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.Caption = APPNAME
End Sub


Private Sub OKButton_Click()
    If cbMessage Then
        SaveSetting APPNAME, "Settings", "ShowMessage", "No"
    Else
        SaveSetting APPNAME, "Settings", "ShowMessage", "Yes"
    End If
    Unload Me
End Sub

Private Sub HelpButton_Click()
    Call ShowHelp
End Sub


