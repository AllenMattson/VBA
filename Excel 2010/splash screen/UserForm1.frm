VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
'   NOTE: This button is is "hidden"
'   behind another object
'   Including this button makes it possible to
'   cancel the splash screen by pressing Escape
    Unload Me
End Sub

Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:05"), "KillTheForm"
End Sub



