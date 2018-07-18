VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'   NOTE: This button is is hidden
'   behind another object.
'   Including this button makes it possible to
'   cancel the splash screen by pressing Escape
    Unload Me
End Sub

Private Sub UserForm_Activate()
    Application.OnTime Now + TimeSerial(0, 0, 5), "KillTheForm"
End Sub



