VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Display a Control Panel Dialog"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub RunButton_Click()
    Dim Arg As String
    Dim TaskID
    Arg = ThisWorkbook.Sheets("Data").Cells(ListBox1.ListIndex + 1, 2)
    On Error Resume Next
    TaskID = Shell(Arg)
    If Err <> 0 Then
        MsgBox ("Cannot start the application.")
    End If
End Sub

Private Sub UserForm_Initialize()
    ListBox1.ListIndex = 0
End Sub
