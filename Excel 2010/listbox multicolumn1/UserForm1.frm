VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Multicolumn ListBox Demo"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_Click()
    If ListBox1.ListIndex <> -1 Then
        Range(Cells(ListBox1.ListIndex + 2, 1), Cells(ListBox1.ListIndex + 2, 3)).Select
    End If
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
