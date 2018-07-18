VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Insert Date"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   2985
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InsertButton_Click()
    ActiveCell = DTPicker1.Value
    ActiveCell.Columns.EntireColumn.AutoFit
End Sub

Private Sub UserForm_Initialize()
    DTPicker1.Value = Date
End Sub
