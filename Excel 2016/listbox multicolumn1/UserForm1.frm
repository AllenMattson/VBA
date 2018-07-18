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
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub lbxProducts_Click()
    If Me.lbxProducts.ListIndex > -1 Then
        Sheet1.Cells(Me.lbxProducts.ListIndex + 2, 1).Resize(, 3).Select
    End If
End Sub

Private Sub UserForm_Initialize()

'   If the user has selected a cell in the range, select that entry
'   in the listbox. Otherwise, select the first entry
    If Intersect(ActiveCell, Sheet1.Range(Me.lbxProducts.RowSource)) Is Nothing Then
        Me.lbxProducts.ListIndex = 0
    Else
        Me.lbxProducts.ListIndex = ActiveCell.Row - 2
    End If
    
End Sub
