VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ListBox Move Item Demo"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUp_Click()
    Dim lSelected As Long
    Dim sSelected As String
    
'   Store the currently selected item
    lSelected = Me.lbxItems.ListIndex
    sSelected = Me.lbxItems.Value

'   Remove the selected item
    Me.lbxItems.RemoveItem lSelected
'   Add back the item one above
    Me.lbxItems.AddItem sSelected, lSelected - 1
'   Reselect the moved item
    Me.lbxItems.ListIndex = lSelected - 1
End Sub

Private Sub cmdDown_Click()
    Dim lSelected As Long
    Dim sSelected As String
    
'   Store the currently selected item
    lSelected = Me.lbxItems.ListIndex
    sSelected = Me.lbxItems.Value

'   Remove the selected item
    Me.lbxItems.RemoveItem lSelected
'   Add back the item one below
    Me.lbxItems.AddItem sSelected, lSelected + 1
'   Reselect the moved item
    Me.lbxItems.ListIndex = lSelected + 1
End Sub

Private Sub cmdUp_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdUp_Click
End Sub

Private Sub cmdDown_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdDown_Click
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub lbxItems_Click()
    Me.cmdDown.Enabled = Me.lbxItems.ListIndex > -1 _
        And Me.lbxItems.ListIndex < Me.lbxItems.ListCount - 1
    Me.cmdUp.Enabled = Me.lbxItems.ListIndex > -1 _
        And Me.lbxItems.ListIndex > 0
End Sub
