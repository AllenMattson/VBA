VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select a Contact"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8325
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private maContacts As Variant

Private Sub UserForm_Initialize()
    maContacts = Sheet1.ListObjects("tblContacts").DataBodyRange.Value
    FillContacts
End Sub

Private Sub tbxSearch_Change()
    FillContacts Me.tbxSearch.Text
End Sub

Private Sub FillContacts(Optional sFilter As String = "*")
    Dim i As Long, j As Long
    
    'Clear any existing entries in the ListBox
    Me.lbxContacts.Clear
    'Loop through all the rows and columns of the contact list
    For i = LBound(maContacts, 1) To UBound(maContacts, 1)
        For j = 1 To 4
            'Compare the contact to the filter
            If UCase(maContacts(i, j)) Like UCase("*" & sFilter & "*") Then
                'Add it to the ListBox
                With Me.lbxContacts
                    .AddItem maContacts(i, 1)
                    .List(.ListCount - 1, 1) = maContacts(i, 2)
                    .List(.ListCount - 1, 2) = maContacts(i, 3)
                    .List(.ListCount - 1, 3) = maContacts(i, 4)
                End With
                'If any column matched, skip the rest of the columns
                'and move to the next contact
                Exit For
            End If
        Next j
    Next i
    'Select the first contact
    If Me.lbxContacts.ListCount > 0 Then Me.lbxContacts.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.lbxContacts.ListIndex >= 0 Then
        MsgBox "You selected " & Me.lbxContacts.Value
        Unload Me
    End If
End Sub
