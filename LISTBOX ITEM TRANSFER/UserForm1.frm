VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ListBox Transfer Demo"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddButton_Click()
    Dim i As Integer
    
    If ListBox1.ListIndex = -1 Then Exit Sub
    If Not cbDuplicates Then
'       See if item already exists
        For i = 0 To ListBox2.ListCount - 1
            If ListBox1.Value = ListBox2.List(i) Then
                Beep
                Exit Sub
            End If
        Next i
    End If
    ListBox2.AddItem ListBox1.Value
End Sub

Private Sub ListBox1_Enter()
    RemoveButton.Enabled = False
End Sub

Private Sub ListBox2_Enter()
    RemoveButton.Enabled = True
End Sub

Private Sub RemoveButton_Click()
    If ListBox2.ListIndex = -1 Then Exit Sub
    ListBox2.RemoveItem ListBox2.ListIndex
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    MsgBox "The 'To list' contains " & ListBox2.ListCount & " items."
    Unload Me
End Sub
