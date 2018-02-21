VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ListBox Demo"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub obExtend_Click()
    ListBox1.MultiSelect = fmMultiSelectExtended
End Sub

Private Sub obMulti_Click()
    ListBox1.MultiSelect = fmMultiSelectMulti
End Sub

Private Sub obSingle_Click()
    ListBox1.MultiSelect = fmMultiSelectSingle
End Sub

Private Sub OKButton_Click()
    Dim Msg As String
    Dim i As Integer
    
    If ListBox1.ListIndex = -1 Then
        Msg = "Nothing"
    Else
        Msg = ""
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then _
              Msg = Msg & ListBox1.List(i) & vbCrLf
        Next i
    End If
    MsgBox "You selected: " & vbCrLf & Msg
    Unload Me
End Sub


