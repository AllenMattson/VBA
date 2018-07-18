VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ListBox Demo"
   ClientHeight    =   2520
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

Private Sub optExtend_Click()
    Me.ListBox1.MultiSelect = fmMultiSelectExtended
End Sub

Private Sub optMulti_Click()
    Me.ListBox1.MultiSelect = fmMultiSelectMulti
End Sub

Private Sub optSingle_Click()
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
End Sub

Private Sub cmdOK_Click()
    Dim Msg As String
    Dim i As Long
    
    If Me.ListBox1.ListIndex = -1 Then
        Msg = "Nothing"
    Else
        For i = 0 To Me.ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then _
              Msg = Msg & Me.ListBox1.List(i) & vbNewLine
        Next i
    End If
    MsgBox "You selected: " & vbNewLine & Msg
    Unload Me
End Sub
