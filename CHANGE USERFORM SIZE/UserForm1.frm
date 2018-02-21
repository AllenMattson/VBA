VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Print Sheets"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SmallSize As Integer = 124
Const LargeSize As Integer = 164


Private Sub UserForm_Initialize()
    Dim sht As Worksheet
    
    For Each sht In ActiveWorkbook.Worksheets
        ListBox1.AddItem sht.Name
    Next sht
    Me.Height = SmallSize
End Sub

Private Sub OptionsButton_Click()
    If OptionsButton.Caption = "Options >>" Then
        Me.Height = LargeSize
        OptionsButton.Caption = "<< Options"
    Else
        Me.Height = SmallSize
        OptionsButton.Caption = "Options >>"
    End If
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    Dim i As Integer
    
    MsgBox "UserForm demo only - Sheets will not be printed."
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            With Sheets(ListBox1.List(i))
                .PageSetup.PrintGridlines = cbGridlines
                If obLandscape Then .PageSetup.Orientation = xlLandscape
                If obPortrait Then .PageSetup.Orientation = xlPortrait
'               .PrintOut
            End With
        End If
    Next i
    Unload Me
End Sub

