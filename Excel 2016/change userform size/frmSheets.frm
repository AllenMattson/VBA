VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheets 
   Caption         =   "Print Sheets"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   OleObjectBlob   =   "frmSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSheets"
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
        Me.lbxSheets.AddItem sht.Name
    Next sht
    Me.Height = SmallSize
End Sub

Private Sub cmdOptions_Click()
    Const OptionsHidden As String = "Options >>"
    Const OptionsShown As String = "<< Options"
    
    If Me.cmdOptions.Caption = OptionsHidden Then
        Me.Height = LargeSize
        Me.cmdOptions.Caption = OptionsShown
    Else
        Me.Height = SmallSize
        Me.cmdOptions.Caption = OptionsHidden
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    MsgBox "UserForm demo only - Sheets will not be printed."
    For i = 0 To Me.lbxSheets.ListCount - 1
        If Me.lbxSheets.Selected(i) Then
            With Sheets(Me.lbxSheets.List(i))
                .PageSetup.PrintGridlines = Me.chkGridlines.Value
                If Me.optLandscape.Value Then .PageSetup.Orientation = xlLandscape
                If Me.optPortrait.Value Then .PageSetup.Orientation = xlPortrait
'               .PrintOut
            End With
        End If
    Next i
    Unload Me
End Sub

