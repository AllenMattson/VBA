VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "SpinButton / TextBox Demo"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me.SpinButton1
'       Specify upper and lower limits
        .Min = -10
        .Max = 10
'       Change label
        Me.Label1.Caption = "Specify a value between " _
          & .Min & " and " & .Max & ":"
'       Initialize Spinner
        .Value = 1
'       Initialize TextBox
        Me.TextBox1.Text = .Value
    End With
End Sub

Private Sub TextBox1_Change()
    Dim NewVal As Long
    If IsNumeric(Me.TextBox1.Text) Then
        NewVal = Val(Me.TextBox1.Text)
        If NewVal >= Me.SpinButton1.Min And _
            NewVal <= Me.SpinButton1.Max Then _
            Me.SpinButton1.Value = NewVal
    End If
End Sub

Private Sub TextBox1_Enter()
'   Selects all text when user enters TextBox
    Me.TextBox1.SelStart = 0
    Me.TextBox1.SelLength = Len(Me.TextBox1.Text)
End Sub

Private Sub SpinButton1_Change()
    Me.TextBox1.Text = Me.SpinButton1.Value
End Sub


Private Sub OKButton_Click()
'   Enter the value into the active cell
    If CStr(Me.SpinButton1.Value) = Me.TextBox1.Text Then
        ActiveCell = Me.SpinButton1.Value
        Unload Me
    Else
        MsgBox "Invalid entry.", vbCritical
        Me.TextBox1.SetFocus
        Me.TextBox1.SelStart = 0
        Me.TextBox1.SelLength = Len(Me.TextBox1.Text)
    End If
End Sub

