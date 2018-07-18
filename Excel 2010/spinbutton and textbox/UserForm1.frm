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
    With SpinButton1
'       Specify upper and lower limits
        .Min = 1
        .Max = 100
'       Change label
        Label1.Caption = "Specify a value between " _
          & .Min & " and " & .Max & ":"
'       Initialize Spinner
        .Value = 1
'       Initialize TextBox
        TextBox1.Text = .Value
    End With
End Sub

Private Sub TextBox1_Change()
    Dim NewVal As Integer
    
    NewVal = Val(TextBox1.Text)
    If NewVal >= SpinButton1.Min And _
        NewVal <= SpinButton1.Max Then _
        SpinButton1.Value = NewVal
End Sub

Private Sub TextBox1_Enter()
'   Selects all text when user enters TextBox
    TextBox1.SelStart = 0
    TextBox1.SelLength = Len(TextBox1.Text)
End Sub

Private Sub SpinButton1_Change()
    TextBox1.Text = SpinButton1.Value
End Sub


Private Sub OKButton_Click()
'   Enter the value into the active cell
    If CStr(SpinButton1.Value) = TextBox1.Text Then
        ActiveCell = SpinButton1.Value
        Unload Me
    Else
        MsgBox "Invalid entry.", vbCritical
        TextBox1.SetFocus
        TextBox1.SelStart = 0
        TextBox1.SelLength = Len(TextBox1.Text)
    End If
End Sub

