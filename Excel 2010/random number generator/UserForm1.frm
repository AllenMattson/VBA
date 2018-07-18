VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Random Number Generator"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Stopped As Boolean

Private Sub StartStopButton_Click()
    Dim Low As Double, Hi As Double
    
    If StartStopButton.Caption = "Start" Then
'       validate low and hi values
        If Not IsNumeric(TextBox1.Text) Then
            MsgBox "Non-numeric starting value.", vbInformation
            With TextBox1
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            Exit Sub
        End If
        
        If Not IsNumeric(TextBox2.Text) Then
            MsgBox "Non-numeric ending value.", vbInformation
            With TextBox2
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            Exit Sub
        End If
        
'       Make sure they aren't in the wrong order
        Low = Application.Min(Val(TextBox1.Text), Val(TextBox2.Text))
        Hi = Application.Max(Val(TextBox1.Text), Val(TextBox2.Text))
        
'       Adjust font size, if necessary
        Select Case Application.Max(Len(TextBox1.Text), Len(TextBox2.Text))
            Case Is < 5: Label1.Font.Size = 72
            Case 5: Label1.Font.Size = 60
            Case 6: Label1.Font.Size = 48
            Case Else: Label1.Font.Size = 36
        End Select
        
        StartStopButton.Caption = "Stop"
        Stopped = False
        Randomize
        Do Until Stopped
            Label1.Caption = Int((Hi - Low + 1) * Rnd + Low)
            DoEvents ' Causes the animation
        Loop
    Else
        Stopped = True
        StartStopButton.Caption = "Start"
    End If
End Sub

Private Sub CancelButton_Click()
    Stopped = True
    Unload Me
End Sub


Private Sub UserForm_Terminate()
    Stopped = True
End Sub

