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

Private Stopped As Boolean

Private Sub cmdCancel_Click()
    Stopped = True
    Unload Me
End Sub

Private Sub cmdStartStop_Click()
    Dim Low As Double, Hi As Double
    Dim wf As WorksheetFunction
    
    Set wf = Application.WorksheetFunction
    
    If Me.cmdStartStop.Caption = "Start" Then
'       validate low and hi values
        If Not IsNumeric(Me.tbxStart.Text) Then
            MsgBox "Non-numeric starting value.", vbInformation
            With Me.tbxStart
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            Exit Sub
        End If
        
        If Not IsNumeric(Me.tbxEnd.Text) Then
            MsgBox "Non-numeric ending value.", vbInformation
            With Me.tbxEnd
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            Exit Sub
        End If
        
'       Make sure they aren't in the wrong order
        Low = wf.Min(Val(Me.tbxStart.Text), Val(Me.tbxEnd.Text))
        Hi = wf.Max(Val(Me.tbxStart.Text), Val(Me.tbxEnd.Text))
        
'       Adjust font size, if necessary
        Select Case _
            wf.Max(Len(Me.tbxStart.Text), Len(Me.tbxEnd.Text))
            
            Case Is < 5: Me.lblRandom.Font.Size = 72
            Case 5: Me.lblRandom.Font.Size = 60
            Case 6: Me.lblRandom.Font.Size = 48
            Case Else: Me.lblRandom.Font.Size = 36
        End Select
        
        Me.cmdStartStop.Caption = "Stop"
        Stopped = False
        Randomize
        Do Until Stopped
            Me.lblRandom.Caption = _
                Int((Hi - Low + 1) * Rnd + Low)
            DoEvents ' Causes the animation
        Loop
    Else
        Stopped = True
        Me.cmdStartStop.Caption = "Start"
    End If
End Sub

Private Sub UserForm_Terminate()
    Stopped = True
End Sub

