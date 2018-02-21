VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Random Number Generator"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
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
    Call UpdateLabels
    Me.Height = 124
End Sub

Private Sub CancelButton_Click()
    End
    Unload Me
End Sub

Private Sub OKButton_Click()
'   Inserts random numbers on the active worksheet
    Dim Counter As Long
    Dim RowMax As Long, ColMax As Long
    Dim r As Long, c As Long
    Dim PctDone As Double
    
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    
'   Increase height
    Me.Height = 172
    
'   Do the work
    Cells.Clear
    Counter = 1
    RowMax = SpinButton1.Value
    ColMax = SpinButton2.Value
    For r = 1 To RowMax
        For c = 1 To ColMax
            Cells(r, c) = Int(Rnd * 1000)
            Counter = Counter + 1
        Next c
        PctDone = Counter / (RowMax * ColMax)
        Call UpdateProgress(PctDone)
    Next r
    Unload UserForm1
End Sub

Sub UpdateProgress(Pct)
    With UserForm1
      .FrameProgress.Caption = Format(Pct, "0%")
      .LabelProgress.Width = Pct * (.FrameProgress.Width - 10)
      DoEvents
    End With
End Sub

Private Sub SpinButton1_Change()
    Call UpdateLabels
End Sub

Private Sub SpinButton2_Change()
    Call UpdateLabels
End Sub


Sub UpdateLabels()
    lblRows = SpinButton1.Value & " Rows"
    lblColumns = SpinButton2.Value & " Columns"
End Sub
