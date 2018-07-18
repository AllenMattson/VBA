VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UProgressMp 
   Caption         =   "Random Number Generator"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   OleObjectBlob   =   "UProgressMp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UProgressMp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    UpdateLabels
End Sub

Private Sub cmdOK_Click()
'   Inserts random numbers on the active worksheet
    Dim Counter As Long
    Dim r As Long, c As Long
    Dim PctDone As Double
    
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    
'   Display the progress panel
    Me.mpProgress.Value = 1
    
'   Do the work
    ActiveSheet.Cells.Clear
    Counter = 1
    For r = 1 To Me.sbRows.Value
        For c = 1 To Me.sbColumns.Value
            ActiveSheet.Cells(r, c) = Int(Rnd * 1000)
            Counter = Counter + 1
        Next c
        PctDone = Counter / (Me.sbRows.Value * Me.sbColumns.Value)
        UpdateProgress PctDone
    Next r
    Unload Me
End Sub

Sub UpdateProgress(Pct As Double)
    With Me
        .frmProgress.Caption = Format(Pct, "0%")
        .lblProgress.Width = Pct * (.frmProgress.Width - 10)
        .Repaint
    End With
End Sub

Private Sub sbRows_Change()
    UpdateLabels
End Sub

Private Sub sbColumns_Change()
    UpdateLabels
End Sub

Sub UpdateLabels()
    Me.lblRows.Caption = Me.sbRows.Value & " Rows"
    Me.lblColumns.Caption = Me.sbColumns.Value & " Columns"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

