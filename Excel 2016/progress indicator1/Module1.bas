Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateRandomNumbers()
'   Inserts random numbers on the active worksheet
    Dim Counter As Long
    Dim r As Long, c As Long
    Dim PctDone As Double
    Const RowMax As Long = 500
    Const ColMax As Long = 40
    
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    ActiveSheet.Cells.Clear
    UProgress.SetDescription "Generating random numbers..."
    UProgress.Show vbModeless
    Counter = 1
    For r = 1 To RowMax
        For c = 1 To ColMax
            ActiveSheet.Cells(r, c) = Int(Rnd * 1000)
            Counter = Counter + 1
        Next c
        PctDone = Counter / (RowMax * ColMax)
        UProgress.UpdateProgress PctDone
    Next r
    Unload UProgress
End Sub

