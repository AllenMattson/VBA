Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateRandomNumbers()
'   Inserts random numbers on the active worksheet
    Dim Counter As Long
    Const RowMax As Long = 500
    Const ColMax As Long = 40
    Dim r As Long, c As Long
    Dim PctDone As Double
    
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Cells.Clear
    Counter = 1
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
        .Repaint
    End With
End Sub

Sub ShowUserForm()
    With UserForm1
        'Use a color from the workbook's theme
        .LabelProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        .LabelProgress.Width = 0
        .Show
    End With
End Sub
