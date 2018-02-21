Attribute VB_Name = "Module1"
Option Explicit
Dim r As Long
Public ChartIsAnimated As Boolean

Sub AnimateButton_Click()
    Dim i As Integer
    If ChartIsAnimated Then GoTo Finish
    ChartIsAnimated = True
    On Error GoTo Finish
    Application.EnableCancelKey = xlErrorHandler
    i = 1
    Do
        Range("animate").Value = i * Range("Speed") * 0.1
        DoEvents
        i = i + 1
        DoEvents
        If Not ChartIsAnimated Then GoTo Finish
    Loop
Finish:
    ChartIsAnimated = False
    Range("animate").Value = 0
    End
End Sub

Sub RandomButton_Click()
    Application.ScreenUpdating = False
    Range("a_inc").Value = Rnd() * 1000
    Range("b_inc").Value = Rnd() * 1000
    Range("t_inc").Value = Rnd() * 1000
    Application.ScreenUpdating = True
End Sub

Sub cbSmoothlines_Click()
    If ActiveSheet.CheckBoxes("cbSmoothLines").Value = xlOn Then
        ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).Smooth = True
    Else
        ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).Smooth = False
    End If
End Sub

