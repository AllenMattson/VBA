Attribute VB_Name = "Module1"
Public AnimationInProgress As Boolean

Sub AnimateChart()
    Dim StartVal As Long, r As Long
    If AnimationInProgress Then
        AnimationInProgress = False
        End
    End If
    AnimationInProgress = True
    StartVal = Range("StartDay")
    For r = StartVal To 5219 - Range("NumDays") Step Range("Increment")
        Range("StartDay") = r
        DoEvents
    Next r
    AnimationInProgress = False
End Sub

