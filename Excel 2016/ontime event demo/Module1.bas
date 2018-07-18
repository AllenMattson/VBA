Attribute VB_Name = "Module1"
    Dim NextTick As Date

Sub UpdateClock()
'   Updates cell A1 with the current time
    ThisWorkbook.Sheets(1).Range("A1") = Time
'   Set up the next event five seconds from now
    NextTick = Now + TimeValue("00:00:05")
    Application.OnTime NextTick, "UpdateClock"
End Sub

Sub StopClock()
'   Cancels the OnTime event (stops the clock)
    On Error Resume Next
    Application.OnTime NextTick, "UpdateClock", , False
End Sub

