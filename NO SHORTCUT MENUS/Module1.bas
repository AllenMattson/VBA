Attribute VB_Name = "Module1"
Sub SetupNoShiftF10()
    Application.OnKey "+{F10}", "NoShiftF10"
End Sub

Sub TurnOffNoShiftF10()
    Application.OnKey "+{F10}"
End Sub

Sub NoShiftF10()
    MsgBox "Nice try, but that doesn't work either."
End Sub

