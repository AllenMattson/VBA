Attribute VB_Name = "Module1"
Sub EraseRange()
    Dim UserRange As Range
    On Error GoTo Canceled
    Set UserRange = Application.InputBox _
        (Prompt:="Range to erase:", _
        Title:="Range Erase", _
        Default:=Selection.Address, _
        Type:=8)
    UserRange.Clear
    UserRange.Select
Canceled:
End Sub


