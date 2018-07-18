Attribute VB_Name = "Module1"
Option Explicit

Sub EraseRange()
    Dim UserRange As Range
    On Error Resume Next
        Set UserRange = Application.InputBox _
            (Prompt:="Select the range to erase:", _
            Title:="Range Erase", _
            Default:=Selection.Address, _
            Type:=8)
    On Error GoTo 0
    If Not UserRange Is Nothing Then
        UserRange.ClearContents
        UserRange.Select
    End If
End Sub


Sub GetValue2()
    Dim Monthly As Variant
    Monthly = Application.InputBox _
        (Prompt:="Enter your monthly salary:", _
         Type:=1)
    If Monthly <> False Then
        MsgBox "Annualized: " & Monthly * 12
    End If
End Sub
