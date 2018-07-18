Attribute VB_Name = "SpeechObject"
Option Explicit

Sub ReadNamesWithHighScores()
    Dim v As Integer
    Dim cell As Variant

    v = InputBox("Enter the minimum expected score:", _
        "Approved Minimum")

    For Each cell In ActiveSheet.UsedRange.Columns("B").Cells
        If IsNumeric(cell.Value) And cell.Value >= v Then
            Application.Speech.Speak "Congratulations " & _
            cell.Offset(0, -1).Text
            Application.Speech.Speak " your score is " & cell.Text
        End If
    Next
End Sub



