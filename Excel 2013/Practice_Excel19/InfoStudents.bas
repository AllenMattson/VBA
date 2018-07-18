Attribute VB_Name = "InfoStudents"
Option Explicit

Sub DoStudents()
    Students.Show
End Sub

Function IsFilled()
    With Students
        If .txtSSN.Text = "" Or _
            .txtLast.Text = "" Or _
            .txtFirst.Text = "" Or _
            .cboxYear.Text = "" Or _
            .cboxMajor.Text = "" Then
            IsFilled = False
        Else
            IsFilled = True
        End If
    End With
End Function



