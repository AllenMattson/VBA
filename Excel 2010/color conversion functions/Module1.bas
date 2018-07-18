Attribute VB_Name = "Module1"
Sub ShowColor()
    With UserForm1
        .Label1.BackColor = ActiveCell.Value
        .Show vbModeless
    End With
End Sub
