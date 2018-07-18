Attribute VB_Name = "Module1"
Sub ShowDialog1()
'   Make sure the RowSource property is empty
    With UserForm1
        .ListBox1.RowSource = "Sheet1!Months"
        .obMonths.Value = True
    End With
    UserForm1.Show
End Sub

