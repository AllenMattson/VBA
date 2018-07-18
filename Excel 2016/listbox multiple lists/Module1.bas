Attribute VB_Name = "Module1"
Option Explicit

Sub ShowDialog1()
'   Make sure the RowSource property is empty
    With UserForm1
        .ListBox1.RowSource = "Sheet1!Months"
        .optMonths.Value = True
    End With
    UserForm1.Show
End Sub

