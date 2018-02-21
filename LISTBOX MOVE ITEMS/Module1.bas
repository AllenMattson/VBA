Attribute VB_Name = "Module1"
Sub ShowDialog()
'   Add some items to the ListBox
    With UserForm1.ListBox1
        .RowSource = ""
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
        .ListIndex = 0
    End With
    UserForm1.Show
End Sub

