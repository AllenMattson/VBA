Attribute VB_Name = "Module1"
Sub ShowDialog()
    With UserForm1.ListBox1
        .MultiSelect = fmMultiSelectSingle
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
    End With
    UserForm1.Show
End Sub

