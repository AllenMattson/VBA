Attribute VB_Name = "Module1"
Option Explicit

Sub ShowDialog()
'   Add some items to the ListBox
    With UserForm1.lbxItems
        .RowSource = ""
        .List = Split("January February March April May June July August September October November December")
        .ListIndex = 0
    End With
    UserForm1.Show
End Sub

