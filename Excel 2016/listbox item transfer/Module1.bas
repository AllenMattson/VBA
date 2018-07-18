Attribute VB_Name = "Module1"
Option Explicit

Sub ShowDialog()
    
    Dim i As Long
    
    UserForm1.lbxFrom.RowSource = ""
'   Add some items to the ListBox
    With UserForm1.lbxFrom
        .RowSource = ""
        For i = 1 To 12
            .AddItem Format(DateSerial(2000, i, 1), "mmmm")
        Next i
    End With
    'Select the first item
    UserForm1.lbxFrom.ListIndex = 0
    UserForm1.Show
End Sub

