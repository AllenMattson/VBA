Attribute VB_Name = "Module1"
Option Explicit

Sub ShowDialog1()
    UserForm1.ListBox1.RowSource = "Sheet1!A1:A12"
    UserForm1.Show
End Sub

Sub ShowDialog2()
    Dim i As Integer
'   Make sure the RowSource property is empty
    UserForm1.ListBox1.RowSource = ""
    
'   Add some items to the ListBox
    For i = 1 To 12
        UserForm1.ListBox1.AddItem "Item #" & i
    Next i
    UserForm1.Show
End Sub

Sub ShowDialog3()
    Dim row As Integer
'   Make sure the RowSource property is empty
    UserForm1.ListBox1.RowSource = ""
    
'   Add some items to the ListBox
    For row = 1 To 12
        UserForm1.ListBox1.AddItem Sheets("Sheet1").Cells(row, 1)
    Next row
    
'   Simpler method, no loop
'   UserForm1.ListBox1.List = Application.Transpose(Sheets("Sheet1").Range("A1:A12"))
    UserForm1.Show
End Sub

