Attribute VB_Name = "Module1"
Option Explicit

Sub ShowDialog()
        
    Dim i As Long
    
    With UserForm1.ListBox1
        .MultiSelect = fmMultiSelectSingle
        .RowSource = ""
        
        For i = 1 To 12
            UserForm1.ListBox1.AddItem Format(DateSerial(Year(Now), i, 1), "mmmm")
        Next i
    End With
    UserForm1.Show
End Sub

