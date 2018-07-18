Attribute VB_Name = "Module1"
Sub ShowChart()
    ThisWorkbook.Sheets("Data").Calculate
    UserForm1.Show
End Sub

