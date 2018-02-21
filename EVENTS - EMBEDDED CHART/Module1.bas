Attribute VB_Name = "Module1"
Dim MyChart As New clsChart

Sub CheckBox1_Click()
    If Worksheets("Sheet1").CheckBoxes("Check Box 1") = xlOn Then
        Set MyChart.clsChart = ActiveSheet.ChartObjects(1).Chart
    Else
        Set MyChart.clsChart = Nothing
    End If
End Sub

