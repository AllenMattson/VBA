Attribute VB_Name = "Module1"
Option Explicit

Dim SummaryChart As New EmbChartClass

Sub CheckBox1_Click()
    If Worksheets("Main").CheckBoxes("Check Box 1") = xlOn Then
        'Enable chart events
        Range("A1").Select
        Set SummaryChart.myChartClass = _
          Worksheets(1).ChartObjects(1).Chart
    Else
        'Disable chart events
        Set SummaryChart.myChartClass = Nothing
        Range("A1").Select
    End If
End Sub

Sub ReturnToMain()
'   Called by worksheet button
    Sheets("Main").Activate
    ActiveWindow.RangeSelection.Select
End Sub

