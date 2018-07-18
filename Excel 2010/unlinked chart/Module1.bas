Attribute VB_Name = "Module1"
Sub CreateUnlinkedChart()
    Dim MyChart As Chart
    Set MyChart = ActiveSheet.Shapes.AddChart.Chart
    With MyChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Sales"
        .SeriesCollection(1).XValues = Array("Jan", "Feb", "Mar")
        .SeriesCollection(1).Values = Array(125, 165, 189)
        .ChartType = xlColumnClustered
        .SetElement msoElementLegendNone
    End With
End Sub

Sub ConvertChartToPicture()
    Dim Cht As Chart
    If ActiveChart Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) = "Chart" Then Exit Sub
    Set Cht = ActiveChart
    Cht.CopyPicture Appearance:=xlPrinter, _
      Size:=xlScreen, Format:=xlPicture
    ActiveWindow.RangeSelection.Select
    ActiveSheet.Paste
End Sub


