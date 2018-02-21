Attribute VB_Name = "Module1"
Sub FormatAllCharts()
Attribute FormatAllCharts.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ChtObj As ChartObject
    For Each ChtObj In ActiveSheet.ChartObjects
      With ChtObj.Chart
        .ChartType = xlLineMarkers
        .ApplyLayout 3
        .ChartStyle = 12
        .ClearToMatchStyle
        .SetElement msoElementChartTitleAboveChart
        .SetElement msoElementLegendNone
        .SetElement msoElementPrimaryValueAxisTitleNone
        .SetElement msoElementPrimaryCategoryAxisTitleNone
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1000
      End With
    Next ChtObj
End Sub

