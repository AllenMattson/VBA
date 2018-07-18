Attribute VB_Name = "Module1"
Sub FormatAllCharts()
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
        With .Axes(xlValue).MajorGridlines.Format.Line
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.25
            .DashStyle = msoLineSysDash
            .Transparency = 0
        End With
      End With
    Next ChtObj
End Sub



