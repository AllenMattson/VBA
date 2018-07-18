Attribute VB_Name = "Module1"
Option Explicit

Sub ChangeSeries1Color()
    Dim MyChartObject As ChartObject
    Dim MyChart As Chart
    Dim MySeries As Series
    Dim MyChartFormat As ChartFormat
    Dim MyFillFormat As FillFormat
    Dim MyColorFormat As ColorFormat

'   Create the objects
    Set MyChartObject = ActiveSheet.ChartObjects("Chart 1")
    Set MyChart = MyChartObject.Chart
    Set MySeries = MyChart.SeriesCollection(1)
    Set MyChartFormat = MySeries.Format
    Set MyFillFormat = MyChartFormat.Fill
    Set MyColorFormat = MyFillFormat.ForeColor

'   Change the color
    MyColorFormat.RGB = vbRed
End Sub

Sub AddPresetGradient()
    Dim MyChart As Chart
    Set MyChart = ActiveSheet.ChartObjects("Chart 2").Chart
    With MyChart.SeriesCollection(1).Format.Fill
        .PresetGradient _
            Style:=msoGradientHorizontal, _
            Variant:=1, _
            PresetGradientType:=msoGradientFire
    End With
End Sub

Sub RecolorChartAndPlotArea()
'   Use theme colors for chart area and plot area
    Dim MyChart As Chart
    Set MyChart = ActiveSheet.ChartObjects("Chart 3").Chart
    With MyChart
        .ChartArea.Format.Fill.ForeColor.ObjectThemeColor = _
             msoThemeColorAccent6
        .ChartArea.Format.Fill.ForeColor.TintAndShade = 0.9
        .PlotArea.Format.Fill.ForeColor.ObjectThemeColor = _
             msoThemeColorAccent6
        .PlotArea.Format.Fill.ForeColor.TintAndShade = 0.5
    End With
End Sub

Sub UseRandomColors()
    Dim MyChart As Chart
    Set MyChart = ActiveSheet.ChartObjects("Chart 4").Chart
    With MyChart
        .ChartArea.Format.Fill.ForeColor.RGB = RandomColor
        .PlotArea.Format.Fill.ForeColor.RGB = RandomColor
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RandomColor
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RandomColor
        .Legend.Font.Color = RandomColor
        .ChartTitle.Font.Color = RandomColor
        .Axes(xlValue).MajorGridlines.Border.Color = RandomColor
        .Axes(xlValue).TickLabels.Font.Color = RandomColor
        .Axes(xlValue).Border.Color = RandomColor
        .Axes(xlCategory).TickLabels.Font.Color = RandomColor
        .Axes(xlCategory).Border.Color = RandomColor
    End With
End Sub

Function RandomColor()
    RandomColor = Application.RandBetween(0, RGB(255, 255, 255))
End Function

Sub UseRandomThemeColors()
    Dim MyChart As Chart
    Set MyChart = ActiveSheet.ChartObjects("Chart 5").Chart
    With MyChart
        .ChartArea.Format.Fill.ForeColor.ObjectThemeColor = RandomThemeColor
        .ChartArea.Format.Fill.ForeColor.TintAndShade = Application.RandBetween(-100, 100) / 100
        
        .PlotArea.Format.Fill.ForeColor.ObjectThemeColor = RandomThemeColor
        .PlotArea.Format.Fill.ForeColor.TintAndShade = Application.RandBetween(-100, 100) / 100
        
        .SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = RandomThemeColor
        .SeriesCollection(1).Format.Fill.ForeColor.TintAndShade = Application.RandBetween(-100, 100) / 100
        
        .SeriesCollection(2).Format.Fill.ForeColor.ObjectThemeColor = RandomThemeColor
        .SeriesCollection(2).Format.Fill.ForeColor.TintAndShade = Application.RandBetween(-100, 100) / 100
        
'       the following use the "old" method
        .Legend.Font.Color = RandomColor
        .ChartTitle.Font.Color = RandomColor
        .Axes(xlValue).MajorGridlines.Border.Color = RandomColor
        .Axes(xlValue).TickLabels.Font.Color = RandomColor
        .Axes(xlValue).Border.Color = RandomColor
        .Axes(xlCategory).TickLabels.Font.Color = RandomColor
        .Axes(xlCategory).Border.Color = RandomColor
    End With
End Sub

Function RandomThemeColor()
    RandomThemeColor = Application.RandBetween(1, 10)
End Function


Sub ResetAllCharts()
    Dim Cht As ChartObject
    For Each Cht In ActiveSheet.ChartObjects
        Cht.Chart.ClearToMatchStyle
    Next Cht
End Sub

