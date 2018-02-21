Attribute VB_Name = "Module1"
Dim myClassModule As New Class1

Sub ConnectChart()
    Application.ShowChartTipNames = False
    Application.ShowChartTipValues = False
    Set myClassModule.MyChart = ActiveSheet.ChartObjects(1).Chart
    ActiveSheet.ChartObjects(1).Activate
End Sub

Sub DisconnectChart()
    Application.ShowChartTipNames = True
    Application.ShowChartTipValues = True
    Set myClassModule = Nothing
    Range("A1").Select
End Sub

