VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Chart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Chart_Activate()
    Dim msg As String
    If Worksheets("Sheet1").CheckBoxes("Check Box 1") = xlOff Then Exit Sub
    msg = "Hello " & Application.UserName & vbCrLf & vbCrLf
    msg = msg & "You are now viewing the six-month sales "
    msg = msg & "summary for Products 1-3." & vbCrLf & vbCrLf
    msg = msg & "Click on items in the chart to find out what they are."
    MsgBox msg, vbInformation, ActiveWorkbook.Name
End Sub

Private Sub Chart_Deactivate()
    Dim msg As String
    If Worksheets("Sheet1").CheckBoxes("Check Box 1") = xlOff Then Exit Sub
    msg = "Thanks for viewing the chart."
    MsgBox msg, , ActiveWorkbook.Name
End Sub

Private Sub Chart_Select(ByVal ElementID As Long, _
  ByVal Arg1 As Long, ByVal Arg2 As Long)
    Dim Id As String
    If Worksheets("Sheet1").CheckBoxes("Check Box 1") = xlOff Then Exit Sub
    Select Case ElementID
        Case xlAxis: Id = "Axis"
        Case xlAxisTitle: Id = "AxisTitle"
        Case xlChartArea: Id = "ChartArea"
        Case xlChartTitle: Id = "ChartTitle"
        Case xlCorners: Id = "Corners"
        Case xlDataLabel: Id = "DataLabel"
        Case xlDataTable: Id = "DataTable"
        Case xlDownBars: Id = "DownBars"
        Case xlDropLines: Id = "DropLines"
        Case xlErrorBars: Id = "ErrorBars"
        Case xlFloor: Id = "Floor"
        Case xlHiLoLines: Id = "HiLoLines"
        Case xlLegend: Id = "Legend"
        Case xlLegendEntry: Id = "LegendEntry"
        Case xlLegendKey: Id = "LegendKey"
        Case xlMajorGridlines: Id = "MajorGridlines"
        Case xlMinorGridlines: Id = "MinorGridlines"
        Case xlNothing: Id = "Nothing"
        Case xlPlotArea: Id = "PlotArea"
        Case xlRadarAxisLabels: Id = "RadarAxisLabels"
        Case xlSeries: Id = "Series"
        Case xlSeriesLines: Id = "SeriesLines"
        Case xlShape: Id = "Shape"
        Case xlTrendline: Id = "Trendline"
        Case xlUpBars: Id = "UpBars"
        Case xlWalls: Id = "Walls"
        Case xlXErrorBars: Id = "XErrorBars"
        Case xlYErrorBars: Id = "YErrorBars"
        Case Else:: Id = "Some unknown thing"
    End Select
    MsgBox "Selection type:" & Id & vbCrLf & Arg1 & vbCrLf & Arg2
End Sub
