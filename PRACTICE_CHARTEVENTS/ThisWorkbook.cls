VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' set up a reference to a class module
Dim myChart As New clsChart

Sub InitializeChart()
    
    ' you must run this procedure before event procedures
    ' written in clsChart class module can be triggered for
    ' the chart embedded in Sheet1
    
    ' connect the class module with the Excel chart object
    Set myChart.xlChart = _
        Worksheets("Sheet1").ChartObjects(1).Chart
End Sub


