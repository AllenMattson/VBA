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
    Application.ShowChartTipNames = False
    Application.ShowChartTipValues = False
End Sub

Private Sub Chart_Deactivate()
    Application.ShowChartTipNames = True
    Application.ShowChartTipValues = True
End Sub

Private Sub Chart_MouseMove(ByVal Button As Long, ByVal Shift As Long, _
  ByVal X As Long, ByVal Y As Long)
    Dim ElementId As Long
    Dim arg1 As Long, arg2 As Long
    On Error Resume Next
    ActiveChart.GetChartElement X, Y, ElementId, arg1, arg2
    If ElementId = xlSeries Then
        ActiveChart.Shapes(1).Visible = msoCTrue
        ActiveChart.Shapes(1).TextFrame.Characters.Text = _
          Sheets("Sheet1").Range("Comments").Offset(arg2, arg1)
    Else
        ActiveChart.Shapes(1).Visible = msoFalse
    End If
End Sub
