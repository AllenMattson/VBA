VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents MyChart As Chart
Attribute MyChart.VB_VarHelpID = -1


Private Sub MyChart_MouseMove(ByVal Button As Long, ByVal Shift As Long, _
  ByVal x As Long, ByVal y As Long)
    Dim ElementId As Long
    Dim arg1 As Long, arg2 As Long
    Dim NewText As String
    'On Error Resume Next
    MyChart.GetChartElement x, y, ElementId, arg1, arg2
    If ElementId = xlSeries Then
        MyChart.Shapes(1).TextFrame.Characters.Text = _
          Sheets("Sheet1").Range("Comments").Offset(arg2, arg1)
        MyChart.Shapes(1).Visible = msoCTrue
    Else
        MyChart.Shapes(1).Visible = msoFalse
    End If

End Sub

