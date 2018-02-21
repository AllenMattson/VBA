VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Dim strPivotPage As String
    Dim r As Integer

    strPivotPage = Target.PivotFields("ProductName"). _
        CurrentPage.Value

    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects(1).Activate
        ActiveChart.ChartTitle.Text = strPivotPage

        r = ActiveSheet.UsedRange.Rows.Count + 3

        With Range("B" & r)
           ActiveSheet.ChartObjects(1).Top = .Top
    End With
    End If
End Sub

