Attribute VB_Name = "Module1"
Option Explicit

Sub UpdateChart()
    Dim ChtObj As ChartObject
    Dim UserRow As Long
    Set ChtObj = ActiveSheet.ChartObjects(1)
    UserRow = ActiveCell.Row
    If UserRow < 4 Or IsEmpty(Cells(UserRow, 1)) Then
        ChtObj.Visible = False
    Else
        ChtObj.Chart.SeriesCollection(1).Values = _
           Range(Cells(UserRow, 2), Cells(UserRow, 6))
        ChtObj.Chart.ChartTitle.Text = Cells(UserRow, 1).Text
        ChtObj.Visible = True
    End If
End Sub

Sub test()
Dim sc As Series

Set sc = ActiveChart.SeriesCollection(1)
MsgBox sc.Name
End Sub
