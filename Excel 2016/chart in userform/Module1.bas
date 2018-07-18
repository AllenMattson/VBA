Attribute VB_Name = "Module1"
Option Explicit

Sub ShowChart()
    Dim UserRow As Long
    UserRow = ActiveCell.Row
    If UserRow < 2 Or IsEmpty(Cells(UserRow, 1)) Then
        MsgBox "Move the cell pointer to a row that contains data."
        Exit Sub
    End If
    CreateChart (UserRow)
    UserForm1.Show
End Sub

Sub CreateChart(r)
    Dim TempChart As Chart
    Dim CatTitles As Range
    Dim SrcRange As Range, SourceData As Range
    Dim fname As String
    
    Set CatTitles = ActiveSheet.Range("A2:F2")
    Set SrcRange = ActiveSheet.Range(Cells(r, 1), Cells(r, 6))
    Set SourceData = Union(CatTitles, SrcRange)
    
'   Add a chart
    Application.ScreenUpdating = False
    Set TempChart = ActiveSheet.Shapes.AddChart2.Chart
    TempChart.SetSourceData Source:=SourceData
    
'   Fix it up
    With TempChart
        .ChartStyle = 25
        .ChartType = xlColumnClustered
        .SetSourceData Source:=SourceData, PlotBy:=xlRows
        .HasLegend = False
        .PlotArea.Interior.ColorIndex = xlNone
        .Axes(xlValue).MajorGridlines.Delete
        .ApplyDataLabels Type:=xlDataLabelsShowValue, LegendKey:=False
        .Axes(xlValue).MaximumScale = 0.6
        .ChartArea.Format.Line.Visible = False
    End With

'   Adjust the ChartObject's size size
    With ActiveSheet.ChartObjects(1)
        .Width = 300
        .Height = 200
        .Activate
    End With

'   Save chart as GIF
    fname = Application.DefaultFilePath & Application.PathSeparator & "temp.gif"
    TempChart.Export Filename:=fname, filterName:="GIF"
    ActiveSheet.ChartObjects(1).Delete
    Application.ScreenUpdating = True
End Sub

