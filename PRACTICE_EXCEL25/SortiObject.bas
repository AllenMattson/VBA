Attribute VB_Name = "SortiObject"
Option Explicit

Sub SortData()
    Range("A2").Select
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A2:A6"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal
        .SetRange Range("A1:B6")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    MsgBox "Data has been sorted.", vbInformation
End Sub



