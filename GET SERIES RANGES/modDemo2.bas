Attribute VB_Name = "modDemo2"
Option Explicit

Sub ContractAllSeries()
    Dim s As Series
    Dim Result As Variant
    Dim DRange As Range
    For Each s In ActiveSheet.ChartObjects(1).Chart.SeriesCollection
        Result = XVALUES_FROM_SERIES(s)
        If Result(1) = "Range" Then
            Set DRange = Range(Result(2))
            If DRange.Rows.Count > 1 Then
                Set DRange = DRange.Resize(DRange.Rows.Count - 1)
                s.XValues = DRange
            End If
        End If
        Result = VALUES_FROM_SERIES(s)
        If Result(1) = "Range" Then
            Set DRange = Range(Result(2))
            If DRange.Rows.Count > 1 Then
                Set DRange = DRange.Resize(DRange.Rows.Count - 1)
                s.Values = DRange
            End If
        End If
    Next s
End Sub

Sub ExpandAllSeries()
    Dim s As Series
    Dim Result As Variant
    Dim DRange As Range
    For Each s In ActiveSheet.ChartObjects(1).Chart.SeriesCollection
        Result = XVALUES_FROM_SERIES(s)
        If Result(1) = "Range" Then
            Set DRange = Range(Result(2))
            Set DRange = DRange.Resize(DRange.Rows.Count + 1)
            s.XValues = DRange
        End If
    
        Result = VALUES_FROM_SERIES(s)
        If Result(1) = "Range" Then
            Set DRange = Range(Result(2))
            Set DRange = DRange.Resize(DRange.Rows.Count + 1)
            s.Values = DRange
        End If
    Next s
End Sub

