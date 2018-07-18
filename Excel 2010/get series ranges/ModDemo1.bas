Attribute VB_Name = "ModDemo1"
Option Explicit

Sub ShowSeriesInfo()
    Dim s As Series
    Dim Result As Variant
    Set s = ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1)
    
    Result = SERIESNAME_FROM_SERIES(s)
    If Result(1) = "Range" Then MsgBox Range(Result(2)).Address(, , , True), vbInformation, "Series Name"
    If Result(1) = "Empty" Then MsgBox "Missing", vbInformation, "Series Name"
    If Result(1) = "String" Then MsgBox Result(2), vbInformation, "Series Name"
    
    Result = XVALUES_FROM_SERIES(s)
    If Result(1) = "Range" Then MsgBox Range(Result(2)).Address(, , , True), vbInformation, "X Values"
    If Result(1) = "Array" Then MsgBox Result(2), vbInformation, "X Values"
    If Result(1) = "Empty" Then MsgBox "Missing", vbInformation, "X Values"
    
    Result = VALUES_FROM_SERIES(s)
    If Result(1) = "Range" Then MsgBox Range(Result(2)).Address(, , , True), vbInformation, "Values"
    If Result(1) = "Array" Then MsgBox Result(2), vbInformation, "Values"

    Result = BUBBLESIZE_FROM_SERIES(s)
    If Result(1) = "Range" Then MsgBox Range(Result(2)).Address(, , , True), vbInformation, "Bubble Size"
    If Result(1) = "Array" Then MsgBox Result(2), vbInformation, "Bubble Size"
    If Result(1) = "Empty" Then MsgBox "Missing", vbInformation, "Bubble Size"
End Sub

