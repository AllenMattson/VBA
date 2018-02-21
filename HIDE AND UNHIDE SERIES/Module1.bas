Attribute VB_Name = "Module1"
Sub Test()
Dim SeriesFormulas(1 To 7) As String

Dim cht As Chart
Dim s As Series
Set cht = ActiveSheet.ChartObjects(1).Chart

For i = 1 To 7

Debug.Print "SeriesFormulas(" & i & ")=" & "" & cht.SeriesCollection(i).Formula; ""

Next i
End Sub
