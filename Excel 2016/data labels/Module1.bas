Attribute VB_Name = "Module1"
Option Explicit

Sub DataLabelsFromRange()
Attribute DataLabelsFromRange.VB_Description = "Macro recorded 11/09/98 by John Walkenbach"
Attribute DataLabelsFromRange.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim DLRange As Range
    Dim Cht As Chart
    Dim i As Integer, Pts As Integer
    
'   Specify chart
    Set Cht = ActiveSheet.ChartObjects(1).Chart
    
'   Prompt for a range
    On Error Resume Next
    Set DLRange = Application.InputBox _
      (prompt:="Range for data labels?", Type:=8)
    If DLRange Is Nothing Then Exit Sub
    On Error GoTo 0
      
'   Add data labels
    Cht.SeriesCollection(1).ApplyDataLabels _
      Type:=xlDataLabelsShowValue, _
      AutoText:=True, _
      LegendKey:=False
      
'   Loop through the Points, and set the data labels
    Pts = Cht.SeriesCollection(1).Points.Count
    For i = 1 To Pts
        'Cht.SeriesCollection(1). _
          Points(i).DataLabel.Text = DLRange(i)
          
        'Use the statement below for formulas
        Cht.SeriesCollection(1).Points(i).DataLabel.Text = _
          "=" & "'" & DLRange.Parent.Name & "'!" & _
            DLRange(i).Address(ReferenceStyle:=xlR1C1)
    
    Next i
End Sub

Sub RemoveDataLabels()
    Dim Cht As Chart
    Dim x As Series

'   Specify chart
    Set Cht = ActiveSheet.ChartObjects(1).Chart
    Cht.SeriesCollection(1).HasDataLabels = False
End Sub
