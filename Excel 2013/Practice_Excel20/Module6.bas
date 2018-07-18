Attribute VB_Name = "Module6"

Sub GetSparklineInfo()
    Dim spGrp As SparklineGroup
    Dim spCount As Long
    Dim i As Long
    
    spCount = Cells.SparklineGroups.count
    If spCount <> 0 Then
        For i = 1 To spCount
            Set spGrp = Cells.SparklineGroups(i)
            Debug.Print "Sparkline Group:" & i
            Select Case spGrp.Type
                Case 1
                    Debug.Print "Type:Line"
                Case 2
                    Debug.Print "Type:Column"
                Case 3
                    Debug.Print "Type:Win/Loss"
            End Select
            Debug.Print "Location:" & spGrp.Location.Address
            Debug.Print "Data Source: " & spGrp.SourceData
        Next i
    Else
        MsgBox "There are no sparklines in the active sheet."
    End If
End Sub


Sub CreateSparklineReport()
    Dim spGrp As SparklineGroup
    Dim sht As Worksheet
    Dim cell As Range
    Dim spLocation As Range
    
    Workbooks.Add
    Set sht = ActiveSheet
   
    EnterData sht, 3, "Month", "Sales Quota", "Sales $", "Difference"
    EnterData sht, 4, "January", "234000", "250000", "=C4-B4"
    EnterData sht, 5, "February", "211000", "180000", "=C5-B5"
    EnterData sht, 6, "March", "304000", "370000", "=C6-B6"
    Range("B4:D6").Style = "Currency"

    Columns("A:D").AutoFit
    
    Range("A1").Value = "Win/Loss"
    Set spLocation = sht.Range("B1")
    Set spGrp = spLocation.SparklineGroups _
        .Add(xlSparkColumnStacked100, "D4:D6")
    spGrp.SeriesColor.ThemeColor = 2
    spLocation.SparklineGroups.Item(1) _
        .Axes.Horizontal.Axis.Visible = True
End Sub


Sub EnterData(sht As Worksheet, rowNum As Integer, _
    ParamArray myValues() As Variant)
    
    Dim j As Integer
    Dim count As Integer
    
    count = UBound(myValues()) + 1
    j = 1
    For j = j To count
        sht.Range(Cells(rowNum, 1), Cells(rowNum, count)) = myValues()
    Next
End Sub

