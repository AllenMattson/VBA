Attribute VB_Name = "Module1"
Option Explicit

Sub SparklineReport()
    Dim sg As SparklineGroup
    Dim sl As Sparkline
    Dim SGType As String
    Dim SLSheet As Worksheet
    Dim i As Long, j As Long, r As Long
    
    If Cells.SparklineGroups.Count = 0 Then
        MsgBox "No sparklines were found on the active sheet."
        Exit Sub
    End If
    
    Set SLSheet = ActiveSheet
'   Insert new worksheet for the report
    Worksheets.Add
    
'   Headings
    With Range("A1")
        .Value = "Sparkline Report: " & SLSheet.Name & " in " & SLSheet.Parent.Name
        .Font.Bold = True
        .Font.Size = 16
    End With
    With Range("A3:F3")
        .Value = Array("Group #", "Sparkline Grp Range", _
           "# in Group", "Type", "Sparkline #", "Source Range")
        .Font.Bold = True
    End With
    r = 4
    
    'Loop through each sparkline group
    For i = 1 To SLSheet.Cells.SparklineGroups.Count
        Set sg = SLSheet.Cells.SparklineGroups(i)
        Select Case sg.Type
            Case 1: SGType = "Line"
            Case 2: SGType = "Column"
            Case 3: SGType = "Win/Loss"
        End Select
        ' Loop through each sparkline in the group
        For j = 1 To sg.Count
            Set sl = sg.Item(j)
            Cells(r, 1) = i 'Group #
            Cells(r, 2) = sg.Location.Address
            Cells(r, 3) = sg.Count
            Cells(r, 4) = SGType
            Cells(r, 5) = j 'Sparkline # within Group
            Cells(r, 6) = sl.SourceData
            r = r + 1
        Next j
        r = r + 1
    Next i
End Sub

Sub ListSparklineGroups()
    Dim sg As SparklineGroup
    Dim i As Long
    For i = 1 To Cells.SparklineGroups.Count
        Set sg = Cells.SparklineGroups(i)
        MsgBox sg.Location.Address
    Next i
End Sub

