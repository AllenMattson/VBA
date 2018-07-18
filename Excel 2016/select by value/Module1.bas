Attribute VB_Name = "Module1"
Option Explicit

Sub SelectByValue()
    Dim Cell As Object
    Dim FoundCells As Range
    Dim WorkRange As Range
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
'   Check all or selection?
    If Selection.CountLarge = 1 Then
        Set WorkRange = ActiveSheet.UsedRange
    Else
       Set WorkRange = Application.Intersect(Selection, ActiveSheet.UsedRange)
    End If
    
'   Reduce the search to numeric cells only
    On Error Resume Next
    Set WorkRange = WorkRange.SpecialCells(xlConstants, xlNumbers)
    If WorkRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
'   Loop through each cell, add to the FoundCells range if it qualifies
    For Each Cell In WorkRange
        If Cell.Value < 0 Then
            If FoundCells Is Nothing Then
                Set FoundCells = Cell
            Else
                Set FoundCells = Application.Union(FoundCells, Cell)
            End If
        End If
    Next Cell

'   Show message, or select the cells
    If FoundCells Is Nothing Then
        MsgBox "No cells qualify."
    Else
        FoundCells.Select
        MsgBox "Selected " & FoundCells.Count & " cells."
    End If
End Sub

