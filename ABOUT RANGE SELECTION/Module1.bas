Attribute VB_Name = "Module1"
Option Explicit

Sub RangeDescription()
    Dim NumCols As Integer
    Dim NumRows As Long
    Dim NumBlocks As Integer
    Dim NumCells As Double
    Dim NumAreas As Integer
    Dim SelType As String
    Dim FirstAreaType As String
    Dim CurrentType  As String
    Dim WhatSelected As String
    Dim UnionRange As Range
    Dim Area As Range
    Dim Msg As String

'   Quit if a range is not selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range."
        Exit Sub
    End If

'   Initialize counters
    NumCols = 0
    NumRows = 0
    NumBlocks = 0
    NumCells = 0

'   Determine number of areas in selection
    NumAreas = Selection.Areas.Count
    If NumAreas = 1 Then
        SelType = "Single Selection"
    Else
        SelType = "Multiple Selection"
    End If

    FirstAreaType = AreaType(Selection.Areas(1))
    WhatSelected = FirstAreaType

'   Build the union of all areas to avoid double-counting
    Set UnionRange = Selection.Areas(1)
    
    For Each Area In Selection.Areas
        CurrentType = AreaType(Area)

'       Count blocks before they're combined in the union
        If CurrentType = "Block" Then NumBlocks = NumBlocks + 1
        Set UnionRange = Union(UnionRange, Area)

'       Change label if multiple selection is "mixed"
        If CurrentType <> FirstAreaType Then WhatSelected = "Mixed"
    Next Area

'   Loop through each area in the Union range
    For Each Area In UnionRange.Areas
        Select Case AreaType(Area)
            Case "Row"
                NumRows = NumRows + Area.Rows.Count
            Case "Column"
                NumCols = NumCols + Area.Columns.Count
            Case "Worksheet"
                NumCols = NumCols + Area.Columns.Count
                NumRows = NumRows + Area.Rows.Count
            Case "Block"
'           Blocks already counted in original selection above
        End Select
     Next Area

'   Count number of non-overlapping cells
    NumCells = UnionRange.CountLarge

    Msg = "Selection Type:" & vbTab & WhatSelected & vbCrLf
    Msg = Msg & "No. of Areas:" & vbTab & NumAreas & vbCrLf
    Msg = Msg & "Full Columns: " & vbTab & NumCols & vbCrLf
    Msg = Msg & "Full Rows: " & vbTab & NumRows & vbCrLf
    Msg = Msg & "Cell Blocks:" & vbTab & NumBlocks & vbCrLf
    Msg = Msg & "Total Cells: " & vbTab & Format(NumCells, "#,###")
    MsgBox Msg, vbInformation, SelType
End Sub

Private Function AreaType(RangeArea As Range) As String
'   Returns the type of a range in an area
    Select Case True
        Case RangeArea.Cells.CountLarge = 1
            AreaType = "Cell"
        Case RangeArea.CountLarge = Cells.CountLarge
            AreaType = "Worksheet"
        Case RangeArea.Rows.Count = Cells.Rows.Count
            AreaType = "Column"
        Case RangeArea.Columns.Count = Cells.Columns.Count
            AreaType = "Row"
        Case Else
            AreaType = "Block"
    End Select
End Function

