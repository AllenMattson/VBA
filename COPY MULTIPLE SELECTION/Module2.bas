Attribute VB_Name = "Module2"
Option Explicit

'This version warns the user if data will be overwritten

Sub CopyMultipleSelection2()
    Dim SelAreas() As Range
    Dim PasteRange As Range
    Dim UpperLeft As Range
    Dim NumAreas As Long, i As Long
    Dim TopRow As Long, LeftCol As Long
    Dim RowOffset As Long, ColOffset As Long
    Dim NonEmptyCellCount As Long
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
'   Store the areas as separate Range objects
    NumAreas = Selection.Areas.Count
    ReDim SelAreas(1 To NumAreas)
    For i = 1 To NumAreas
        Set SelAreas(i) = Selection.Areas(i)
    Next
    
'   Determine the upper left cell in the multiple selection
    TopRow = ActiveSheet.Rows.Count
    LeftCol = ActiveSheet.Columns.Count
    For i = 1 To NumAreas
        If SelAreas(i).Row < TopRow Then TopRow = SelAreas(i).Row
        If SelAreas(i).Column < LeftCol Then LeftCol = SelAreas(i).Column
    Next
    Set UpperLeft = Cells(TopRow, LeftCol)
    
'   Get the paste address
    On Error Resume Next
    Set PasteRange = Application.InputBox _
      (Prompt:="Specify the upper left cell for the paste range:", _
      Title:="Copy Mutliple Selection", _
      Type:=8)
    On Error GoTo 0
'   Exit if canceled
    If TypeName(PasteRange) <> "Range" Then Exit Sub

'   Make sure only the upper left cell is used
    Set PasteRange = PasteRange.Range("A1")
    
'   Check paste range for existing data
    NonEmptyCellCount = 0
    For i = 1 To NumAreas
        RowOffset = SelAreas(i).Row - TopRow
        ColOffset = SelAreas(i).Column - LeftCol
        NonEmptyCellCount = NonEmptyCellCount + _
            Application.CountA(Range(PasteRange.Offset(RowOffset, ColOffset), _
            PasteRange.Offset(RowOffset + SelAreas(i).Rows.Count - 1, _
            ColOffset + SelAreas(i).Columns.Count - 1)))
    Next i
    
'   If paste range is not empty, warn user
    If NonEmptyCellCount <> 0 Then _
        If MsgBox("Overwrite existing data?", vbQuestion + vbYesNo, _
        "Copy Multiple Selection") <> vbYes Then Exit Sub

'   Copy and paste each area
    For i = 1 To NumAreas
        RowOffset = SelAreas(i).Row - TopRow
        ColOffset = SelAreas(i).Column - LeftCol
        SelAreas(i).Copy PasteRange.Offset(RowOffset, ColOffset)
    Next i
End Sub

