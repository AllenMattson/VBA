Attribute VB_Name = "Module1"
Option Explicit

Sub CopyMultipleSelection()
    Dim SelAreas() As Range
    Dim PasteRange As Range
    Dim UpperLeft As Range
    Dim NumAreas As Integer, i As Integer
    Dim TopRow As Long, LeftCol As Integer
    Dim RowOffset As Long, ColOffset As Integer
    
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
    
'   Copy and paste each area
    For i = 1 To NumAreas
        RowOffset = SelAreas(i).Row - TopRow
        ColOffset = SelAreas(i).Column - LeftCol
        SelAreas(i).Copy PasteRange.Offset(RowOffset, ColOffset)
    Next i
End Sub

