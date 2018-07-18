Attribute VB_Name = "Examples"
Option Explicit

Sub SelectDown()
Attribute SelectDown.VB_ProcData.VB_Invoke_Func = " \n14"
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
End Sub


Sub SelectUp()
Attribute SelectUp.VB_ProcData.VB_Invoke_Func = " \n14"
    Range(ActiveCell, ActiveCell.End(xlUp)).Select
End Sub


Sub SelectToRight()
Attribute SelectToRight.VB_ProcData.VB_Invoke_Func = " \n14"
    Range(ActiveCell, ActiveCell.End(xlToRight)).Select
End Sub


Sub SelectToLeft()
Attribute SelectToLeft.VB_ProcData.VB_Invoke_Func = " \n14"
    Range(ActiveCell, ActiveCell.End(xlToLeft)).Select
End Sub


Sub SelectCurrentRegion()
Attribute SelectCurrentRegion.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.CurrentRegion.Select
End Sub


Sub SelectActiveArea()
Attribute SelectActiveArea.VB_ProcData.VB_Invoke_Func = " \n14"
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).Select
End Sub


Sub SelectActiveColumn()
Attribute SelectActiveColumn.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim TopCell As Range
    Dim BottomCell As Range
    
    If IsEmpty(ActiveCell) Then Exit Sub
'   ignore error if activecell is in Row 1
    On Error Resume Next
    If IsEmpty(ActiveCell.Offset(-1, 0)) Then Set TopCell = ActiveCell Else Set TopCell = ActiveCell.End(xlUp)
    If IsEmpty(ActiveCell.Offset(1, 0)) Then Set BottomCell = ActiveCell Else Set BottomCell = ActiveCell.End(xlDown)
    Range(TopCell, BottomCell).Select

End Sub


Sub SelectActiveRow()
Attribute SelectActiveRow.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim LeftCell As Range
    Dim RightCell As Range
    
    If IsEmpty(ActiveCell) Then Exit Sub
'   ignore error if activecell is in Column A
    On Error Resume Next
    If IsEmpty(ActiveCell.Offset(0, -1)) Then Set LeftCell = ActiveCell Else Set LeftCell = ActiveCell.End(xlToLeft)
    If IsEmpty(ActiveCell.Offset(0, 1)) Then Set RightCell = ActiveCell Else Set RightCell = ActiveCell.End(xlToRight)
    Range(LeftCell, RightCell).Select
End Sub


Sub SelectEntireColumn()
Attribute SelectEntireColumn.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.EntireColumn.Select
End Sub


Sub SelectEntireRow()
Attribute SelectEntireRow.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.EntireRow.Select
End Sub


Sub SelectEntireSheet()
Attribute SelectEntireSheet.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.Select
End Sub


Sub ActivateNextBlankDown()
Attribute ActivateNextBlankDown.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.Offset(1, 0).Select
    Do While Not IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub


Sub ActivateNextBlankToRight()
Attribute ActivateNextBlankToRight.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.Offset(0, 1).Select
    Do While Not IsEmpty(ActiveCell)
        ActiveCell.Offset(0, 1).Select
    Loop
End Sub


Sub SelectFirstToLastInRow()
Attribute SelectFirstToLastInRow.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim LeftCell As Range
    Dim RightCell As Range
    
    Set LeftCell = Cells(ActiveCell.Row, 1)
    Set RightCell = Cells(ActiveCell.Row, 256)

    If IsEmpty(LeftCell) Then Set LeftCell = LeftCell.End(xlToRight)
    If IsEmpty(RightCell) Then Set RightCell = RightCell.End(xlToLeft)
    If LeftCell.Column = 256 And RightCell.Column = 1 Then ActiveCell.Select Else Range(LeftCell, RightCell).Select
End Sub


Sub SelectFirstToLastInColumn()
Attribute SelectFirstToLastInColumn.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim TopCell As Range
    Dim BottomCell As Range
    
    Set TopCell = Cells(1, ActiveCell.Column)
    Set BottomCell = Cells(16384, ActiveCell.Column)

    If IsEmpty(TopCell) Then Set TopCell = TopCell.End(xlDown)
    If IsEmpty(BottomCell) Then Set BottomCell = BottomCell.End(xlUp)
    If TopCell.Row = 16384 And BottomCell.Row = 1 Then ActiveCell.Select Else Range(TopCell, BottomCell).Select
End Sub


