Attribute VB_Name = "Module1"
Option Explicit

Sub onActionUppercase(ctl As IRibbonControl)
    Dim cell As Variant
    For Each cell In Selection
        If WorksheetFunction.IsText(cell) Then
            cell.Value = UCase(cell.Value)
        End If
    Next
End Sub

Sub onActionSelSpec(ctl As IRibbonControl)
    Select Case ctl.ID
        Case "text"
            Selection.SpecialCells(xlCellTypeConstants, 2).Select
        Case "num"
            Selection.SpecialCells(xlCellTypeConstants, 1).Select
        Case "blank"
            Selection.SpecialCells(xlCellTypeBlanks).Select
        Case "zero"
            Dim cell As Variant
            Dim myRange As Range
            Dim foundFirst As Boolean
            
            foundFirst = True
            
            Selection.SpecialCells(xlCellTypeConstants, 1).Select
                For Each cell In Selection
                    If cell.Value = 0 Then
                        If foundFirst Then
                            Set myRange = cell
                            foundFirst = False
                        End If
                        Set myRange = Application.Union(myRange, cell)
                    End If
                Next
              myRange.Select
        Case Else
           MsgBox "Missing Case statement for control id=" & ctl.ID, _
            vbOKOnly + vbExclamation, "Check your VBA Procedure"
    End Select
End Sub


Sub onActionBuiltInCmd(ctl As IRibbonControl)
    CommandBars.ExecuteMso "FileOpenRecentFile"
End Sub
