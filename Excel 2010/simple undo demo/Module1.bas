Attribute VB_Name = "Module1"
Option Explicit
  'Custom data type for undoing
    Type SaveRange
        Val As Variant
        Addr As String
    End Type
    
'   Stores info about current selection
    Public OldWorkbook As Workbook
    Public OldSheet As Worksheet
    Public OldSelection() As SaveRange


Sub ZeroRange()
'   Inserts zero into all selected cells
    Dim i As Long
    Dim cell As Range
    
'   Abort if a range isn't selected
    If TypeName(Selection) <> "Range" Then Exit Sub
    
'   Abort if too many cells are selected
    If Selection.Count > 50000 Then
        MsgBox "You selected too many cells.", vbCritical
        Exit Sub
    End If

'   The next block of statements
'   Save the current values for undoing
    ReDim OldSelection(Selection.Count)
    Set OldWorkbook = ActiveWorkbook
    Set OldSheet = ActiveSheet
    i = 0
    For Each cell In Selection
        i = i + 1
        OldSelection(i).Addr = cell.Address
        OldSelection(i).Val = cell.Formula
    Next cell
            
'   Insert 0 into current selection
    Application.ScreenUpdating = False
    Selection.Value = 0
    
'   Specify the Undo Sub
    Application.OnUndo "Undo the ZeroRange macro", "UndoZero"
End Sub


Sub UndoZero()
'   Undoes the effect of the ZeroRange sub
    Dim i As Long
    
'   Tell user if a problem occurs
    On Error GoTo Problem

    Application.ScreenUpdating = False
    
'   Make sure the correct workbook and sheet are active
    OldWorkbook.Activate
    OldSheet.Activate
    
'   Restore the saved information
    For i = 1 To UBound(OldSelection)
        Range(OldSelection(i).Addr).Formula = OldSelection(i).Val
    Next i
    Exit Sub

'   Error handler
Problem:
    MsgBox "Can't undo", vbCritical
End Sub
