VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRangeSelect 
   Caption         =   "Range Selection Demo"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   OleObjectBlob   =   "ufRangeSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufRangeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    Me.refRange.Text = ActiveWindow.RangeSelection.Address
End Sub

Private Sub cmdOK_Click()
    Dim UserRange As Range
    Dim WorkRange As Range
    Dim Operand As Double
    Dim cell As Range
    
'   Validate range entry
    On Error Resume Next
    Set UserRange = Range(Me.refRange.Text)
    If Err <> 0 Then
        MsgBox "Invalid range selected"
        Me.refRange.SetFocus
        Exit Sub
    End If
    On Error GoTo 0

'   Validate the operand
    If Not IsNumeric(Me.tbxOperand.Text) Then
        MsgBox "Invalid or missing operand."
        tbxOperand.SetFocus
        Exit Sub
    End If
    
'   Create a range that consists of numeric constants
'   Check for single cell selection
'   If so, skip SpecialCells since it will work with entire sheet!
    If UserRange.Count = 1 Then
        If UserRange.HasFormula Or _
           Application.WorksheetFunction.IsText(UserRange) Or _
           IsEmpty(UserRange) Then
            MsgBox "No constant cells were found in the selection.", vbInformation
            Me.refRange.SetFocus
            Exit Sub
        End If
        Set WorkRange = UserRange
    Else
        On Error Resume Next
        Set WorkRange = UserRange.SpecialCells(xlCellTypeConstants, 1)
        If TypeName(WorkRange) = "Empty" Then
            MsgBox "No constant cells were found in the selection.", vbInformation
            Me.refRange.SetFocus
            Exit Sub
        End If
        On Error GoTo 0
    End If

'   Do the math!
    Operand = Val(Me.tbxOperand.Text)
    On Error Resume Next 'Ignore errors - such as Divide by 0
    Select Case True
        Case Me.optAdd.Value
            For Each cell In WorkRange
                cell.Value = cell.Value + Operand
            Next cell
        Case Me.optSubtract.Value
            For Each cell In WorkRange
                cell.Value = cell.Value - Operand
            Next cell
        Case Me.optMultiply.Value
            For Each cell In WorkRange
                cell.Value = cell.Value * Operand
            Next cell
        Case Me.optDivide.Value
            For Each cell In WorkRange
                cell.Value = cell.Value / Operand
            Next cell
    End Select
    Unload Me
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

