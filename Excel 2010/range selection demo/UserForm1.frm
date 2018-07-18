VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Range Selection Demo"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    RefEdit1.Text = ActiveWindow.RangeSelection.Address
End Sub

Private Sub OKButton_Click()
    Dim UserRange As Range
    Dim WorkRange As Range
    Dim Operand As Double
    Dim cell As Range
    
'   Validate range entry
    On Error Resume Next
    Set UserRange = Range(RefEdit1.Text)
    If Err <> 0 Then
        MsgBox "Invalid range selected"
        RefEdit1.SetFocus
        Exit Sub
    End If
    On Error GoTo 0

'   Validate the operand
    If Not IsNumeric(tbOperand) Then
        MsgBox "Invalid or missing operand."
        tbOperand.SetFocus
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
            RefEdit1.SetFocus
            Exit Sub
        End If
        Set WorkRange = UserRange
    Else
        On Error Resume Next
        Set WorkRange = UserRange.SpecialCells(xlCellTypeConstants, 1)
        If TypeName(WorkRange) = "Empty" Then
            MsgBox "No constant cells were found in the selection.", vbInformation
            RefEdit1.SetFocus
            Exit Sub
        End If
        On Error GoTo 0
    End If

'   Do the math!
    Operand = Val(tbOperand.Text)
    On Error Resume Next 'Ignore errors - such as Divide by 0
    Select Case True
        Case obAdd
            For Each cell In WorkRange
                cell.Value = cell.Value + Operand
            Next cell
        Case obSubtract
            For Each cell In WorkRange
                cell.Value = cell.Value - Operand
            Next cell
        Case obMultiply
            For Each cell In WorkRange
                cell.Value = cell.Value * Operand
            Next cell
        Case obDivide
            For Each cell In WorkRange
                cell.Value = cell.Value / Operand
            Next cell
    End Select
    Unload Me
End Sub


Private Sub CommandButton1_Click()
    Unload Me
End Sub

