VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub sWorksheet_Change(ByVal Target As Excel.Range)
    Static AbortProc As Boolean
    Dim VRange As Range, cell As Range
    Dim Msg As String
    Dim ValidateCode As Variant
    If AbortProc Then Exit Sub
    Set VRange = Range("InputRange")
    For Each cell In Target
        If Union(cell, VRange).Address = VRange.Address Then
            ValidateCode = EntryIsValid(cell)
            If ValidateCode = True Then
                AbortProc = False
                Exit Sub
            Else
                Msg = "Cell " & cell.Address(False, False) & ":"
                Msg = Msg & vbCrLf & vbCrLf & ValidateCode
                MsgBox Msg, vbCritical, "Invalid Entry"
                AbortProc = True
                cell.ClearContents
                cell.Activate
                AbortProc = False
            End If
        End If
    Next cell
    AbortProc = False
End Sub

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Static AbortProc As Boolean
    Dim VRange As Range, cell As Range
    Dim Msg As String
    Dim ValidateCode As Variant
    If AbortProc Then Exit Sub
    Set VRange = Range("InputRange")
    If Intersect(VRange, Target) Is Nothing Then Exit Sub
    For Each cell In Intersect(VRange, Target)
        ValidateCode = EntryIsValid(cell)
        If TypeName(ValidateCode) = "String" Then
            Msg = "Cell " & cell.Address(False, False) & ":"
            Msg = Msg & vbCrLf & vbCrLf & ValidateCode
            MsgBox Msg, vbCritical, "Invalid Entry"
            AbortProc = True
            cell.ClearContents
            cell.Activate
            AbortProc = False
        End If
    Next cell
End Sub



Private Function EntryIsValid(cell) As Variant
'   Returns True if cell is an integer between 1 and 12
'   Otherwise it returns a string that describes the problem
    
'   Numeric
    If Not WorksheetFunction.IsNumber(cell) Then
        EntryIsValid = "Non-numeric entry."
        Exit Function
    End If
    
'   Integer?
    If CInt(cell) <> cell Then
        EntryIsValid = "Integer required."
        Exit Function
    End If
            
'   Between 1 and 12?
    If cell < 1 Or cell > 12 Then
        EntryIsValid = "Valid values are between 1 and 12."
        Exit Function
    End If

'   It passed all the tests
    EntryIsValid = True
End Function

