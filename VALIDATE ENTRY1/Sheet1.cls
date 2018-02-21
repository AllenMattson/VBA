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

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim VRange As Range, cell As Range
    Dim Msg As String
    Dim ValidateCode As Variant
    Set VRange = Range("InputRange")
    If Intersect(VRange, Target) Is Nothing Then Exit Sub
    For Each cell In Intersect(VRange, Target)
        ValidateCode = EntryIsValid(cell)
        If TypeName(ValidateCode) = "String" Then
            Msg = "Cell " & cell.Address(False, False) & ":"
            Msg = Msg & vbCrLf & vbCrLf & ValidateCode
            MsgBox Msg, vbCritical, "Invalid Entry"
            Application.EnableEvents = False
            cell.ClearContents
            cell.Activate
            Application.EnableEvents = True
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

