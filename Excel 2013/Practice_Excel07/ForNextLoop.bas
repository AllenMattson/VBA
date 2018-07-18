Attribute VB_Name = "ForNextLoop"
Option Explicit

Sub DeleteZeroRows()
  Dim totalR As Integer
  Dim r As Integer
    
  Range("A1").CurrentRegion.Select
  totalR = Selection.Rows.Count
  Range("B2").Select
    
  For r = 1 To totalR - 1
    If ActiveCell = 0 Then
          Selection.EntireRow.Delete
          totalR = totalR - 1
    Else
          ActiveCell.Offset(1, 0).Select
    End If
  Next r
End Sub

