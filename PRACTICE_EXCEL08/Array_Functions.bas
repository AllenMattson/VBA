Attribute VB_Name = "Array_Functions"
Option Explicit

Option Base 1

Sub CarInfo()
  Dim auto As Variant
  auto = Array("Ford", "Black", "1999")
  MsgBox auto(2) & " " & auto(1) & ", " & auto(3)
  auto(2) = "4-door"
  MsgBox auto(2) & " " & auto(1) & ", " & auto(3)
End Sub

Sub ColumnHeads()
      Dim heading As Variant
      Dim cell As Range
      Dim i As Integer
      i = 1
      heading = Array("First Name", "Last Name", "Position", "Salary")
      Workbooks.Add
    
      For Each cell In Range("A1:D1")
        cell.Formula = heading(i)
      i = i + 1
      Next
    
      Columns("A:D").Select
      Selection.Columns.AutoFit
      Range("A1").Select
End Sub

Function ReplaceIllegalChars(strInput As String) As String
  Dim illegal As Variant
  Dim i As Integer
    
  illegal = Array("~", "!", "?", "<", ">", "[", "]", ":", "|", _
        "*", "/")
    
  For i = LBound(illegal) To UBound(illegal)
      Do While InStr(strInput, illegal(i))
          Mid(strInput, InStr(strInput, illegal(i)), 1) = "_"
      Loop
  Next i
    
  ReplaceIllegalChars = strInput
End Function

