Attribute VB_Name = "Module1"
Sub DupeRows()
  Dim cell As Range
' 1st cell with number of tickets
  Set cell = Range("B2")
  Do While Not IsEmpty(cell)
    If cell > 1 Then
      Range(cell.Offset(1, 0), cell.Offset(cell.Value - 1, 0)).EntireRow.Insert
      Range(cell, cell.Offset(cell.Value - 1, 1)).EntireRow.FillDown
    End If
   Set cell = cell.Offset(cell.Value, 0)
    Loop
End Sub

