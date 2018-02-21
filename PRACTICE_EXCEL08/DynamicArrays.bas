Attribute VB_Name = "DynamicArrays"
Option Explicit

Sub DynArray()
  Dim counter As Integer

  'declare a dynamic array
  Dim myArray() As Integer

  'specify the initial size of the array
  ReDim myArray(1 To 5)

  Workbooks.Add

  'populate myArray with values
  For counter = 1 To 5
      myArray(counter) = counter + 1
      ActiveCell.Offset(counter - 1, 0).Value = myArray(counter)
  Next

  'change the size of myArray to hold 10 elements
  ReDim Preserve myArray(1 To 10)

  'add new values to myArray
  For counter = 6 To 10
      myArray(counter) = counter * counter
      With ActiveCell.Offset(counter - 1, 0)
          .Value = myArray(counter)
          .Font.Bold = True
      End With
  Next counter
End Sub

Sub LoadArrayFromWorksheet()
  Dim myDataRng As Range
  Dim myArray() As Variant
  Dim cnt As Integer
  Dim i As Integer
  Dim cell As Variant
  Dim r As Integer
  Dim last As Integer
        
  Set myDataRng = ActiveSheet.UsedRange
      
  'get the count of nonempty cells (text and numbers only)
  last = myDataRng.SpecialCells(xlCellTypeConstants, 3).Count
    
  If IsEmpty(myDataRng) Then
      MsgBox "Sheet is empty."
      Exit Sub
  End If
        
  ReDim myArray(1 To last)

  i = 1
  
  'fill the array from worksheet data
  'reformat all numeric values
  For Each cell In myDataRng
      If cell.Value <> "" Then
          If IsNumeric(cell.Value) Then
              myArray(i) = Format(cell.Value, "$#,#00.00")
          Else
              myArray(i) = cell.Value
          End If
          i = i + 1
      End If
  Next
      
  'print array values to the Immediate window
  For i = 1 To last
      Debug.Print myArray(i)
  Next
  Debug.Print "Items in the array: " & UBound(myArray)
End Sub

