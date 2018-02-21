Attribute VB_Name = "Sorting"
Option Explicit

    Sub SortArrayWithExcel()
      Dim myIntArray() As Integer
      Dim i As Integer
      Dim x As Integer
      Dim y As Integer
      Dim r As Integer
      Dim myDataRng As Range
    
      'initialize random number generator
      Randomize
    
      ReDim myIntArray(1 To 10)
    
      ' Fill the array with 10 random numbers between 1 and 100
      For i = 1 To 10
          myIntArray(i) = Int((100 * Rnd) + 1)
          Debug.Print "aValue" & i & ":" & vbTab & myIntArray(i)
      Next

      'write array to a worksheet
      Worksheets.Add

      r = 1   'row counter
      With ActiveSheet
        For i = 1 To 10
            Cells(r, 1).Value = myIntArray(i)
            r = r + 1
        Next i
      End With
    
      'Use Excel Sort to order values in the worksheet
      Set myDataRng = ActiveSheet.UsedRange
    
      With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), _
          SortOn:=xlSortOnValues, Order:=xlAscending, _
          DataOption:=xlSortNormal
        .SetRange myDataRng
        .Header = xlNo
        .MatchCase = False
        .Apply
      End With
    
      'free the memory used by array by using Erase statement
      Erase myIntArray
 
      ReDim myIntArray(1 To 10)
    
      'load sorted values back into an array
    
      For i = 1 To 10
        myIntArray(i) = ActiveSheet.Cells(i, 1).Value
      Next
        
      'write out sorted array to the Immediate Window
    
      i = 1
      For i = 1 To 10
        Debug.Print "aValueSorted: " & myIntArray(i)
      Next
    
      'find minimum and maximum values stored in the array
      x = myIntArray(1)
      y = myIntArray(UBound(myIntArray))
    
      Debug.Print "Min value=" & x & vbTab; "Max value=" & y
    
End Sub


Sub ResortedArray()
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
  'reformat all numeric values as currency
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
    
  'call the procedure to sort array
  BubbleSort myArray
    
  'empty the sorted array into a new worksheet
  Worksheets.Add

  r = 1   'row counter
  With ActiveSheet
      For i = 1 To UBound(myArray)
          Cells(r, 1).Value = myArray(i)
          r = r + 1
      Next i
  End With
End Sub


Sub BubbleSort(myArray As Variant)
  Dim i As Integer
  Dim j As Integer
  Dim uBnd As Integer
  Dim Temp As Variant
  uBnd = UBound(myArray)
    
  For i = LBound(myArray) To uBnd - 1
      For j = i + 1 To uBnd
          If UCase(myArray(i)) > UCase(myArray(j)) Then
              Temp = myArray(j)
              myArray(j) = myArray(i)
              myArray(i) = Temp
          End If
      Next j
  Next i
End Sub


