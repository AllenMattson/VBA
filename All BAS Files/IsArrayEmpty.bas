Public Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
'Originally from http://www.cpearson.com/excel/vbaarrays.htm
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Dim LB As Long
  Dim UB As Long
  
  err.Clear
  On Error Resume Next
  If IsArray(Arr) = False Then
      ' we weren't passed an array, return True
      IsArrayEmpty = True
  End If
  
  ' Attempt to get the UBound of the array. If the array is
  ' unallocated, an error will occur.
  UB = UBound(Arr, 1)
  If (err.Number <> 0) Then
      IsArrayEmpty = True
  Else
      ''''''''''''''''''''''''''''''''''''''''''
      ' On rare occassion, under circumstances I
      ' cannot reliably replictate, Err.Number
      ' will be 0 for an unallocated, empty array.
      ' On these occassions, LBound is 0 and
      ' UBoung is -1.
      ' To accomodate the weird behavior, test to
      ' see if LB > UB. If so, the array is not
      ' allocated.
      ''''''''''''''''''''''''''''''''''''''''''
      err.Clear
      LB = LBound(Arr)
      If LB > UB Then
          IsArrayEmpty = True
      Else
          IsArrayEmpty = False
      End If
  End If

End Function
