Attribute VB_Name = "modCounting"
Option Explicit

Sub CountingSort(list)
    Dim counts()
    Dim i As Long
    Dim j As Long
    Dim next_index As Long
    Dim min, max
    Dim min_value As Variant, max_value As Variant

'   Allocate the counts array. VBA automatically
'   initialises all entries to 0.

    min_value = Minimum(list)
    max_value = Maximum(list)

    min = LBound(list)
    max = UBound(list)
    
    ReDim counts(min_value To max_value)
    
    ' Count the values.
    For i = min To max
        counts(list(i)) = counts(list(i)) + 1
    Next i

    ' Write the items back into the list array.
    next_index = min
    For i = min_value To max_value
        For j = 1 To counts(i)
            list(next_index) = i
            next_index = next_index + 1
        Next j
    Next i
End Sub

Function Minimum(list)
    Dim i As Long
    Minimum = list(LBound(list))
    For i = LBound(list) To UBound(list)
        If list(i) < Minimum Then Minimum = list(i)
    Next i
End Function

Function Maximum(list)
    Dim i As Long
    Maximum = list(LBound(list))
    For i = LBound(list) To UBound(list)
        If list(i) > Maximum Then Maximum = list(i)
    Next i
End Function

