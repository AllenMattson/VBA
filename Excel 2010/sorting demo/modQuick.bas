Attribute VB_Name = "modQuick"
Option Explicit

'This technique works only with Integer or Long values.

Public Sub Quicksort(list() As Long, ByVal min As Long, ByVal max As Long)
    Dim med_value As Long
    Dim hi As Long
    Dim lo As Long
    Dim i As Long

    ' If min >= max, the list contains 0 or 1 items so it
    ' is sorted.
    If min >= max Then Exit Sub

    ' Pick the dividing value.
    i = Int((max - min + 1) * Rnd + min)
    med_value = list(i)

    ' Swap it to the front.
    list(i) = list(min)

    lo = min
    hi = max
    Do
        ' Look down from hi for a value < med_value.
        Do While list(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            list(lo) = med_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        list(lo) = list(hi)
        
        ' Look up from lo for a value >= med_value.
        lo = lo + 1
        Do While list(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            list(hi) = med_value
            Exit Do
        End If
        
        ' Swap the lo and hi values.
        list(hi) = list(lo)
    Loop
    
    ' Sort the two sublists.
    Quicksort list(), min, lo - 1
    Quicksort list(), lo + 1, max
End Sub
