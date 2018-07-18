Attribute VB_Name = "modBubble"
Option Explicit

Sub BubbleSort(list() As Long)
'   Sorts an array using bubble sort algorithm
'   Works with any data type
    Dim First As Integer, Last As Long
    Dim i As Long, j As Long
    Dim Temp
    
    First = LBound(list)
    Last = UBound(list)
    For i = First To Last - 1
        For j = i + 1 To Last
            If list(i) > list(j) Then
                Temp = list(j)
                list(j) = list(i)
                list(i) = Temp
            End If
        Next j
    Next i
End Sub

