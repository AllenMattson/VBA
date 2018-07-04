Attribute VB_Name = "UniqueItems_Function"
Option Base 1

Function UniqueItems(ArrayIn, Optional Count As Variant) As Variant
'   Accepts an array or range as input
'   If Count = True or is missing, the function returns the number of unique elements
'   If Count = False, the function returns a variant array of unique elements
    Dim Unique() As Variant ' array that holds the unique items
    Dim Element As Variant
    Dim i As Integer
    Dim FoundMatch As Boolean
'   If 2nd argument is missing, assign default value
    If IsMissing(Count) Then Count = True
'   Counter for number of unique elements
    NumUnique = 0
'   Loop thru the input array
    For Each Element In ArrayIn
        FoundMatch = False
'       Has item been added yet?
        For i = 1 To NumUnique
            If Element = Unique(i) Then
                FoundMatch = True
                Exit For '(exit loop)
            End If
        Next i
AddItem:
'       If not in list, add the item to unique list
        If Not FoundMatch And Not IsEmpty(Element) Then
            NumUnique = NumUnique + 1
            ReDim Preserve Unique(NumUnique)
            Unique(NumUnique) = Element
        End If
    Next Element
'   Assign a value to the function
    If Count Then UniqueItems = NumUnique Else UniqueItems = Unique
End Function
