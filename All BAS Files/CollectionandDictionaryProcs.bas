Attribute VB_Name = "CollectionandDictionaryProcs"
 'http://www.cpearson.com/excel/CollectionsAndDictionaries.htm
 'The following procedures are described on this page:
 '         ArrayToCollection
 '         ArrayToDictionary
 '         CollectionToArray
 '         CollectionToDictionary
 '         CollectionToRange
 '         CreateDictionaryKeyFromCollectionItem
 '         DictionaryToArray
 '         DictionaryToCollection
 '         DictionaryToRange
 '         KeyExistsInCollection
 '         RangeToCollection
 '         RangeToDictionary
 '         SortCollection
 '         SortDictionary
 
 
Public Function ArrayToCollection(Arr As Variant, ByRef Coll As Collection) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ArrayToCollection
' This function converts an array to a Collection. Arr may be either a 1-dimensional
' arrary or a two-dimensional array. If Arr is a 1-dimensional array, each element
' of the array is added to Coll without a key. If Arr is a 2-dimensional array,
' the first column is assumed to the be Item to be added, and the second column
' is assumed to be the Key for that item.
' Items are added to the Coll collection. Existing contents are preserved.
' This function returns True if successful, or False if an error occurs.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim KeyVal As String

''''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(Arr) = False Then
    ArrayToCollection = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Work with either a 1-dimensional
' or 2-dimensional array. Any other
' number of dimensions will cause
' a error. Use On Error to
' trap for errors (most likely a
' duplicate key error).
'''''''''''''''''''''''''''''''''''
On Error GoTo ErrH:
Select Case NumberOfArrayDimensions(Arr:=Arr)
    Case 0
        '''''''''''''''''''''''''''''''
        ' Unallocated array. Exit with
        ' error.
        '''''''''''''''''''''''''''''''
        ArrayToCollection = False
        Exit Function
        
    Case 1
        ''''''''''''''''''''''''''''''
        ' Arr is a single dimensional
        ' array. Load the elements of
        ' the array without keys.
        ''''''''''''''''''''''''''''''
        For Ndx = LBound(Arr) To UBound(Arr)
            Coll.Add Item:=Arr(Ndx)
        Next Ndx
    
    Case 2
        '''''''''''''''''''''''''''''
        ' Arr is a two-dimensional
        ' array. The first column
        ' is the Item and the second
        ' column is the Key.
        '''''''''''''''''''''''''''''
        For Ndx = LBound(Arr, 1) To UBound(Arr, 1)
            KeyVal = Arr(Ndx, 1)
            If Trim(KeyVal) = vbNullString Then
                '''''''''''''''''''''''''''''''''
                ' Key is empty. Add to collection
                ' without a key.
                '''''''''''''''''''''''''''''''''
                Coll.Add Item:=Arr(Ndx, 1)
            Else
                '''''''''''''''''''''''''''''''''
                ' Key is not empty. Add with key.
                '''''''''''''''''''''''''''''''''
                Coll.Add Item:=Arr(Ndx, 0), key:=KeyVal
            End If
        Next Ndx
    
    Case Else
        '''''''''''''''''''''''''''''
        ' The array has 3 or more
        ' dimensions. Return an
        ' error.
        '''''''''''''''''''''''''''''
        ArrayToCollection = False
        Exit Function

End Select

ArrayToCollection = True
Exit Function

ErrH:
    ''''''''''''''''''''''''''''''''
    ' An error occurred, most likely
    ' a duplicate key error. Return
    ' False.
    ''''''''''''''''''''''''''''''''
    ArrayToCollection = False

End Function

Public Function ArrayToDictionary(Arr As Variant, Dict As Scripting.Dictionary) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ArrayToDictionary
' This function loads the contents of a two dimensional array into the Dict dictionary
' object. Arr must be two dimensional. The first column is the Item to add to the Dict
' dictionary, and the second column is the Key value of the Item. The existing items
' in the dictionary are left intact.
' The function returns True if successful, or False if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim ItemVar As Variant
Dim KeyVal As String

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(Arr) = False Then
    ArrayToDictionary = False
    Exit Function
End If

'''''''''''''''''''''''''''''''
' Ensure Arr is two dimensional
'''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=Arr) <> 2 Then
    ArrayToDictionary = False
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''''
' Loop through the arary and
' add the items to the Dictionary.
'''''''''''''''''''''''''''''''''''
On Error GoTo ErrH:
For Ndx = LBound(Arr, 1) To UBound(Arr, 1)
    Dict.Add key:=Arr(Ndx, LBound(Arr, 2) + 1), Item:=Arr(Ndx, LBound(Arr, 2))
Next Ndx
    
'''''''''''''''''
' Return Success.
'''''''''''''''''
ArrayToDictionary = True
Exit Function

ErrH:
ArrayToDictionary = False

End Function

Public Function CollectionToArray(Coll As Collection, Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CollectionToArray
' This function converts a collection object to a single dimensional array.
' The elements of Collection may be any type of data except User Defined Types.
' The procedure will populate the array Arr with the elements of the collection.
' Only the collection items, not the keys, are stored in Arr. The function returns
' True if the the Collection was successfully converted to an array, or False
' if an error occcurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim V As Variant
Dim Ndx As Long

''''''''''''''''''''''''''''''
' Ensure Coll is not Nothing.
''''''''''''''''''''''''''''''
If Coll Is Nothing Then
    CollectionToArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''
' Ensure Arr is an array and
' is dynamic.
''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    CollectionToArray = False
    Exit Function
End If
If IsArrayDynamic(Arr:=Arr) = False Then
    CollectionToArray = False
    Exit Function
End If

''''''''''''''''''''''''''''
' Ensure Coll has at least
' one item.
''''''''''''''''''''''''''''
If Coll.Count < 1 Then
    CollectionToArray = False
    Exit Function
End If
    
''''''''''''''''''''''''''''''
' Redim Arr to the number of
' elements in the collection.
'''''''''''''''''''''''''''''
ReDim Arr(1 To Coll.Count)
'''''''''''''''''''''''''''''
' Loop through the colletcion
' and add the elements of
' Collection to Arr.
'''''''''''''''''''''''''''''
For Ndx = 1 To Coll.Count
    If IsObject(Coll(Ndx)) = True Then
        Set Arr(Ndx) = Coll(Ndx)
    Else
        Arr(Ndx) = Coll(Ndx)
    End If
Next Ndx

CollectionToArray = True

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The  CreateDictionaryKeyFromCollectionItem  procedure is declared as follows:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '   Private Function CreateDictionaryKeyFromCollectionItem( _
 '       Dict As Scripting.Dictionary, _
  '      Item As Variant) As String
 Public Function CollectionToDictionary(Coll As Collection, Dict As Scripting.Dictionary) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CollectionToDictionary
'
' This function converts a Collection Objct to a
' Dictionary object. This code requires a reference
' the Microsoft Scripting RunTime Library.
'
' It calls a private procedure named
' CreateDictionaryKeyFromCollectionItem that you supply
' to create a Dictionary Key from an Item in the Collection.
' This must return a String value that will be unique within
' the Dictionary.
'
' If an error occurs (e.g., a Key value returned by
' CreateDictionaryKeyFromCollectionItem already exists
' in the Dictionary object), Dictionary is set to Nothing.
' The function returns True if the conversion from Collection
' to Dictionary was successful, or False if an error occurred.
' If it returns False, the Dictionary is set to Nothing.
'
' The code destroys the existing contents of Dict
' and replaces them with the new elements. The Coll
' Collection is left intact with no changes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim ItemKey As String
Dim ItemVar As Variant

''''''''''''''''''''''''''''''''''''''''''''
' Ensure Coll is not Nothing.
''''''''''''''''''''''''''''''''''''''''''''
If (Coll Is Nothing) Then
    CollectionToDictionary = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''
' Reset Dict to a new, empty Dictionary
''''''''''''''''''''''''''''''''''''''''''''
Set Dict = Nothing
Set Dict = New Scripting.Dictionary
'''''''''''''''''''''''''''''''''''''''''''
' Ensure we have at least one element in
' the collection object.
'''''''''''''''''''''''''''''''''''''''''''
If Coll.Count = 0 Then
    Set Dict = Nothing
    CollectionToDictionary = False
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''''''''''''
' Loop through the collection and convert
' each item in the collection to an item
' for the dictionary. Call
' CreateDictionaryKeyFromCollectionItem
' to get the Key to be used in the Dictionary
' item.
'''''''''''''''''''''''''''''''''''''''''''
For Ndx = 1 To Coll.Count
    '''''''''''''''''''''''''''''''''''''''
    ' Coll may contain object variables.
    ' Test for this condition and set
    ' ItemVar appropriately.
    '''''''''''''''''''''''''''''''''''''''
    If IsObject(Coll(Ndx)) = True Then
        Set ItemVar = Coll(Ndx)
    Else
        ItemVar = Coll(Ndx)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call the user-supplied CreateDictionaryKeyFromCollectionItem
    ' function to get the Key to be used in the Dictionary.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ItemKey = CreateDictionaryKeyFromCollectionItem(Dict:=Dict, Item:=ItemVar)
    ''''''''''''''''''''''''''''''''
    ' ItemKey must not be spaces or
    ' an empty string.
    ''''''''''''''''''''''''''''''''
    If Trim(ItemKey) = vbNullString Then
        CollectionToDictionary = False
        Exit Function
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' See if ItemKey already exists in the Dictionary.
    ' If so, return False. You can't have duplicate keys.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Dict.Exists(key:=ItemKey) = True Then
        Set Dict = Nothing
        CollectionToDictionary = False
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ItemKey does not exist in Dict, so add ItemVar to
    ' Dict with a key of ItemKey.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dict.Add key:=ItemKey, Item:=ItemVar
Next Ndx
CollectionToDictionary = True

End Function

Private Function CreateDictionaryKeyFromCollectionItem( _
    Dict As Scripting.Dictionary, _
    Item As Variant) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateDictionaryKeyFromCollectionItem
' This function is called by CollectionToDictionary to create
' a Key for a Dictionary item that is take from a Collection
' item. The collection item is passed in the Item parameter.
' It is up to you to create a unique key based on the
' Item parameter.
' Dict is the Dictionary for which the result of this function
' will be used as a Key, and Item is the item of the
' Dictionary for which this procedure is creating a Key.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ItemKey As String
''''''''''''''''''''''''''''''''''''''''''
' Your code to set ItemKey to the
' appropriate string value. ItemKey
' must not be all spaces or vbNullString.
''''''''''''''''''''''''''''''''''''''''''


CreateDictionaryKeyFromCollectionItem = ItemKey
End Function

Public Function CollectionToRange(Coll As Collection, StartCells As Range) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CollectionToRange
' This procedure writes the contents of a Collection Coll to a range of cells starting
' in StartCells. If StartCells is a single cell, the contents of Collection are
' written downward in a single column starting in StartCell. If StartCell is
' two cells, the Collection is written in the same orientation (down a column or
' across a row) as StartCells. If StartCells is more than two cells, ONLY those
' cells will be written to, moving across then down. StartCells must be a single
' area range.
'
' If an item in Coll is an object, it is skipped.
'
' The function returns True if successful or False if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim DestRng As Range
Dim V As Variant
Dim Ndx As Long

'''''''''''''''''''''''''''''''''''''
' Ensure parameters are not Nothing.
'''''''''''''''''''''''''''''''''''''
If (Coll Is Nothing) Or (StartCells Is Nothing) Then
    CollectionToRange = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure StartCells is a single area.
'''''''''''''''''''''''''''''''''''''
If StartCells.Areas.Count > 1 Then
    CollectionToRange = False
    Exit Function
End If

If StartCells.Cells.Count = 1 Then
    '''''''''''''''''''''''''''''''''''''
    ' StartCells is one cell. Write out
    ' the collection moving downwards.
    '''''''''''''''''''''''''''''''''''''
    Set DestRng = StartCells
    For Each V In Coll
        If IsObject(V) = False Then
            DestRng.value = V
            If DestRng.Row < DestRng.Parent.Rows.Count Then
                Set DestRng = DestRng(2, 1)
            Else
                CollectionToRange = False
                Exit Function
            End If
                
        End If
    Next V
    CollectionToRange = True
    Exit Function
End If

If StartCells.Cells.Count = 2 Then
    ''''''''''''''''''''''''''''''''''
    ' Test the orientation of the two
    ' cells in StartCells.
    ''''''''''''''''''''''''''''''''''
    If StartCells.Rows.Count = 1 Then
        '''''''''''''''''''''''''''''''''
        ' Write out the Colleciton moving
        ' across the row.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In Coll
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.Column < StartCells.Parent.Columns.Count Then
                    Set DestRng = DestRng(1, 2)
                Else
                    CollectionToRange = False
                    Exit Function
                End If
            End If
        Next V
        CollectionToRange = True
        Exit Function
    Else
        '''''''''''''''''''''''''''''''''
        ' Write out the Colleciton moving
        ' down the column.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In Coll
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.Row < StartCells.Parent.Rows.Count Then
                    Set DestRng = DestRng(2, 1)
                Else
                    CollectionToRange = False
                    Exit Function
                End If
            End If
        Next V
        CollectionToRange = True
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''''''''''
' Write the collection only into
' Cells of StartCells.
'''''''''''''''''''''''''''''''''''''
For Ndx = 1 To StartCells.Cells.Count
    If Ndx <= Coll.Count Then
        V = Coll(Ndx)
        If IsObject(V) = False Then
            StartCells.Cells(Ndx).value = V
        End If
    End If
Next Ndx

CollectionToRange = True


End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This procedure populates the array Arr with the items and keys for the Dict Dictionary. Arr must be declared as an array of variants, e.g.,

'    Dim Arr() As Variant
'The array created is a 0-based 2-dimensional array. Each row in Arr is one item of the Dict. Column 0 of Arr is the Key of the dictionary element, and column 1 of Arr is the Item in the Dictionary.
''The Item of the Dictionary element may be any type, including objects, of variable except a User-Defined type.   The Arr is Redimed as follows:

'    ReDim Arr(0 To Dict.Count - 1, 0 To 1)
'The original contents of Arr are destroyed. The code for DictionaryToArray is shown below:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DictionaryToArray(Dict As Scripting.Dictionary, Arr As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DictionaryToArray
' This creates a 0-based, 2-dimensional array Arr from a Dictionary object. Each
' row of the array is one element of the Dictionary. The first column of the array is the
' Key of the dictionary item, and the second column is the Key of the item in the
' dictionary. Arr MUST be an dynamic array of Variants, e.g.,
' Dim Arr() As Variant
' The VarType of Arr is tested, and if it does not equal 8204 (vbArray + vbVariant) an
' error occurs.
'
' The existing content of Arr is destroyed. The function returns True if successsful
' or False if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

'''''''''''''''''''''''''''''
' Ensure that Arr is an array
' of Variants.
'''''''''''''''''''''''''''''
If VarType(Arr) <> (vbArray + vbVariant) Then
    DictionaryToArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Arr is a dynamic array.
''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=Arr) = False Then
    DictionaryToArray = False
    Exit Function
End If
   
'''''''''''''''''''''''''''''
' Ensure Dict is not nothing.
'''''''''''''''''''''''''''''
If Dict Is Nothing Then
    DictionaryToArray = False
    Exit Function
End If
    
'''''''''''''''''''''''''''
' Ensure that Dict contains
' at least one entry.
'''''''''''''''''''''''''''
If Dict.Count = 0 Then
    DictionaryToArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''
' Redim the Arr variable.
'''''''''''''''''''''''''''''
ReDim Arr(0 To Dict.Count - 1, 0 To 1)

For Ndx = 0 To Dict.Count - 1
    Arr(Ndx, 0) = Dict.Keys(Ndx)
    '''''''''''''''''''''''''''''''''''''''''
    ' Test to see if the item in the Dict is
    ' an object. If so, use Set.
    '''''''''''''''''''''''''''''''''''''''''
    If IsObject(Dict.Items(Ndx)) = True Then
        Set Arr(Ndx, 1) = Dict.Items(Ndx)
    Else
        Arr(Ndx, 1) = Dict.Items(Ndx)
    End If

Next Ndx

'''''''''''''''''
' Return success.
'''''''''''''''''
DictionaryToArray = True

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This procedure converts a Dictionary object Dict to a Collection object Coll.
'This procedure converts an existing Dictionary to a new Collection object.
'Keys from the Dictionary are used as the keys for the Collection.
'This function returns True if successful, or False if an error occurred.
'The contents of Dict are not modified.  If PreserveColl is omitted or False, the existing contents of the Coll collection are destroyed.
'If PreserveColl is True, the existing contents of Coll are preserved. If PreserveColl is true, then the possibility exists that we will run into duplicate key values for the Collection.
'If StopOnDuplicateKey is omitted or false, this error is ignored, but the item from the Dict Dictionary will not be added to Coll Collection.
'If StopOnDuplicateKey is True, the procedure will terminate, and not all of the items in the Dict Dictionary will have copied to the Coll Collection.
'The Coll Collection will be in an indeterminant state. The code for the DictionaryToCollection is shown below:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DictionaryToCollection(Dict As Scripting.Dictionary, Coll As Collection, _
    Optional PreserveColl As Boolean = False, _
    Optional StopOnDuplicateKey As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DictionaryToCollection
' This procedure converts an existing Dictionary to a new Collection object. Keys from
' the Dictionary are used as the keys for the Collection. This function returns True
' if successful, or False if an error occurred. The contents of Dict are not modified.
' If PreserveColl is omitted or False, the existing contents of the Coll collection are
' destroyed. If PreserveColl is True, the existing contents of Coll are preserved.
' If PreserveColl is true, then the possibility exists that we will run into duplicate
' key values for the Collection. If StopOnDuplicateKey is omitted or false, this error
' is ignored, but the item from the Dict Dictionary will not be added to Coll Collection.
' If StopOnDuplicateKey is True, the procedure will terminate, and not all of the items in
' the Dict Dictionary will have copied to the Coll Collection. The Coll Collection will
' be in an indeterminant state.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim ItemVar As Variant
Dim KeyVal As String

''''''''''''''''''''''''''''''''
' Ensure Dict is not Nothing
''''''''''''''''''''''''''''''''
If Dict Is Nothing Then
    DictionaryToCollection = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' If PreseveColl is omitted or
' False, destroy the existing
' Coll Collection.
'''''''''''''''''''''''''''''''''
If PreserveColl = False Then
    Set Coll = Nothing
    Set Coll = New Collection
End If

'''''''''''''''''''''''''''''''''
' Loop through the Dictionary
' and transfer the data to
' the Collection.
'''''''''''''''''''''''''''''''''
On Error Resume Next
For Ndx = 0 To Dict.Count - 1
    If IsObject(Dict.Items(Ndx)) = True Then
        Set ItemVar = Dict.Items(Ndx)
    Else
        ItemVar = Dict.Items(Ndx)
    End If
    KeyVal = Dict.Keys(Ndx)
    Err.Clear
    Coll.Add Item:=ItemVar, key:=KeyVal
    If Err.Number <> 0 Then
        If StopOnDuplicateKey = True Then
            DictionaryToCollection = False
            Exit Function
        End If
    End If
Next Ndx
DictionaryToCollection = True
End Function

Public Function DictionaryToRange(Dict As Scripting.Dictionary, StartCells As Range) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DictionaryToRange
' This procedure writes the contents of a Dictionary Dict to a range of cells starting
' in StartCells. If StartCells is a single cell, the contents of Dict are
' written downward in a single column starting in StartCell. If StartCell is
' two cells, the Dictionary is written in the same orientation (down a column or
' across a row) as StartCells. If StartCells is more than two cells, ONLY those
' cells will be written to, moving across then down. StartCells must be a single
' area range.
'
' If an item in Dict is an object, it is skipped.
'
' The function returns True if successful or False if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim DestRng As Range
Dim V As Variant
Dim Ndx As Long

'''''''''''''''''''''''''''''''''''''
' Ensure parameters are not Nothing.
'''''''''''''''''''''''''''''''''''''
If (Dict Is Nothing) Or (StartCells Is Nothing) Then
    DictionaryToRange = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure StartCells is a single area.
'''''''''''''''''''''''''''''''''''''
If StartCells.Areas.Count > 1 Then
    DictionaryToRange = False
    Exit Function
End If

If StartCells.Cells.Count = 1 Then
    '''''''''''''''''''''''''''''''''''''
    ' StartCells is one cell. Write out
    ' the collection moving downwards.
    '''''''''''''''''''''''''''''''''''''
    Set DestRng = StartCells
    For Each V In Dict.Items
        If IsObject(V) = False Then
            DestRng.value = V
            If DestRng.Row < DestRng.Parent.Rows.Count Then
                Set DestRng = DestRng(2, 1)
            Else
                DictionaryToRange = False
                Exit Function
            End If
                
        End If
    Next V
    DictionaryToRange = True
    Exit Function
End If

If StartCells.Cells.Count = 2 Then
    ''''''''''''''''''''''''''''''''''
    ' Test the orientation of the two
    ' cells in StartCells.
    ''''''''''''''''''''''''''''''''''
    If StartCells.Rows.Count = 1 Then
        '''''''''''''''''''''''''''''''''
        ' Write out the Colleciton moving
        ' across the row.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In Dict.Items
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.Column < StartCells.Parent.Columns.Count Then
                    Set DestRng = DestRng(1, 2)
                Else
                    DictionaryToRange = False
                    Exit Function
                End If
            End If
        Next V
        DictionaryToRange = True
        Exit Function
    Else
        '''''''''''''''''''''''''''''''''
        ' Write out the Dictionary moving
        ' down the column.
        '''''''''''''''''''''''''''''''''
        Set DestRng = StartCells.Cells(1, 1)
        For Each V In Dict.Items
            If IsObject(V) = False Then
                DestRng.value = V
                If DestRng.Row < StartCells.Parent.Rows.Count Then
                    Set DestRng = DestRng(2, 1)
                Else
                    DictionaryToRange = False
                    Exit Function
                End If
            End If
        Next V
        DictionaryToRange = True
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''''''''''
' Write the Dictionary only into
' Cells of StartCells.
'''''''''''''''''''''''''''''''''''''
For Ndx = 1 To StartCells.Cells.Count
    If Ndx <= Dict.Count Then
        V = Dict.Items(Ndx - 1)
        If IsObject(V) = False Then
            StartCells.Cells(Ndx).value = V
        End If
    End If
Next Ndx

DictionaryToRange = True


End Function

Public Function KeyExistsInCollection(Coll As Collection, KeyName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' KeyExistsInCollection
' This function returns True if KeyName exists the Coll Collection,
' or False if KeyName does not exist.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim V As Variant
    On Error Resume Next
    Err.Clear
    V = Coll(KeyName)
    If Err.Number = 0 Then
        KeyExistsInCollection = True
    Else
        KeyExistsInCollection = False
    End If
End Function

Function RangeToCollection(KeyRange As Range, ItemRange As Range, Coll As Collection, _
    Optional RangeAsObject As Boolean = False, _
    Optional StopOnDuplicateKey As Boolean = True, _
    Optional ReplaceOnDuplicateKey As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RangeToCollection
' This function load an existing Collection Coll with items from worksheet
' ranges.
'
' The KeyRange and ItemRange must be the same size. Each element in KeyRange
' is the Key value for the corresponding item in ItemRange.
'
' KeyRange may be Nothing. In this case, the items in ItemRange are added to
' the Collection Coll without keys.
'
' If RangeAsObject is omitted of False, the Items added to the Collection are
' the values in the cells of ItemRange. If RangeAsObject is True, the cells
' are added as objects to the Collection.
'
' If a duplicate key is encountered when adding an item to Coll, the code
' will do one of the following:
'   If StopOnDuplicateKey is omitted or True, the funcion stops processing
'   and returns False. Items added to the Collection before the duplicate key
'   was encountered remain in the Collection.
'
'   If StopOnDuplicateKey is False, then if ReplaceOnDuplicateKey is False,
'   the Item that caused the duplicate key error is not added to the Collection
'   but processing continues with the rest of the items in the range. If
'   ReplaceOnDuplicateKey if True, the existing item in the Collection is
'   deleted and replaced with the new item.
'
' If Coll is Nothing, it will be created as a new Collection.
'
' The function returns True if all items were added to the Collection, or False
' if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim IRng As Range
Dim KeyExists As Boolean
Dim KeyNdx As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the KeyRange and ItemRange variables are not
' Nothing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (ItemRange Is Nothing) Then
    RangeToCollection = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure KeyRange and ItemRange as the same size.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not KeyRange Is Nothing Then
    If (KeyRange.Rows.Count <> ItemRange.Rows.Count) Or _
        (KeyRange.Columns.Count <> ItemRange.Columns.Count) Then
        RangeToCollection = False
        Exit Function
    End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure both KeyRange and ItemRange are single area
' ranges.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If ItemRange.Areas.Count > 1 Then
    RangeToCollection = False
    Exit Function
End If

If Not KeyRange Is Nothing Then
    If KeyRange.Areas.Count > 1 Then
        RangeToCollection = False
        Exit Function
    End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If Coll is Nothing, create a new Collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Coll Is Nothing Then
    Set Coll = New Collection
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through ItemRange, testing whether the Key exists
' and adding items to the Collection.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Each IRng In ItemRange.Cells
    KeyNdx = KeyNdx + 1
    If KeyRange Is Nothing Then
        KeyExists = False
    Else
        KeyExists = KeyExistsInCollection(Coll:=Coll, KeyName:=KeyRange.Cells(KeyNdx))
    End If
    
    If KeyExists = True Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The key already exists in the Collection.
        ' Determine what to do.
        '''''''''''''''''''''''''''''''''''''''''''
        If StopOnDuplicateKey = True Then
            RangeToCollection = False
            Exit Function
        Else
            ''''''''''''''''''''''''''''''''''''''
            ' Do nothing here. Test the value of
            ' ReplaceOnDuplicateKey below.
            ''''''''''''''''''''''''''''''''''''''
        End If
        '''''''''''''''''''''''''''''''''''''''''
        ' If ReplaceOnDuplicateKey is True then
        ' remove the existing entry. Otherwise,
        ' exit the function.
        '''''''''''''''''''''''''''''''''''''''''
        If ReplaceOnDuplicateKey = True Then
            Coll.Remove KeyRange.Cells(KeyNdx)
            KeyExists = False
        Else
            If StopOnDuplicateKey = True Then
                RangeToCollection = False
                Exit Function
            End If
        End If
    End If
    If KeyExists = False Then
        '''''''''''''''''''''''''''''''
        ' Check KeyRange  to see if
        ' we're adding with Keys.
        '''''''''''''''''''''''''''''''
        If Not KeyRange Is Nothing Then
            '''''''''''''''''''''''''
            ' Add with key.
            '''''''''''''''''''''''''
            If RangeAsObject = True Then
                Coll.Add Item:=IRng, key:=KeyRange.Cells(KeyNdx)
            Else
                Coll.Add Item:=IRng.Text, key:=KeyRange.Cells(KeyNdx)
            End If
        Else
            '''''''''''''''''''''
            ' Add without key.
            If RangeAsObject = True Then
                Coll.Add Item:=IRng
            Else
                Coll.Add Item:=IRng.Text
            End If
            '''''''''''''''''''''
            
        End If
    End If
Next IRng

'''''''''''''''''
' Return Success.
'''''''''''''''''
RangeToCollection = True

End Function
Function RangeToDictionary(KeyRange As Range, ItemRange As Range, Dict As Scripting.Dictionary, _
    Optional RangeAsObject As Boolean = False, _
    Optional StopOnDuplicateKey As Boolean = True, _
    Optional ReplaceOnDuplicateKey As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RangeToDictionary
' This funciton loads an existing Dictionary Dict with the keys and value from
' worksheet ranges.
' The KeyRange and ItemRange must be the same size. Each element in KeyRange
' is the Key value for the corresponding item in ItemRange.
'
' If RangeAsObject is omitted of False, the Items added to the Dictionary are
' the values in the cells of ItemRange. If RangeAsObject is True, the cells
' are added as objects to the Dictionary.
'
' If a duplicate key is encountered when adding an item to Dict, the code
' will do one of the following:
'   If StopOnDuplicateKey is omitted or True, the funcion stops processing
'   and returns False. Items added to the Dictionary before the duplicate key
'   was encountered remain in the Dictionary.
'
'   If StopOnDuplicateKey is False, then if ReplaceOnDuplicateKey is False,
'   the Item that caused the duplicate key error is not added to the Dictionary
'   but processing continues with the rest of the items in the range. If
'   ReplaceOnDuplicateKey if True, the existing item in the Dictionary is
'   deleted and replaced with the new item.
'
' If Dict is Nothing, it will be created as a new Dictionary.
'
' The function returns True if all items were added to the dictionary, or False
' if an error occurred.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim KRng As Range
Dim KeyExists As Boolean
Dim ItemNdx As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the KeyRange and ItemRange variables are not
' Nothing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (KeyRange Is Nothing) Or (ItemRange Is Nothing) Then
    RangeToDictionary = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure KeyRange and ItemRange as the same size.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (KeyRange.Rows.Count <> ItemRange.Rows.Count) Or _
    (KeyRange.Columns.Count <> ItemRange.Columns.Count) Then
    RangeToDictionary = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure both KeyRange and ItemRange are single area
' ranges.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (KeyRange.Areas.Count > 1) Or (ItemRange.Areas.Count > 1) Then
    RangeToDictionary = False
    Exit Function
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If Dict is Nothing, create a new dictionary.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Dict Is Nothing Then
    Set Dict = New Scripting.Dictionary
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through KeyRange, testing whether the Key exists
' and adding items to the Dictionary.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Each KRng In KeyRange.Cells
    ItemNdx = ItemNdx + 1
    KeyExists = Dict.Exists(key:=KRng.Text)
    If KeyExists = True Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The key already exists in the Dictionary.
        ' Determine what to do.
        '''''''''''''''''''''''''''''''''''''''''''
        If StopOnDuplicateKey = True Then
            RangeToDictionary = False
            Exit Function
        Else
            ''''''''''''''''''''''''''''''''''''''
            ' Do nothing here. Test the value of
            ' ReplaceOnDuplicateKey below.
            ''''''''''''''''''''''''''''''''''''''
        End If
        '''''''''''''''''''''''''''''''''''''''''
        ' If ReplaceOnDuplicateKey is True then
        ' remove the existing entry. Otherwise,
        ' exit the function.
        '''''''''''''''''''''''''''''''''''''''''
        If ReplaceOnDuplicateKey = True Then
            Dict.Remove key:=KRng.Text
            KeyExists = False
        Else
            If StopOnDuplicateKey = True Then
                RangeToDictionary = False
                Exit Function
            End If
        End If
    End If
    If KeyExists = False Then
        If RangeAsObject = True Then
            Dict.Add key:=KRng.Text, Item:=ItemRange.Cells(ItemNdx)
        Else
            Dict.Add key:=KRng.Text, Item:=ItemRange.Cells(ItemNdx).Text
        End If
    End If
Next KRng

'''''''''''''''''
' Return Success.
'''''''''''''''''
RangeToDictionary = True

End Function

Page

Public Sub SortDictionary(Dict As Scripting.Dictionary, _
    SortByKey As Boolean, _
    Optional Descending As Boolean = False, _
    Optional CompareMode As VbCompareMethod = vbTextCompare)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SortDictionary
' This sorts a Dictionary object. If SortByKey is False, the
' the sort is done based on the Items of the Dictionary, and
' these items must be simple data types. They may not be
' Object, Arrays, or User-Defined Types. If SortByKey is True,
' the Dictionary is sorted by Key value, and the Items in the
' Dictionary may be Object as well as simple variables.
'
' If sort by key is True, all element of the Dictionary
' must have a non-blank Key value. If Key is vbNullString
' the procedure will terminate.
'
' By defualt, sorting is done in Ascending order. You can
' sort by Descending order by setting the Descending parameter
' to True.
'
' By default, text comparisons are done case-INSENSITIVE (e.g.,
' "a" = "A"). To use case-SENSITIVE comparisons (e.g., "a" <> "A")
' set CompareMode to vbBinaryCompare.
'
' Note: This procedure requires the
' QSortInPlace function, which is described and available for
' download at www.cpearson.com/excel/qsort.htm .
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim KeyValue As String
Dim ItemValue As Variant
Dim Arr() As Variant
Dim KeyArr() As String
Dim VTypes() As VbVarType


Dim V As Variant
Dim SplitArr As Variant

Dim TempDict As Scripting.Dictionary
'''''''''''''''''''''''''''''
' Ensure Dict is not Nothing.
'''''''''''''''''''''''''''''
If Dict Is Nothing Then
    Exit Sub
End If
''''''''''''''''''''''''''''
' If the number of elements
' in Dict is 0 or 1, no
' sorting is required.
''''''''''''''''''''''''''''
If (Dict.Count = 0) Or (Dict.Count = 1) Then
    Exit Sub
End If

''''''''''''''''''''''''''''
' Create a new TempDict.
''''''''''''''''''''''''''''
Set TempDict = New Scripting.Dictionary

If SortByKey = True Then
    ''''''''''''''''''''''''''''''''''''''''
    ' We're sorting by key. Redim the Arr
    ' to the number of elements in the
    ' Dict object, and load that array
    ' with the key names.
    ''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To Dict.Count - 1)
    
    For Ndx = 0 To Dict.Count - 1
        Arr(Ndx) = Dict.Keys(Ndx)
    Next Ndx
    
    ''''''''''''''''''''''''''''''''''''''
    ' Sort the key names.
    ''''''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=CompareMode
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Load TempDict. The key value come from
    ' our sorted array of keys Arr, and the
    ' Item comes from the original Dict object.
    ''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = 0 To Dict.Count - 1
        KeyValue = Arr(Ndx)
        TempDict.Add key:=KeyValue, Item:=Dict.Item(KeyValue)
    Next Ndx
    '''''''''''''''''''''''''''''''''
    ' Set the passed in Dict object
    ' to our TempDict object.
    '''''''''''''''''''''''''''''''''
    Set Dict = TempDict
    ''''''''''''''''''''''''''''''''
    ' This is the end of processing.
    ''''''''''''''''''''''''''''''''
Else
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Here, we're sorting by items. The Items must
    ' be simple data types. They may NOT be Objects,
    ' arrays, or UserDefineTypes.
    ' First, ReDim Arr and VTypes to the number
    ' of elements in the Dict object. Arr will
    ' hold a string containing
    '   Item & vbNullChar & Key
    ' This keeps the association between the
    ' item and its key.
    '''''''''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To Dict.Count - 1)
    ReDim VTypes(0 To Dict.Count - 1)

    For Ndx = 0 To Dict.Count - 1
        If (IsObject(Dict.Items(Ndx)) = True) Or _
            (IsArray(Dict.Items(Ndx)) = True) Or _
            VarType(Dict.Items(Ndx)) = vbUserDefinedType Then
            Debug.Print "***** ITEM IN DICTIONARY WAS OBJECT OR ARRAY OR UDT"
            Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Here, we create a string containing
        '       Item & vbNullChar & Key
        ' This preserves the associate between an item and its
        ' key. Store the VarType of the Item in the VTypes
        ' array. We'll use these values later to convert
        ' back to the proper data type for Item.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Arr(Ndx) = Dict.Items(Ndx) & vbNullChar & Dict.Keys(Ndx)
            VTypes(Ndx) = VarType(Dict.Items(Ndx))
            
    Next Ndx
    ''''''''''''''''''''''''''''''''''
    ' Sort the array that contains the
    ' items of the Dictionary along
    ' with their associated keys
    ''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=vbTextCompare
    
    For Ndx = LBound(Arr) To UBound(Arr)
        '''''''''''''''''''''''''''''''''''''
        ' Loop trhogh the array of sorted
        ' Items, Split based on vbNullChar
        ' to get the Key from the element
        ' of the array Arr.
        SplitArr = Split(Arr(Ndx), vbNullChar)
        ''''''''''''''''''''''''''''''''''''''''''
        ' It may have been possible that item in
        ' the dictionary contains a vbNullChar.
        ' Therefore, use UBound to get the
        ' key value, which will necessarily
        ' be the last item of SplitArr.
        ' Then Redim Preserve SplitArr
        ' to UBound - 1 to get rid of the
        ' Key element, and use Join
        ' to reassemble to original value
        ' of the Item.
        '''''''''''''''''''''''''''''''''''''''''
        KeyValue = SplitArr(UBound(SplitArr))
        ReDim Preserve SplitArr(LBound(SplitArr) To UBound(SplitArr) - 1)
        ItemValue = Join(SplitArr, vbNullChar)
        '''''''''''''''''''''''''''''''''''''''
        ' Join will set ItemValue to a string
        ' regardless of what the original
        ' data type was. Test the VTypes(Ndx)
        ' value to convert ItemValue back to
        ' the proper data type.
        '''''''''''''''''''''''''''''''''''''''
        Select Case VTypes(Ndx)
            Case vbBoolean
                ItemValue = CBool(ItemValue)
            Case vbByte
                ItemValue = CByte(ItemValue)
            Case vbCurrency
                ItemValue = CCur(ItemValue)
            Case vbDate
                ItemValue = CDate(ItemValue)
            Case vbDecimal
                ItemValue = CDec(ItemValue)
            Case vbDouble
                ItemValue = CDbl(ItemValue)
            Case vbInteger
                ItemValue = CInt(ItemValue)
            Case vbLong
                ItemValue = CLng(ItemValue)
            Case vbSingle
                ItemValue = CSng(ItemValue)
            Case vbString
                ItemValue = CStr(ItemValue)
            Case Else
                ItemValue = ItemValue
        End Select
        ''''''''''''''''''''''''''''''''''''''
        ' Finally, add the Item and Key to
        ' our TempDict dictionary.
        
        TempDict.Add key:=KeyValue, Item:=ItemValue
    Next Ndx
End If


'''''''''''''''''''''''''''''''''
' Set the passed in Dict object
' to our TempDict object.
'''''''''''''''''''''''''''''''''
Set Dict = TempDict
End Sub


