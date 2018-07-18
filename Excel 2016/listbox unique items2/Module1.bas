Attribute VB_Name = "Module1"
Option Explicit

Sub RemoveDuplicates()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As Collection, NoDupesSorted As Collection
    Dim i As Long, j As Long
    Dim Item As Variant
    
    Set NoDupes = New Collection
    
'   The items are in a range named Countries
    Set AllCells = Sheet1.Range("Countries")
    
'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.Add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Update the labels on UserForm1
    With UserForm1
        .Label1.Caption = "Total Items: " & AllCells.Count
        .Label2.Caption = "Unique Items: " & NoDupes.Count
    End With
    
    Set NoDupesSorted = New Collection
    
'   Sort the collection (optional)
'   Add the first item
    NoDupesSorted.Add NoDupes(1), CStr(NoDupes(1))
'   Add subsequent items in the correct order
    For i = 2 To NoDupes.Count
        For j = 1 To NoDupesSorted.Count
            If NoDupes(i) < NoDupesSorted(j) Then
                NoDupesSorted.Add NoDupes(i), CStr(NoDupes(i)), j
                'Once we found the right place, move to the next i
                Exit For
            End If
        Next j
    Next i
    
'   Add the sorted, non-duplicated items to a ListBox
    For Each Item In NoDupesSorted
        UserForm1.ListBox1.AddItem Item
    Next Item

'   Show the UserForm
    UserForm1.Show
End Sub

