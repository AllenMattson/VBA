Attribute VB_Name = "Module1"
Option Explicit
'   This example is based on a tip by J.G. Hussey,
'   published in "Visual Basic Programmer's Journal"

Sub RemoveDuplicates()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim i As Integer, j As Integer
    Dim Swap1, Swap2, Item
    
'   The items are in a range named Countries
    Set AllCells = Range("Countries")
    
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
    
'   Sort the collection (optional)
    For i = 1 To NoDupes.Count - 1
        For j = i + 1 To NoDupes.Count
            If NoDupes(i) > NoDupes(j) Then
                Swap1 = NoDupes(i)
                Swap2 = NoDupes(j)
                NoDupes.Add Swap1, before:=j
                NoDupes.Add Swap2, before:=i
                NoDupes.Remove i + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next i
    
'   Add the sorted, non-duplicated items to a ListBox
    For Each Item In NoDupes
        UserForm1.ListBox1.AddItem Item
    Next Item

'   Show the UserForm
    UserForm1.Show
End Sub

