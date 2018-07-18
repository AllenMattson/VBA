Attribute VB_Name = "Module1"
Option Explicit

Sub GetData()
    Dim NextRow As Long
    Dim Entry1 As String, Entry2 As String
  Do
    'Determine next empty row
    NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
'   Prompt for the data
    Entry1 = InputBox("Enter the name")
    If Entry1 = "" Then Exit Sub
    Entry2 = InputBox("Enter the amount")
    If Entry2 = "" Then Exit Sub
    
'   Write the data
    Cells(NextRow, 1) = Entry1
    Cells(NextRow, 2) = Entry2
  Loop
End Sub

