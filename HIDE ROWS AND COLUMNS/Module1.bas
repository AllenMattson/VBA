Attribute VB_Name = "Module1"
Option Explicit

Sub HideRowsAndColumns()
Attribute HideRowsAndColumns.VB_ProcData.VB_Invoke_Func = "H\n14"
    Dim row1 As Long, row2 As Long
    Dim col1 As Long, col2 As Long
        
    If TypeName(Selection) <> "Range" Then Exit Sub
    
'   If last row or last column is hidden, unhide all and quit
    If Rows(Rows.Count).EntireRow.Hidden Or Columns(Columns.Count).EntireColumn.Hidden Then
        Cells.EntireColumn.Hidden = False
        Cells.EntireRow.Hidden = False
        Exit Sub
    End If
    
    row1 = Selection.Rows(1).Row
    row2 = row1 + Selection.Rows.Count - 1
    col1 = Selection.Columns(1).Column
    col2 = col1 + Selection.Columns.Count - 1
    
    Application.ScreenUpdating = False
    On Error Resume Next
'   Hide rows
    Range(Cells(1, 1), Cells(row1 - 1, 1)).EntireRow.Hidden = True
    Range(Cells(row2 + 1, 1), Cells(Rows.Count, 1)).EntireRow.Hidden = True
'   Hide columns
    Range(Cells(1, 1), Cells(1, col1 - 1)).EntireColumn.Hidden = True
    Range(Cells(1, col2 + 1), Cells(1, Columns.Count)).EntireColumn.Hidden = True
End Sub

