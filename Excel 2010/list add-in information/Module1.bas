Attribute VB_Name = "Module1"
Option Explicit

Sub ListAllAddins()
    Dim ai As AddIn
    Dim Row As Long
    Dim Table1 As ListObject
    Cells.Clear
    Application.ScreenUpdating = False
    Range("A1:E1") = Array("Name", "Title", "Installed", _
      "Comments", "Path")
    Row = 2
    On Error Resume Next
    For Each ai In Application.AddIns
        Cells(Row, 1) = ai.Name
        Cells(Row, 2) = ai.Title
        Cells(Row, 3) = ai.Installed
        Cells(Row, 4) = ai.Comments
        Cells(Row, 5) = ai.Path
        Row = Row + 1
    Next ai
    On Error GoTo 0
    Range("A1").Select
    ActiveSheet.ListObjects.Add
    ActiveSheet.ListObjects(1).TableStyle = _
      "TableStyleMedium2"
End Sub
