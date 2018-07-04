Attribute VB_Name = "CombineAllTables"
Option Explicit
Sub CombineAllTables()
    Dim J As Integer
    On Error Resume Next
    Sheets("All").Select
    ' copy headings
    Sheets("Sheet1").Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets("All").Range("A1")
    ' work through sheets
    For J = 1 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(J).Activate ' make the sheet active
        Range("A1").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets
        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets("All").Range("A65536").End(xlUp)(2)
    Next
End Sub
