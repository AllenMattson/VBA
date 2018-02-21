Attribute VB_Name = "Module1"
Option Explicit

Function CellType(Rng)
'   Returns the cell type of the upper left
'   cell in a range
    Dim TheCell As Range
    Set TheCell = Rng.Range("A1")
    Select Case True
        Case IsEmpty(TheCell)
            CellType = "Blank"
        Case TheCell.NumberFormat = "@"
            CellType = "Text"
        Case Application.IsText(TheCell)
            CellType = "Text"
        Case Application.IsLogical(TheCell)
            CellType = "Logical"
        Case Application.IsErr(TheCell)
            CellType = "Error"
        Case IsDate(TheCell)
            CellType = "Date"
        Case InStr(1, TheCell.Text, ":") <> 0
            CellType = "Time"
        Case IsNumeric(TheCell)
            CellType = "Number"
    End Select
End Function

