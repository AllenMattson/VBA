Attribute VB_Name = "Module1"
Option Explicit

Function CELLTYPE(Rng)
'   Returns the cell type of the upper left
'   cell in a range
    Dim TheCell As Range
    Set TheCell = Rng.Range("A1")
    Select Case True
        Case IsEmpty(TheCell)
            CELLTYPE = "Blank"
        Case TheCell.NumberFormat = "@"
            CELLTYPE = "Text"
        Case Application.IsText(TheCell)
            CELLTYPE = "Text"
        Case Application.IsLogical(TheCell)
            CELLTYPE = "Logical"
        Case Application.IsErr(TheCell)
            CELLTYPE = "Error"
        Case IsDate(TheCell)
            CELLTYPE = "Date"
        Case InStr(1, TheCell.Text, ":") <> 0
            CELLTYPE = "Time"
        Case IsNumeric(TheCell)
            CELLTYPE = "Number"
    End Select
End Function

