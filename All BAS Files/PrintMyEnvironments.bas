Attribute VB_Name = "PrintMyEnvironments"
Option Explicit
Sub PrintMyEnvironments()
Cells.Clear
Cells(1, 1).Value = "Environment Number"
Cells(1, 2).Value = "Environment"
'running test of i to 100 resulted in only 48 environments
Dim i As Integer
For i = 2 To 48
    Cells(i, 1).Value = i
    Cells(i, 2).Value = Environ(i)
Next
Columns.AutoFit
End Sub

