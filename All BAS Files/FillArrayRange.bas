Attribute VB_Name = "FillArrayRange"
Sub FillArrayRange()
Dim WS As Worksheet: Set WS = ActiveSheet
Dim NewWS As Worksheet
'Fill range by transferring array
Dim CellsDown As Long, CellsAcross As Long, StartCol As Integer
Dim i As Long, j As Long, k As Long
Dim TempArray() As Long
Dim TheRange As Range
Dim CurrVal As Long
'get dimensions
CellsDown = Cells(Rows.Count, 1).End(xlUp).Row
If CellsDown = 0 Then Exit Sub

CellsAcross = Cells(1, Columns.Count).End(xlToLeft).Column
For k = 1 To CellsAcross
        If Mid(Cells(1, k), 5, 1) = "-" Then StartCol = k
Next k
'Redimension temp array
ReDim TempArray(2 To CellsDown, StartCol To CellsAcross)
'set worksheet range
Set NewWS = Sheets.Add
Set TheRange = ActiveCell.Range(Cells(1, 1), Cells(CellsDown, CellsAcross))

'Fill temp array
CurrVal = 0
Application.ScreenUpdating = False
For i = 2 To CellsDown
    For j = StartCol To CellsAcross
        TempArray(i, j) = CurrVal + 1
        CurrVal = CurrVal + 1
    Next j
Next i

'Transfer temp array to worksheet
TheRange.Value = TempArray
Application.ScreenUpdating = True
ActiveSheet.Range("A1").End(xlDown).EntireRow.Delete
End Sub

