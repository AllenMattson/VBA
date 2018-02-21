Attribute VB_Name = "Module2"
Option Explicit

Sub LoopFillRange()
'   Fill a range by looping through cells

    Dim CellsDown As Long, CellsAcross As Long
    Dim CurrRow As Long, CurrCol As Long
    Dim StartTime As Double
    Dim CurrVal As Long

'   Get the dimensions
    CellsDown = InputBox("How many cells down?")
    If CellsDown = 0 Then Exit Sub
    CellsAcross = InputBox("How many cells across?")
    If CellsAcross = 0 Then Exit Sub
    
'   Record starting time
    StartTime = Timer

'   Loop through cells and insert values
    CurrVal = 1
    Application.ScreenUpdating = False
    For CurrRow = 1 To CellsDown
        For CurrCol = 1 To CellsAcross
            Range("A1").Offset(CurrRow - 1, CurrCol - 1).Value = CurrVal
            CurrVal = CurrVal + 1
        Next CurrCol
    Next CurrRow

'   Display elapsed time
    Application.ScreenUpdating = True
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
End Sub


Sub ArrayFillRange()
'   Fill a range by transferring an array

    Dim CellsDown As Long, CellsAcross As Long
    Dim i As Long, j As Long
    Dim StartTime As Double
    Dim TempArray() As Double
    Dim TheRange As Range
    Dim CurrVal As Long

'   Get the dimensions
    CellsDown = InputBox("How many cells down?")
    If CellsDown = 0 Then Exit Sub
    CellsAcross = InputBox("How many cells across?")
    If CellsAcross = 0 Then Exit Sub

 '  Record starting time
    StartTime = Timer

'   Redimension temporary array
    ReDim TempArray(1 To CellsDown, 1 To CellsAcross)

'   Set worksheet range
    Set TheRange = Range(Cells(1, 1), Cells(CellsDown, CellsAcross))

'   Fill the temporary array
    CurrVal = 0
    Application.ScreenUpdating = False
    For i = 1 To CellsDown
        For j = 1 To CellsAcross
            TempArray(i, j) = CurrVal
            CurrVal = CurrVal + 1
        Next j
    Next i

'   Transfer temporary array to worksheet
    TheRange.Value = TempArray

'   Display elapsed time
    Application.ScreenUpdating = True
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
End Sub
