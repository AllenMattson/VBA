Attribute VB_Name = "Module1"
Sub TimeTest1()
'   VARIABLES DECLARED
    Dim x As Long, y As Long
    Dim A As Double, B As Double, C As Double
    Dim i As Long, j As Long
    Dim StartTime As Date, EndTime As Date
'   Store the starting time
    StartTime = Timer
'   Perform some calculations
    x = 0
    y = 0
    For i = 1 To 10000
        x = x + 1
        y = x + 1
        For j = 1 To 10000
            A = x + y + i
            B = y - x - i
            C = x / y * i
        Next j
    Next i
'   Get ending time
    EndTime = Timer
'   Display total time in seconds
    MsgBox Format(EndTime - StartTime, "0.0")
End Sub

Sub TimeTest2()
'   VARIABLES NOT DECLARED
'   Store the starting time
    StartTime = Timer
'   Perform some calculations
    x = 0
    y = 0
    For i = 1 To 10000
        x = x + 1
        y = x + 1
        For j = 1 To 10000
            A = x + y + i
            B = y - x - i
            C = x / y * i
        Next j
    Next i
'   Get ending time
    EndTime = Timer
'   Display total time in seconds
    MsgBox Format(EndTime - StartTime, "0.0")
End Sub


