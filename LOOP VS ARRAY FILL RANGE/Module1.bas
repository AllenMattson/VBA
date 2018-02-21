Attribute VB_Name = "Module1"
Option Explicit

Sub WriteReadRange()
    Dim MyArray()
    Dim Time1 As Double
    Dim NumElements As Long, i As Long
    Dim WriteTime As String, ReadTime As String
    Dim Msg As String
    
    NumElements = 60000
    ReDim MyArray(1 To NumElements)

'   Fill the array
    For i = 1 To NumElements
        MyArray(i) = i
    Next i
    
'   Write the array to a range
    Time1 = Timer
    For i = 1 To NumElements
        Cells(i, 1) = MyArray(i)
    Next i
    WriteTime = Format(Timer - Time1, "00:00")
    
'   Read the range into the array
    Time1 = Timer
    For i = 1 To NumElements
        MyArray(i) = Cells(i, 1)
    Next i
    ReadTime = Format(Timer - Time1, "00:00")
    
'   Show results
    Msg = "Write: " & WriteTime
    Msg = Msg & vbCrLf
    Msg = Msg & "Read: " & ReadTime
    MsgBox Msg, vbOKOnly, NumElements & " Elements"
End Sub

