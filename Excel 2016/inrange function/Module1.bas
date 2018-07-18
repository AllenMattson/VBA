Attribute VB_Name = "Module1"
Option Explicit

Function InRange(rng1, rng2) As Boolean
'   Returns True if rng1 is a subset of rng2
    On Error GoTo ErrHandler
    If Union(rng1, rng2).Address = rng2.Address Then
        InRange = True
        Exit Function
    End If
ErrHandler:
    InRange = False
End Function

Sub Test()
    Dim r1 As Range, r2 As Range
    Set r1 = Range("c1")
    Set r2 = Sheets(2).Range("c1:D3")
    MsgBox InRange(r1, r2)
End Sub
