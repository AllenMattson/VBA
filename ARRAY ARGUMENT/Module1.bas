Attribute VB_Name = "Module1"
Option Explicit

Function SumArray(List) As Double
    Dim Item As Variant
    SumArray = 0
    For Each Item In List
        If WorksheetFunction.IsNumber(Item) Then _
            SumArray = SumArray + Item
    Next Item
End Function

Sub MakeList()
    Dim Nums(1 To 100) As Double
    Dim i As Integer
    For i = 1 To 100
        Nums(i) = Rnd * 1000
    Next i
    MsgBox SumArray(Nums)
End Sub



