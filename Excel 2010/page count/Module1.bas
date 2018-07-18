Attribute VB_Name = "Module1"
Option Explicit

Sub ShowPageCount()
    Dim PageCount As Integer
    Dim sht As Worksheet
    PageCount = 0
    For Each sht In Worksheets
       PageCount = PageCount + (sht.HPageBreaks.Count + 1) * _
        (sht.VPageBreaks.Count + 1)
    Next sht
    MsgBox "Total Pages = " & PageCount
End Sub


