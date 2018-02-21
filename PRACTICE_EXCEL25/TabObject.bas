Attribute VB_Name = "TabObject"
Option Explicit

Sub ColorTabs()
    Dim wks As Worksheet
    Dim i As Integer

    i = 5

    For Each wks In ThisWorkbook.Worksheets
        If wks.Tab.ColorIndex = xlColorIndexNone Then
            wks.Tab.ColorIndex = i
            i = i + 1
        End If
    Next
End Sub


