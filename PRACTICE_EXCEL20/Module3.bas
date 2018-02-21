Attribute VB_Name = "Module3"
Option Explicit

Sub Themes4Thru10()
    Dim tintshade As Variant
    Dim heading As Variant
    Dim cell As Range
    Dim themeC As Integer
    Dim r As Integer
    Dim c As Integer
    Dim i As Integer
    
    heading = Array("ThemeColorIndex", "Neutral", "Lighter 80%", _
             "Lighter 60%", "Lighter 40%", "Darker 25%", "Darker 50%")
    tintshade = Array(0, 0.8, 0.6, 0.4, -0.25, -0.5)
    
    i = 0
    For Each cell In Range("A1:G1")
        cell.Formula = heading(i)
        i = i + 1
    Next
    
    For r = 2 To 8
       themeC = r + 2
         For c = 1 To 7
           If c = 1 Then
             Cells(r, c).Formula = themeC
           Else
               With Cells(r, c)
                   With .Interior
                     .ThemeColor = themeC
                     .TintAndShade = tintshade(c - 2)
                   End With
               End With
           End If
       Next c
    Next r
    ActiveSheet.Columns("A:G").AutoFit
End Sub



