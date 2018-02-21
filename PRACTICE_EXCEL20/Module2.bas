Attribute VB_Name = "Module2"
Option Base 1

Sub Macro1()
    '
    ' Macro1 Macro
    '
    '
    Range("F4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Sub


Sub GetThemeColors()
    Dim tColorScheme As ThemeColorScheme
    Dim colorArray(10) As Variant
    Dim i As Long
    Dim r As Long
    
    Set tColorScheme = ActiveWorkbook.Theme.ThemeColorScheme
    For i = 1 To 10
      colorArray(i) = tColorScheme.Colors(i).RGB
      ActiveSheet.Cells(i, 1).Value = colorArray(i)
    Next i
    i = 0
    For r = 1 To 10
      ActiveSheet.Cells(r, 2).Interior.Color = colorArray(i + 1)
      i = i + 1
    Next r
End Sub

Sub ApplyThemeColors()
    Dim i As Integer
    
    For i = 1 To 10
       ActiveSheet.Cells(i, 3).Interior.ThemeColor = i
       ActiveSheet.Cells(i, 4).Value = i
    Next i
End Sub





