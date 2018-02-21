Attribute VB_Name = "Module1"
Option Explicit

Sub TintAndShadeDemo()
    Dim TaS As Double
    Dim TheColor As Long
    Dim r As Long
    TaS = -1
    TheColor = Sheets(1).Range("StartingColor").Interior.Color
    Application.ScreenUpdating = False
    For r = 5 To 55
        Cells(r, 1).Interior.Color = TheColor
        Cells(r, 1).Interior.TintAndShade = TaS
        Cells(r, 2) = Cells(r, 1).Interior.TintAndShade
        Cells(r, 3) = Cells(r, 1).Interior.Color
        Range(Cells(r, 4), Cells(r, 6)) = DECIMAL2RGB(Cells(r, 1).Interior.Color)
        TaS = Round(TaS + 0.04, 2)
    Next r
End Sub

Function DECIMAL2RGB(ColorVal) As Variant
'   Converts a color value to an RGB triplet
'   Returns a 3-element variant array
    DECIMAL2RGB = Array(ColorVal \ 256 ^ 0 And 255, ColorVal \ 256 ^ 1 And 255, ColorVal \ 256 ^ 2 And 255)
End Function


