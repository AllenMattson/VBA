Attribute VB_Name = "Module1"
Option Explicit

Sub ChangeColorSwatch()
Attribute ChangeColorSwatch.VB_Description = "Macro recorded 6/18/2002 by John Walkenbach"
Attribute ChangeColorSwatch.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim GrayScale
    ActiveSheet.Range("RedSwatch").Interior.Color = RGB(Range("RedValue"), 0, 0)
    ActiveSheet.Range("GreenSwatch").Interior.Color = RGB(0, Range("GreenValue"), 0)
    ActiveSheet.Range("BlueSwatch").Interior.Color = RGB(0, 0, Range("BlueValue"))
    ActiveSheet.Range("CombinedColor").Interior.Color = RGB(Range("RedValue"), Range("GreenValue"), Range("BlueValue"))
    GrayScale = Range("RedValue") * 0.287 + Range("GreenValue") * 0.589 + Range("BlueValue") * 0.114
    ActiveSheet.Range("GrayScaleSwatch").Interior.Color = RGB(GrayScale, GrayScale, GrayScale)
End Sub

