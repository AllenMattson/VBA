Attribute VB_Name = "Module1"
Public Sub OpenCalculator(ctl As IRibbonControl)
    Shell "Calc.exe", vbNormalFocus
End Sub

