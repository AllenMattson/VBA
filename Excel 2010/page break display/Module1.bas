Attribute VB_Name = "Module1"
Public MyRibbon As IRibbonUI

Sub Initialize(Ribbon As IRibbonUI)
'   Executed when the workbook loads
    Set MyRibbon = Ribbon
End Sub

Sub TogglePageBreakDisplay(control As IRibbonControl, pressed As Boolean)
'   Executed when Checkbox is clicked
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = pressed
End Sub

Sub CheckPageBreakDisplay()
'   Executed when a sheet is activated
    MyRibbon.InvalidateControl ("Checkbox1")
End Sub

Sub GetPressed(control As IRibbonControl, ByRef returnedVal)
'   Executed when the control is invalidated
    On Error Resume Next
    returnedVal = ActiveSheet.DisplayPageBreaks
End Sub

Sub GetEnabled(control As IRibbonControl, ByRef returnedVal)
'   Executed when the control is invalidated
    returnedVal = TypeName(ActiveSheet) = "Worksheet"
End Sub
