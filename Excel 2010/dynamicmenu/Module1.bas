Attribute VB_Name = "Module1"
Option Explicit

Public MyRibbon As IRibbonUI

'Callback for customUI.onLoad
Sub ribbonLoaded(ribbon As IRibbonUI)
    Set MyRibbon = ribbon
End Sub

'Callback for DynamicMenu getContent
Sub dynamicMenuContent(control As IRibbonControl, ByRef returnedVal)
'   This procedure is executed whenever a sheet is activated
'   (See the Worksheet_Activate procedure in ThisWorkbook)
    
    Dim r As Long
    Dim XMLcode As String
    
'   Read the XML markup from the active sheet
    For r = 1 To Application.CountA(Range("A:A"))
        XMLcode = XMLcode & ActiveSheet.Cells(r, 1) & " "
    Next r
    returnedVal = XMLcode
    
    'Debug.Print xml
End Sub

Sub UpdateDynamicRibbon()
'   Invalidate the ribbon to force a call to dynamicMenuContent
    On Error Resume Next
    MyRibbon.Invalidate
    If Err.Number <> 0 Then
        MsgBox "Lost the Ribbon object. Save and reload."
    End If
End Sub

Sub Macro1(control As IRibbonControl)
'   Executed when Sheet1 is active
    MsgBox "Greetings from the Sheet1 Macro", vbInformation
End Sub

Sub Macro2(control As IRibbonControl)
'   Executed when Sheet2 is active
    MsgBox "Hello from the Sheet2 Macro", vbCritical
End Sub

Sub Macro3(control As IRibbonControl)
'   Executed when Sheet3 is active
    MsgBox "Aloha from the Sheet3 Macro", vbQuestion
End Sub

