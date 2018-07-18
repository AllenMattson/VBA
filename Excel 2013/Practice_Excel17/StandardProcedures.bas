Attribute VB_Name = "StandardProcedures"
Option Explicit

Sub EnterData()
    With ActiveSheet.Range("A1:B1")
        .Font.Color = vbRed
        .Value = 15
    End With
    Application.EnableEvents = False
    ActiveWorkbook.Save
    Application.EnableEvents = True
End Sub

