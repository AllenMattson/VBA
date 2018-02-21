Attribute VB_Name = "Breaks"
Option Explicit

Sub ChangeCode()
    Workbooks.Open Filename:="C:\Excel2013_ByExample\Codes.xlsx"
    Windows("Practice_Excel10.xlsm").Activate
    Columns("D:D").Insert Shift:=xlToRight
    Range("D1").Formula = "Code"
    Columns("D:D").SpecialCells(xlBlanks).Select
    ActiveCell.FormulaR1C1 = "=VLookup(RC[1],Codes.xlsx!R1C1:R6C2,2)"
    Selection.FillDown
        With Columns("D:D")
            .EntireColumn.AutoFit
            .Select
        End With
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues
    Rows("1:1").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .Orientation = xlHorizontal
        End With
    Workbooks("Codes.xlsx").Close
End Sub


Sub StopExample()
    Dim curCell As Range
    Dim num As Integer

    ActiveWorkbook.Sheets(1).Select
    ActiveSheet.UsedRange.Select
    num = Selection.Columns.Count
    Selection.Resize(1, num).Select
    Stop
    For Each curCell In Selection
        Debug.Print curCell.Text
    Next
End Sub


Sub TestDebugAssert()
    Dim i As Integer
    
    For i = 1 To 100
        Debug.Assert i <> 50
    Next
End Sub


Sub WhatDate()
    Dim curDate As Date
    Dim newDate As Date
    Dim x As Integer
    
    curDate = Date
    For x = 1 To 365
        newDate = Date + x
    Next
End Sub

Sub MyProcedure()
    Dim strName As String
    Workbooks.Add
    strName = ActiveWorkbook.Name
    ' choose the Step Over to avoid stepping through the
    ' lines of code in the called procedure - SpecialMsg
    SpecialMsg strName
    Workbooks(strName).Close
End Sub

Sub SpecialMsg(n As String)
    If n = "Book2" Then
        MsgBox "You must change the name."
    End If
End Sub
