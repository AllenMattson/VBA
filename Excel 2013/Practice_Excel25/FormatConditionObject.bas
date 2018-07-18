Attribute VB_Name = "FormatConditionObject"
Option Explicit

Sub ApplyConditionalFormat()
    Dim objFormatCon As FormatCondition
    Dim objFormatColl As FormatConditions
    Dim myRange As Range

    ' select range containing numeric cells only
    Set myRange = ActiveSheet.UsedRange. _
        SpecialCells(xlCellTypeConstants, 1)
    Set objFormatColl = myRange.FormatConditions

    ' find out if any conditional formatting already exists
    If objFormatColl.Count > 0 Then
        MsgBox "There are " & objFormatColl.Count & " conditions defined for the used range."
    End If

    ' remove existing conditions if they exist
    myRange.FormatConditions.Delete
 
    ' add first condition
    Set objFormatCon = objFormatColl.Add(Type:=xlCellValue, _
        Operator:=xlGreaterEqual, _
        Formula1:="150")
    With objFormatCon
        .Font.Bold = True
        .Font.ColorIndex = 2 ' white
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(0, 0, 255) ' blue
    End With
End Sub


