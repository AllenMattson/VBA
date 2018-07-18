Attribute VB_Name = "Module1"
Option Explicit

Sub FormatUsedRange()
    ActiveSheet.UsedRange.Select
    Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.NumberFormat = "[<150][Red];[>250][Green];[Yellow]"
   ' Selection.NumberFormat = "[<150][Red];[>250][Color 32];[Yellow]"

End Sub

Sub FormatVariable()
    Dim myResult, frmResult
    myResult = "1435.60"
    frmResult = Format(myResult, "Currency")
    Debug.Print frmResult
    ActiveSheet.Range("G1").FormulaR1C1 = frmResult
End Sub

Sub FormatDateFields()
    Dim wks As Worksheet
    Dim cell As Range
        
    Set wks = ActiveWorkbook.ActiveSheet
    
    For Each cell In wks.UsedRange
      If cell.NumberFormat = "mm:ss.0" Then
        cell.NumberFormat = "m/dd/yyyy h:mm:ss AM/PM"
      End If
    Next
End Sub

Sub ColorLoop()
    Dim r As Integer
    Dim c As Integer
    Dim k As Integer

    k = 0

    For r = 1 To 8
        For c = 1 To 7
            Cells(r, c).Select
            k = k + 1
            ActiveCell.Value = k
            With Selection.Interior
                .ColorIndex = k
                .Pattern = xlSolid
            End With
        Next c
    Next r
End Sub


Sub ApplyCellFormat()
    With ActiveSheet.Range("A1").Font
        .Name = "Tahoma"
        .FontStyle = "italic"
        .Size = 14
        .Underline = xlUnderlineStyleDouble
        .ColorIndex = 3
    End With
End Sub


Sub FormatQtrText()
    With ActiveSheet.UsedRange
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlTextString, String:="Qtr", _
            TextOperator:=xlContains
        .FormatConditions(1).Interior.Color = RGB(123, 130, 0)
    End With
End Sub

Sub AddAboveAverageFormat()
    With Selection
        .FormatConditions.Delete
        .FormatConditions.AddAboveAverage
        .FormatConditions(1).AboveBelow = xlAboveAverage
        .FormatConditions(1).Font.Bold = True
    End With
End Sub

Sub HighlightAltRows()
    With ActiveSheet.UsedRange
        .FormatConditions.Add Type:=xlExpression, _
                Formula1:="=MOD(ROW(),2)=0"
        .FormatConditions(1).Interior.ColorIndex = 6
    End With
End Sub

Sub HighlightNonEmptyCells()
    Range("A1:B12").Select
    Selection.FormatConditions.Add _
        Type:=xlNoBlanksCondition
    With Selection.FormatConditions(1).Interior
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399945066682943
    End With
End Sub

Sub HighlightCellsWithNoErrors()
    Range("F1:F7").Select
    Selection.FormatConditions.Add _
            Type:=xlNoErrorsCondition
    With Selection.FormatConditions(1).Interior
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399945066682943
    End With
End Sub

Sub FormatWithDataBars()
    With Range("B2:E6").FormatConditions
    .AddDatabar
    .Add Type:=xlDatabar, _
        Operator:=xlGreaterEqual, Formula1:="200"
    End With
End Sub



Sub IconSetRules()
    Dim iSC As IconSetCondition
    
    Columns("C:C").Select
    With Selection
        .SpecialCells(xlCellTypeConstants, 23).Select
        .FormatConditions.Delete
        .NumberFormat = "$#,##0.00"
        Set iSC = Selection.FormatConditions.AddIconSetCondition
        iSC.IconSet = ActiveWorkbook.IconSets(xl3Symbols)
    End With
End Sub

Sub IconSetRulesRevised()
    Dim iSC As IconSetCondition
    
    Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    With Selection
        .FormatConditions.Delete
        .AutoFilter
        .NumberFormat = "$#,##0.00"
        Set iSC = Selection.FormatConditions.AddIconSetCondition
        iSC.IconSet = ActiveWorkbook.IconSets(xl3Symbols)
        With iSC.IconCriteria(2)
            .Type = xlConditionValueNumber
            .Value = 50000
            .Operator = xlGreaterEqual
        End With
        
        With iSC.IconCriteria(3)
            .Type = xlConditionValueNumber
            .Value = 80000
            .Operator = xlGreaterEqual
        End With
       
       .AutoFilter Field:=1, Criteria1:=iSC.IconSet.Item(3), _
               Operator:=xlFilterIcon
    End With
End Sub

Sub IconSetHideIcons()
    Dim iSC As IconSetCondition

    Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    With Selection
        .FormatConditions.Delete
        .NumberFormat = "$#,##0.00"
        Set iSC = Selection.FormatConditions.AddIconSetCondition
        iSC.IconSet = ActiveWorkbook.IconSets(xl3Symbols)
        .FormatConditions(1).IconCriteria(1).Icon = xlIconRedCrossSymbol
       
        With iSC.IconCriteria(2)
            .Type = xlConditionValueNumber
            .Value = 50000
            .Operator = xlGreaterEqual
            .Icon = xlIconNoCellIcon
        End With

        With iSC.IconCriteria(3)
            .Type = xlConditionValueNumber
            .Value = 80000
            .Operator = xlGreaterEqual
            .Icon = xlIconNoCellIcon
        End With

    End With
End Sub
