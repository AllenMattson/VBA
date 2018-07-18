Attribute VB_Name = "WorksheetFormatting"
Sub WhatsInACell()
Attribute WhatsInACell.VB_Description = "Indicates the contents of the underlying cells: text, numbers, and formulas."
Attribute WhatsInACell.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' WhatsInACell Macro
' Indicates the contents of the underlying cells: text, numbers, and formulas.
'

'
    
    Range("A1").Select
    ' Find and format cells containing text
    Selection.SpecialCells(xlCellTypeConstants, 2).Select
    Selection.Style = "20% - Accent4"
    Range("B2").Select
    ' Find and format cells containing numbers
    Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.Style = "Neutral"
    Range("E2").Select
    ' Find and format cells containing formulas
    Selection.SpecialCells(xlCellTypeFormulas, 23).Select
    Selection.Style = "Calculation"
    ' Create a legend
    Range("A1:A4").EntireRow.Insert
    Range("A1").Select
    Selection.Style = "20% - Accent4"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Text"
    Range("A2").Select
    Selection.Style = "Neutral"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Numbers"
    Range("A3").Select
    Selection.Style = "Calculation"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Formulas"
    Range("B1:B3").Select
    With Selection.Font
        .Name = "Arial Narrow"
        .FontStyle = "Bold Italic"
        .Size = 10
    End With
End Sub

Sub RemoveFormats()
Attribute RemoveFormats.VB_ProcData.VB_Invoke_Func = "L\n14"
'
' RemoveFormats Macro
'

'
    Cells.Select
    Range("C2").Activate
    Selection.ClearFormats
    Range("A1:A4").Select
    Selection.EntireRow.Delete
    Range("A1").Select
End Sub





Sub PrintView()
    ActiveWindow.View = xlPageBreakPreview
    ActiveWorkbook.SaveAs "C:\Excel2013_ByExample\CopyPractice_Excel01.xlsm"
End Sub



















