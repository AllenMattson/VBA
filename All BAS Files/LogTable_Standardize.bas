Attribute VB_Name = "LogTable_Standardize"
Sub LogTable()
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False
Sheets("test").Select
'After gathering all stocks, this will build a historical volatility table for each asset

With Sheets("test")
    Dim LR As Long, LC As Long, LCdata As Long
    LR = Cells(Rows.Count, 1).End(xlUp).Row - 1
    LC = Cells(5, Columns.Count).End(xlToLeft).Column
    LCdata = Cells(5, Columns.Count).End(xlToLeft).Column
End With
Application.CutCopyMode = False

'Logarithmic difference between current and previous day
'Calculates historical volatility
Dim Rng As Range: Set Rng = Range(Cells(6, LC).Offset(0, 1), Cells(LR, LC).Offset(0, LC - 4)) 'subtract 4 for year month day date columns
Cells(6, LC).Offset(0, 1).FormulaR1C1 = "=LN(R[1]C[-6]/RC[-6])"
Cells(6, LC).Offset(0, 1).AutoFill Destination:=Range(Cells(6, LC).Offset(0, 1), Cells(6, LC).Offset(0, LC - 4)), Type:=xlFillDefault
Range(Cells(6, LC).Offset(0, 1), Cells(6, LC).Offset(0, LC - 4)).AutoFill Destination:=Range(Cells(6, LC).Offset(0, 1), Cells(LR, LC).Offset(0, LC - 4)), Type:=xlFillDefault

'Logarithmic difference between current and previous day
'Calculates historical volatility
With Rng
    .Calculate
    .NumberFormat = "0.00%"
    .Cells.Copy
    .PasteSpecial xlPasteValues
End With
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub
Sub NameVolRanges()
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False
Dim NLastCol As Long, LastRo As Long, Col_Headers As Integer, i As Integer
Dim myRANGE As Range
Dim MyStr As String
'delete named ranges
Dim sName As Name
For Each sName In ThisWorkbook.Names
    If InStr(1, sName, "test") Then
        sName.Delete
    End If
Next
Sheets("test").Select


NLastCol = Cells(5, Columns.Count).End(xlToLeft).Column 'Cells.Find(what:="*", after:=[A1], searchorder:=xlByColumns, searchdirection:=xlPrevious).Column




Col_Headers = Cells(6, Columns.Count).End(xlToLeft).Column
LastRo = Cells(Rows.Count, NLastCol).Offset(0, 1).End(xlUp).Row


Cells(LastRo, "D").Offset(3, 0) = "Mean"
Cells(LastRo, "D").Offset(4, 0) = "Std Dev"
For i = 5 To Col_Headers - 1

Set myRANGE = Range(Cells(6, i), Cells(LastRo, i))
Dim FirstSpace As Integer: FirstSpace = InStr(Cells(5, i).Value, " ")
If FirstSpace = 0 Then FirstSpace = Len(Cells(5, i).Offset(0, -5))
    MyStr = Left(Cells(5, i), FirstSpace)
    myRANGE.Select
    'Insert Named Range
    On Error Resume Next
    ActiveWorkbook.Names.Add Name:=MyStr, RefersTo:=myRANGE
    'Find average and standard dev to normalize
    If i <= NLastCol Then
        Range("A5").End(xlToRight).Offset(0, 1).Value = MyStr
        Cells(LastRo, i).Offset(3, 0) = Application.WorksheetFunction.Average(myRANGE)
        Cells(LastRo, i).Offset(4, 0) = Application.WorksheetFunction.StDev_P(myRANGE)
    End If
Next
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True


Dim NumCol As Integer: NumCol = Range("A5").End(xlToRight).Row
Dim j As Integer: j = 0
While j <= NumCol
    Standardize_Method
j = j + 1
Wend
'Fill Down Selection
Range(Range("A5").End(xlToRight).Offset(1, 1), Range("A6").End(xlToRight)).Select
Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
End Sub
Sub Standardize_Method()
Application.DisplayAlerts = False
Application.Calculation = xlAutomatic
Application.ScreenUpdating = False
Application.CutCopyMode = False
Dim LRow As Long: LRow = Cells(Rows.Count, "D").End(xlUp).Row
Dim TGT As Range: Set TGT = Range("A6").End(xlToRight).Offset(0, 1): TGT.Offset(0, j).FormulaR1C1 = "=STANDARDIZE(RC[-12],R" & LRow & "C[-12],R" & LRow & "C[-12])"
'TGT.Offset(0, j).AutoFill Destination:=Range(TGT, TGT.End(xlDown))

Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub
