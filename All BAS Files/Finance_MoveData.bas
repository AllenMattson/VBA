Attribute VB_Name = "Finance_MoveData"
Sub Main()
Application.DisplayAlerts = False
Application.Calculation = xlAutomatic
Application.ScreenUpdating = False


Sheets("Sheet1").Activate
Dim LR As Long: LR = Cells(Rows.Count, 1).End(xlUp).Row
Dim LC As Integer: LC = Cells(7, Columns.Count).End(xlToLeft).Column
'Change Date Values to General to maintain consistent transfer of data
'Range("A8:A" & LR).NumberFormat = "General"
'Move Range to New Sheet seperating year, month, day and compiling date
Dim MyRNG As Range: Set MyRNG = Range(Cells(7, "B"), Cells(LR, LC))
MyRNG.Copy
Sheets.Add: Cells(5, 5).PasteSpecial xlPasteAll
Dim Nlr As Long: Nlr = Cells(Rows.Count, 5).End(xlUp).Row
Cells(1, 1) = "Asset 1": Cells(2, 1).Value = "Asset 2": Cells(3, 1) = "Start Date": Cells(4, 1) = "End Date"
Cells(5, 1) = "Year": Cells(5, 2) = "Month": Cells(5, 3) = "Day": Cells(5, 4) = "Date"

'Enter in Formulas
Range("A6").FormulaR1C1 = "=YEAR(Sheet1!R[2]C1)"
Range("B6").FormulaR1C1 = "=MONTH(Sheet1!R[2]C1)"
Range("C6").FormulaR1C1 = "=DAY(Sheet1!R[2]C1)"
Range("D6").FormulaR1C1 = "=DATE(RC[-3],RC[-2],RC[-1])"
Range("A6:D6").AutoFill Destination:=Range("A6:D" & Nlr)
'Keep Values Only
Range("A6:D" & Nlr).Copy
Range("A6:D" & Nlr).PasteSpecial xlPasteValues
Columns.AutoFit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*********************REQUEST USER INPUT BOXES B1:B4***************************'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Add Asset Drop Downs
Dim NLC As Integer: NLC = Range("E5").End(xlToRight).Column
Dim Arng As Range: Set Arng = Range("E5", Cells(5, NLC))
Application.CutCopyMode = False
ActiveWorkbook.Names.Add Name:="assets", RefersTo:=Arng
With Range("B1:B2").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=assets")
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With
'Start by displaying max and min dates, seperate year, month, day
'For seasonal trend forcasting we want individual values as well
Range("B3").FormulaR1C1 = "=MAX(R[3]C[2]:R[6199]C[2])"
Range("B4").FormulaR1C1 = "=MIN(R[2]C[2]:R[6198]C[2])"
Range("B1:G1").Merge: Range("B2:G2").Merge: Range("B3:D3").Merge: Range("B4:D4").Merge
Range("E3").FormulaR1C1 = "=YEAR(RC[-3])"
Range("F3").FormulaR1C1 = "=CHOOSE(MONTH(RC[-4]),""Jan"",""Feb"",""Mar"",""Apr"",""May"",""Jun"",""Jul"",""Aug"",""Sep"",""Oct"",""Dec"")"
Range("F4").FormulaR1C1 = "=CHOOSE(MONTH(RC[-4]),""Jan"",""Feb"",""Mar"",""Apr"",""May"",""Jun"",""Jul"",""Aug"",""Sep"",""Oct"",""Dec"")"
Range("G3").FormulaR1C1 = "=DAY(RC[-5])"
Range("E3:G3").AutoFill Destination:=Range("E3:G4")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Format the coloring of COLUMN A TO LEFT OF input boxes
With Range("A1:A4").Interior
    .Pattern = xlSolid
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.349986266670736
End With
'ADD BORDERS
With Range("A1:A4").Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
'Format coloring of input boxes
With Range("B1:G2").Interior
    .Pattern = xlSolid
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.149998474074526
End With
'ADD BORDERS
With Range("B1:G2", "B3:B4").Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = xlAutomatic
End With
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
CreateTwoAxisChart
End Sub
Sub CreateTwoAxisChart()
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False

'Clear old Charts
Dim CHobj As ChartObject
For Each CHobj In ActiveSheet.ChartObjects
    CHobj.Delete
Next CHobj

'Set up Variables
Dim ShName As String
Dim Str1 As String: Str1 = Range("B1").Value
Dim Str2 As String: Str2 = Range("B2").Value
Dim Rng1 As Range
Dim Rng2 As Range
Dim DateRNG As Range
Dim i As Integer
Dim LC As Integer: LC = Cells(5, Columns.Count).End(xlToRight).Column
Dim LR As Long: LR = Cells(Rows.Count, 1).End(xlUp).Row
Set DateRNG = Range("D6:D" & LR)
'Locate Series to add to chart
For i = 5 To LC
    If Cells(5, i).Value = Str1 Then
        Set Rng1 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
    Else
        If Cells(5, i) = Str2 Then
            Set Rng2 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
        End If
    End If
Next i


With ActiveSheet
    ShName = .Name
End With
Charts.Add
With ActiveChart
    .ChartType = xlLine
    .SetSourceData Source:=Rng1
    .FullSeriesCollection(1).Name = Str1
    .FullSeriesCollection(1).XValues = DateRNG
    With .SeriesCollection.NewSeries
            .Values = Rng2
            .XValues = DateRNG
            .Name = Str2
            .AxisGroup = 2
    End With
    .Location Where:=xlLocationAsObject, Name:=ShName
End With

Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub
Sub Insert_Y2Series()
Dim Cell As Range, MyAssets As Range
Set MyAssets = Range("E5:J5")
Dim CHobj As ChartObject
'for each chobj in thi
ActiveSheet.ChartObjects("Chart 41").Activate
    ActiveChart.FullSeriesCollection(2).Name = "=$B$2"
    For Each Cell In MyAssets
        If Cell.Value = Range("B2").Value Then
            Dim MyVals As Range
            MyVals = Range(Cell.Offset(1, 0), Cells(Rows.Count, Cell.Column))
            With ActiveChart
                .FullSeriesCollection(2).Values
                .FullSeriesCollection(2).XValues = "=Sheet46!$D$6:$D$6202"
                .FullSeriesCollection(2).AxisGroup = 2
            End With
        End If
    Next Cell
End Sub
