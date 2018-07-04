Attribute VB_Name = "TimeSeriesCorrelationCharts"
Sub RESFRESH_BUTTON()
Application.DisplayAlerts = False
Application.Calculation = xlAutomatic
Application.ScreenUpdating = False
Dim CHobj As ChartObject
Dim TestSH As Worksheet
For Each TestSH In ThisWorkbook.Worksheets
    If TestSH.Name = "test" Then
        TestSH.Activate: Cells.Clear
        'Clear old Charts
        For Each CHobj In ActiveSheet.ChartObjects
            CHobj.Delete
        Next CHobj
    End If
Next TestSH

Sheets("Sheet1").Activate
'Sort Newest To Oldest
Sheets("Sheet1").Select
With ActiveSheet
    .AutoFilter.Sort.SortFields.Clear
    .AutoFilter.Sort.SortFields.Add Key:=Range("A7"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With .AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End With

Dim LR As Long: LR = Cells(Rows.Count, 1).End(xlUp).Row
Dim LC As Integer: LC = Cells(7, Columns.Count).End(xlToLeft).Column

'Move Range to New Sheet seperating year, month, day and compiling date
Dim MyRNG As Range: Set MyRNG = Range(Cells(7, "B"), Cells(LR, LC))
MyRNG.Copy Destination:=Sheets("test").Cells(5, 5)
Sheets("test").Activate
'Sheets.Add: ActiveSheet.Name = "test"
    
'Cells(5, 5).PasteSpecial xlPasteAll
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
Range("K1:ZZ" & LR).Cells.Clear
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'creates log table for chart values
LogTable
'Convert Error Cells
Cells.Replace what:="#N/A", Replacement:="-"
'Name Ranges of Normalized Data
NameVolRanges
'CALL MACRO TO BUILD MATRIX
MATRIXX
ColorMatrix


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

End Sub
Sub LogTable()
'CREATES LOG TABLE OF ASSETS
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
For i = 5 To Col_Headers


Set myRANGE = Range(Cells(6, i), Cells(LastRo, i))
Dim FirstSpace As Integer: FirstSpace = InStr(Cells(5, i).Value, " ")
If FirstSpace = 0 Then FirstSpace = Len(Cells(5, i).Offset(0, -5))
    If i < (Col_Headers - 3) / 2 Then
        MyStr = Left(Cells(5, i).Offset(0, 0), FirstSpace)
        'insert empty cell in top left of correlation matrix
        Range("A6").End(xlToRight).Offset(-1, 1).Cells.Insert shift:=xlToRight
        GoTo LastName
    End If
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
    'setting up correlation matrix
LastName:
    If Range("ZZ5").End(xlToLeft) <> MyStr Then Range("ZZ5").End(xlToLeft).Offset(0, 1) = MyStr
Next
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub

Sub ChangeChartAssets()
Dim TargetChart As Worksheet: Set TargetChart = Sheets("test")
If Cells(1, 2).Value = "" Or Cells(2, 2).Value = "" Then
    MsgBox "You must select two assets to measure in cells B1 and B2", vbOKOnly, "Missing Assets"
    Exit Sub
End If
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False

'Clear old Charts
Dim CHobj As ChartObject
For Each CHobj In ActiveSheet.ChartObjects
    CHobj.Delete
Next CHobj

'Set up Variables
Dim ser As Series
Dim ShName As String
Dim Str1 As String: Str1 = Trim(Range("B1").Value)
Dim Str2 As String: Str2 = Trim(Range("B2").Value)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IF THE ASSET IS COMPARED PERCENTAGES TO VALUES, CALL DIFFERENT MACRO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InStr(Str1, " ") Or InStr(Str2, " ") Then
    CreateTwoAxisChart
    Exit Sub
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RNG1 As Range
Dim RNG2 As Range
Dim DateRNG As Range
Dim i As Integer
Dim LC As Integer: LC = Range("A5").End(xlToRight).Column 'Cells(5, Columns.Count).End(xlToRight).Column
Dim LR As Long: LR = Cells(Rows.Count, 1).End(xlUp).Row
Set DateRNG = Range("D6:D" & LR)
'Locate Series to add to chart
Dim Num_Count As Long: Num_Count = 5
'CREATE LOOP SO ALL SERIES CAN BE MEASURED
While Num_Count <= LC
For i = 5 To LC
    If Trim(Cells(5, i).Value) = Str1 Then    'Series Values
        Set RNG1 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
    Else
        If Trim(Cells(5, i)) = Str2 Then      'Series Values
            Set RNG2 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
        End If
    End If
Next i
Num_Count = Num_Count + 1
Wend
'If values are same, do not update chart, exit sub
If Str1 = Str2 Then
    MsgBox "Values Must Be Different..." & vbNewLine _
    & vbNewLine & "Change Values in Cells B1 or B2", vbOKOnly, "Change an Asset..."
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

With ActiveSheet
    ShName = .Name
End With
'Source of Data
'Dim FullRNG As Range: Set FullRNG = Union(RNG1, Rng2)
Dim source As Range
Set source = Union(DateRNG, RNG1, RNG2) 'FullRNG)


'Set up Chart Elements
Dim AssetChart As Object
Set AssetChart = TargetChart.Shapes.AddChart(xlLine).Chart
    With AssetChart
        .SetSourceData source:=source
        .ChartType = xlLine
        .HasTitle = False
        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .SetElement (msoElementPrimaryValueAxisTitleHorizontal)
        .Axes(xlCategory).CategoryType = xlCategoryScale
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Change"
        .HasLegend = True
        .SetElement (msoElementLegendTop)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(220, 230, 241)
        .ChartArea.Format.Line.Visible = msoFalse
        .Parent.Name = "Performance"
        .SeriesCollection(1).Name = Str1
        .SeriesCollection(2).Name = Str1
        .Location Where:=xlLocationAsObject, Name:="test"
        With .ChartArea
            .Top = [A5].Top
            .Left = [A5].Left
            .Width = 784
            .Height = 510
        End With
End With
'Make chart series skinny
For Each ser In AssetChart.SeriesCollection
ser.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1
    End With
Next ser
AdjustVerticalAxis
End Sub
Sub AdjustVerticalAxis()
Dim cht As ChartObject
Dim srs As Series
Dim FirstTime  As Boolean
Dim MaxNumber As Double
Dim MinNumber As Double
Dim MaxChartNumber As Double
Dim MinChartNumber As Double
Dim Padding As Double

'Input Padding on Top of Min/Max Numbers (Percentage)
  Padding = 0.01  'Number between 0-1

'Optimize Code
  Application.ScreenUpdating = False
  
'Loop Through Each Chart On ActiveSheet
  For Each cht In ActiveSheet.ChartObjects
    
    'First Time Looking at This Chart?
      FirstTime = True
      
    'Determine Chart's Overall Max/Min From Connected Data Source
      For Each srs In cht.Chart.SeriesCollection
        'Determine Maximum value in Series
          MaxNumber = Application.WorksheetFunction.Max(srs.Values)
        
        'Store value if currently the overall Maximum Value
          If FirstTime = True Then
            MaxChartNumber = MaxNumber
          ElseIf MaxNumber > MaxChartNumber Then
            MaxChartNumber = MaxNumber
          End If
        
        'Determine Minimum value in Series (exclude zeroes)
          MinNumber = Application.WorksheetFunction.Min(srs.Values)
          
        'Store value if currently the overall Minimum Value
          If FirstTime = True Then
            MinChartNumber = MinNumber
          ElseIf MinNumber < MinChartNumber Or MinChartNumber = 0 Then
            MinChartNumber = MinNumber
          End If
        
        'First Time Looking at This Chart?
          FirstTime = False
      Next srs
      
    'Rescale Y-Axis
      cht.Chart.Axes(xlValue).MinimumScale = MinChartNumber * (1 - Padding)
      cht.Chart.Axes(xlValue).MaximumScale = MaxChartNumber * (1 + Padding)
  
  Next cht

'Optimize Code
  Application.ScreenUpdating = True

End Sub
Sub MATRIXX()
Dim AssetNames As Range
Cells(6, Columns.Count).End(xlToLeft).Offset(-1, 1).Insert shift:=xlToRight
Range("ZZ5").End(xlToLeft).Select
Set AssetNames = Range(Range("ZZ5").End(xlToLeft), Range("ZZ5").End(xlToLeft).End(xlToLeft))
AssetNames.Copy
Range("A6").End(xlToRight).Offset(0, 1).PasteSpecial xlPasteValues, , , Transpose:=True

Dim MatrixLC As Long: MatrixLC = Cells(5, Columns.Count).End(xlToLeft).Column
Dim MatrixLR As Long: MatrixLR = AssetNames.Cells.Count
Application.CutCopyMode = False
'Fill matrix with the absolute value of the indirect function of the correlation between the named ranges
With Range("R6")
    .FormulaR1C1 = "=ABS(CORREL(INDIRECT(RC17),INDIRECT(R5C)))"
    .AutoFill Destination:=Range(Cells(6, "R"), Cells(6, MatrixLC)), Type:=xlFillDefault
End With
'Add 5 to the matrix last row since the row count started in the 5th row down
Range(Cells(6, "R"), Cells(6, MatrixLC)).AutoFill Destination:=Range(Cells(6, "R"), Cells(MatrixLR + 5, MatrixLC)), Type:=xlFillDefault
End Sub
Sub ColorMatrix()
'Applies special formatting to correlation chart

'FIND LAST COLUMN OF MATRIX
Dim Col As Integer: Col = Cells(5, Columns.Count).End(xlToLeft).Column
Range("A5").End(xlToRight).Offset(1, 1).Select
'FIND NUMBER OF MEASURED ASSETS
Dim RoCount As Integer: RoCount = ActiveCell.End(xlDown).Row

'FIND STARTING COLUMN
Dim col2 As Integer: col2 = Range("A5").End(xlToRight).Offset(1, 2).Column

'NAME MATRIX TABLE RANGE
Dim TBLrange As Range: Set TBLrange = Range(Cells(6, col2), Cells(RoCount, Col))

'FORMAT MATRIX TABLE
With TBLrange
    .FormatConditions.AddColorScale ColorScaleType:=3
    .FormatConditions(TBLrange.FormatConditions.Count).SetFirstPriority
    .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
End With
With TBLrange.FormatConditions(1).ColorScaleCriteria(1).FormatColor
    .Color = 8109667
    .TintAndShade = 0
End With

With TBLrange
    .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    .FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With TBLrange.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
End With
TBLrange.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
With TBLrange.FormatConditions(1).ColorScaleCriteria(3).FormatColor
    .Color = 7039480
    .TintAndShade = 0
End With
End Sub
Sub CreateTwoAxisChart()
If Cells(1, 2).Value = "" Or Cells(2, 2).Value = "" Then
    MsgBox "You must select two assets to measure in cells B1 and B2", vbOKOnly, "Missing Assets"
    Exit Sub
End If
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False

'Clear old Charts
Dim CHobj As ChartObject
For Each CHobj In ActiveSheet.ChartObjects
    CHobj.Delete
Next CHobj

'Set up Variables
Dim ser As Series
Dim ShName As String
Dim Str1 As String: Str1 = Range("B1").Value
Dim Str2 As String: Str2 = Range("B2").Value
Dim RNG1 As Range
Dim RNG2 As Range
Dim DateRNG As Range
Dim i As Integer
Dim LC As Integer: LC = Range("A5").End(xlToRight).Column 'Cells(5, Columns.Count).End(xlToRight).Column
Dim LR As Long: LR = Cells(Rows.Count, 1).End(xlUp).Row
Set DateRNG = Range("D6:D" & LR)
Dim Counter As Long: Counter = 5
'CREATE LOOP SO ALL SERIES CAN BE MEASURED
While Counter <= LC

'Locate Series to add to chart
For i = 5 To LC
    If Cells(5, i).Value = Str1 Then
        Set RNG1 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
    Else
        If Cells(5, i) = Str2 Then
            Set RNG2 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
        End If
    End If
Next i
Counter = Counter + 1
Wend
'If values are same, do not update chart, exit sub
If Str1 = Str2 Then
    MsgBox "Values Must Be Different..." & vbNewLine _
    & vbNewLine & "Change Values in Cells B1 or B2", vbOKOnly, "Change an Asset..."
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

With ActiveSheet
    ShName = .Name
End With
Charts.Add
With ActiveChart
    .ChartType = xlLine
    .SetSourceData source:=Union(DateRNG, RNG1)
    .SeriesCollection(1).Name = Str1
    .SeriesCollection(1).XValues = DateRNG
    With .SeriesCollection.NewSeries
            .Values = RNG2
            .XValues = DateRNG
            .Name = Str2
            .AxisGroup = 2
    End With
    .SetElement (msoElementLegendTop)
    .Location Where:=xlLocationAsObject, Name:=ShName
    With ActiveChart.ChartArea
        .Left = [A6].Left
        .Top = [A6].Top
        .Width = 800
        .Height = 600
    End With

For Each ser In ActiveChart.SeriesCollection
ser.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1
    End With
Next ser
    
End With

'The end result adjusts Y2 correctly, below the macro call, Seriescollection(1) will need to be adjusted as well
AdjustVerticalAxis

'Series(1) adjustments, Changes Collections to display max min scale correctly
Dim MaxNumber As Double
Dim MinNumber As Double
MaxNumber = Application.WorksheetFunction.Max(RNG1)
MinNumber = Application.WorksheetFunction.Min(RNG1)
With ActiveChart
    If .SeriesCollection(1).AxisGroup = 1 Then
        .Axes(xlValue).MinimumScale = MinNumber
        .Axes(xlValue).MaximumScale = MaxNumber
    End If
End With

Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub
Sub AdjustDateRanges()
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
Dim MaxDate As Date, MinDate As Date
Dim MyLastDate As Integer: MyLastDate = Cells(Rows.Count, "D").End(xlUp).Row
Dim RNG1 As Range: Set RNG1 = Range("D6:D" & MyLastDate)
MaxDate = Application.WorksheetFunction.Max(RNG1)
MinDate = Application.WorksheetFunction.Min(RNG1)
If Range("B3") = MaxDate And Range("B4") = MinDate Then
    ChangeChartAssets
    Exit Sub
Else
    GoTo UpdateNewDates
End If
UpdateNewDates:
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False
Dim LC As Long: LC = Range("D6").End(xlToRight).Column 'Cells(5, Columns.Count).End(xlToRight).Column
Dim LR As Long: LR = Range("D5").End(xlDown).Row 'Cells(Rows.Count, "D").End(xlUp).Row
Dim i As Long
Dim mydates As Range

For i = 6 To LR
    If Year(Cells(i, "D")) <= Year(Cells(3, 2)) And Year(Cells(i, "D")) >= Year(Cells(4, 2)) Then
        Cells(i, LC).Offset(0, 1).End(xlUp).Value = Cells(i, "D").Value
    End If
Next i
Dim New_LR As Long: New_LR = Cells(Rows.Count, LC).Offset(0, 1).End(xlUp).Row
Dim MyRNG As Range: Set MyRNG = Range(Cells(6, LC).Offset(0, 1), Cells(New_LR, LC).Offset(0, 1))

MyRNG.Copy
MyRNG.PasteSpecial xlPasteValues
'MyRNG.Font.ThemeColor = xlThemeColorDark1

Application.CutCopyMode = False
Columns(LC).Offset(0, 1).NumberFormat = "m/d/yyyy"
Update_Xvalues
Cells(1, 1).Select

End Sub
Sub Update_Xvalues()
If Cells(1, 2).Value = "" Or Cells(2, 2).Value = "" Then
    MsgBox "You must select two assets to measure in cells B1 and B2", vbOKOnly, "Missing Assets"
    Exit Sub
End If
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.ScreenUpdating = False

'Clear old Charts
Dim CHobj As ChartObject
For Each CHobj In ActiveSheet.ChartObjects
    CHobj.Delete
Next CHobj

'Set up Variables
Dim ser As Series
Dim ShName As String
Dim Str1 As String: Str1 = Range("B1").Value
Dim Str2 As String: Str2 = Range("B2").Value
Dim RNG1 As Range
Dim RNG2 As Range
Dim DateRNG As Range
Dim i As Integer
Dim LC As Integer: LC = Cells(6, Columns.Count).End(xlToLeft).Column
Dim LR As Long: LR = Cells(Rows.Count, LC).Offset(0, 1).End(xlUp).Row
Set DateRNG = Range(Cells(6, LC).Offset(0, 1), Cells(LR, LC).Offset(0, 1)) '& LR)
'Locate Series to add to chart
For i = 5 To LC
    If Cells(5, i).Value = Str1 Then
        Set RNG1 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
    Else
        If Cells(5, i) = Str2 Then
            Set RNG2 = Range(Cells(5, i).Offset(1, 0), Cells(LR, i))
        End If
    End If
Next i
'If values are same, do not update chart, exit sub
If Str1 = Str2 Then
    MsgBox "Values Must Be Different..." & vbNewLine _
    & vbNewLine & "Change Values in Cells B1 or B2", vbOKOnly, "Change an Asset..."
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If
Dim source As Range: Set source = Union(DateRNG, RNG1)
With ActiveSheet
    ShName = .Name
End With
Charts.Add
With ActiveChart
    .ChartType = xlLine
    .SetSourceData source:=Union(RNG1, RNG2)
    .SeriesCollection(1).Name = Str1
    .SeriesCollection(1).XValues = DateRNG
    With .SeriesCollection.NewSeries
            .Values = RNG2
            .XValues = DateRNG
            .Name = Str2
            .AxisGroup = 2
    End With
    .SetElement (msoElementLegendTop)
    .Location Where:=xlLocationAsObject, Name:=ShName
    With ActiveChart.ChartArea
        .Left = [A6].Left
        .Top = [A6].Top
        .Width = 800
        .Height = 600
    End With

For Each ser In ActiveChart.SeriesCollection
ser.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1
    End With
Next ser
    
End With

'The end result adjusts Y2 correctly, below the macro call, Seriescollection(1) will need to be adjusted as well
AdjustVerticalAxis

'Series(1) adjustments, Changes Collections to display max min scale correctly
Dim MaxNumber As Double
Dim MinNumber As Double
MaxNumber = Application.WorksheetFunction.Max(RNG1)
MinNumber = Application.WorksheetFunction.Min(RNG1)
With ActiveChart
    If .SeriesCollection(1).AxisGroup = 1 Then
        .Axes(xlValue).MinimumScale = MinNumber
        .Axes(xlValue).MaximumScale = MaxNumber
    End If
End With



Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub
Sub ResetToMaxDateRange()
Dim MyLastDate As Integer: MyLastDate = Cells(Rows.Count, "D").End(xlUp).Row
Dim RNG1 As Range: Set RNG1 = Range("D6:D" & MyLastDate)
Range("B3") = Application.WorksheetFunction.Max(RNG1)
Range("B4") = Application.WorksheetFunction.Min(RNG1)
Finance_MoveData.CreateTwoAxisChart
End Sub

