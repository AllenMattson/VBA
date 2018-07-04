Attribute VB_Name = "chartMacros"
Option Private Module
Option Explicit


Sub copyFormattingFromOtherSeries(sourceSeries As series, targetSeries As series)
    On Error Resume Next

    With targetSeries
        .Border.Color = RGB(1, 1, 1)

        .MarkerBackgroundColor = sourceSeries.MarkerBackgroundColor
        .MarkerForegroundColor = sourceSeries.MarkerForegroundColor

        .MarkerStyle = sourceSeries.MarkerStyle
        .MarkerSize = sourceSeries.MarkerSize
        .Format.Line.weight = sourceSeries.Format.Line.weight
        .Format.Line.Visible = sourceSeries.Format.Line.Visible

    End With
End Sub

Sub toggleSecondSeries()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Dim chartNum As Integer
    Dim newMetricNum As Integer
    Dim metricNum As Integer
    Dim chartObj As Object
    Dim firstMetricNum As Integer
    Dim newMetricNameDisp As String
    Dim firstMetricNameDisp As String


    sheetID = fetchValue("sheetID", ActiveSheet)
    chartNum = Right(Application.Caller, Len(Application.Caller) - InStrRev(Application.Caller, "_"))
    Set chartObj = ActiveSheet.ChartObjects(chartNum)
    metricNum = fetchValue("C" & chartNum & "2ndSeries", ActiveSheet)
    metricsCount = fetchValue("metricItemCount", ActiveSheet)
    ' firstMetricNum = Right(chartObj.Name, Len(chartObj.Name) - InStrRev(chartObj.Name, "_M") - 1)
    firstMetricNum = parseVarFromName(chartObj.Name, "M")
    comparisonType = fetchValue("comparisonType", ActiveSheet)

    If comparisonType <> "none" Then firstMetricNum = firstMetricNum * 2 - 1

    If metricNum + 1 = firstMetricNum Then
        newMetricNum = metricNum + 2
    Else
        newMetricNum = metricNum + 1
    End If



    With chartObj.Chart
        If newMetricNum > metricsCount Then
            newMetricNum = 0
            Call removeSecondarySeries(chartObj)
            Call storeValue("C" & chartNum & "2ndSeries", newMetricNum, ActiveSheet)
            ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).TextFrame.Characters.Text = "None"
            If fetchValue("showLegendInCharts", ActiveSheet) = True Then
                .HasLegend = False
                .HasLegend = True
                Call chartLegend(chartObj)
            End If
        Else
            If chartObj.Chart.SeriesCollection.Count >= 1 Then Call removeSecondarySeries(chartObj)
            Call setSecondDataSeries(chartObj, newMetricNum, metricNum)
            Call storeValue("C" & chartNum & "2ndSeries", newMetricNum, ActiveSheet)
            newMetricNameDisp = fetchValue("metricItemDisp" & newMetricNum, ActiveSheet)
            ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).TextFrame.Characters.Text = newMetricNameDisp
            If metricNum = 0 Then
                With .Axes(xlValue, xlSecondary)
                    .HasTitle = True
                    '  .Format.Line.DashStyle = 7   'msoLineLongDash
                    .Format.Line.Visible = False
                    .TickLabels.Font.Name = fontName
                    .AxisTitle.Font.Name = fontName
                End With
                If fetchValue("showLegendInCharts", ActiveSheet) = True Then
                    Call chartLegend(chartObj)
                Else
                    .PlotArea.Width = .PlotArea.Left + (.Parent.Width - .PlotArea.Left - 60)
                End If
            End If
            With .Axes(xlValue, xlSecondary)
                .HasTitle = True
                .AxisTitle.Text = newMetricNameDisp
            End With
            .Axes(xlValue, xlPrimary).AxisTitle.Font.ColorIndex = 1
        End If
    End With


    If debugMode = True Then Debug.Print "firstMetricNum: " & firstMetricNum & " newMetricNum: " & newMetricNum & "  " & newMetricNameDisp

End Sub


Sub setSecondDataSeries(chartObj As Object, newMetricNum As Integer, previous2ndMetricNum As Integer)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim seriesNum As Integer
    Dim seriesCount As Integer
    Dim firstMetricNum As Integer
    Dim comparisonType As String
    Dim groupByMetric As Boolean
    Dim showLegendInCharts As Boolean

    comparisonType = fetchValue("comparisonType", ActiveSheet)
    '  firstMetricNum = Right(chartObj.Name, Len(chartObj.Name) - InStrRev(chartObj.Name, "_M") - 1)
    firstMetricNum = parseVarFromName(chartObj.Name, "M")

    groupByMetric = CBool(fetchValue("groupByMetric", ActiveSheet))
    profileCount = fetchValue("profileCount", ActiveSheet)

    showLegendInCharts = fetchValue("showLegendInCharts", ActiveSheet)

    If comparisonType <> "none" Then firstMetricNum = firstMetricNum * 2 - 1

    If debugMode = True Then Debug.Print "    previous2ndMetricNum: " & previous2ndMetricNum & "->  newMetricNum: " & newMetricNum


    With chartObj.Chart
        seriesCount = .SeriesCollection.Count
        ' For seriesNum = 1 To .SeriesCollection.Count

        If previous2ndMetricNum = 0 Or seriesCount = 1 Then
            seriesNum = 1
            If seriesCount < 256 Then
                .SeriesCollection.NewSeries
                seriesCount = seriesCount + 1
                With .SeriesCollection(seriesCount)
                    If groupByMetric = False Then
                        '                            If comparisonType <> "none" Then
                        '                                .Values = GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, (newMetricNum - firstMetricNum) * 2)
                        '                            Else
                        .Values = GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, newMetricNum - firstMetricNum)
                        '   End If
                    Else
                        '                            If comparisonType <> "none" Then
                        '                                .Values = GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, (newMetricNum - firstMetricNum) * 2 * profileCount)
                        '                            Else
                        .Values = GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, (newMetricNum - firstMetricNum) * profileCount)
                        '       End If
                    End If
                    .Name = "%%" & seriesCount & "%" & seriesNum
                    .ChartType = xlLineMarkers
                    .Format.Fill.ForeColor.RGB = chartSeriesBlue
                    .Border.Color = chartSeriesBlue
                    .Interior.Color = chartSeriesBlue
                    .MarkerBackgroundColor = chartSeriesBlue
                    .MarkerForegroundColor = chartSeriesBlue
                    '.MarkerStyle = xlMarkerStyleNone
                    .MarkerStyle = 8
                    .MarkerSize = 3
                    .Format.Line.weight = 1

                    .Format.Fill.Transparency = 0.15


                    On Error Resume Next
                    .AxisGroup = xlSecondary
                    '   Call copyFormattingFromOtherSeries(chartObj.Chart.SeriesCollection(seriesNum), chartObj.Chart.SeriesCollection(seriesCount))
                    ' .Format.Line.DashStyle = msoLineLongDash
                    chartObj.Chart.Axes(xlValue, xlSecondary).Format.Line.Visible = False
                    ' .MarkerStyle = xlNone

                    If debugMode = True Then On Error GoTo 0
                End With
                If showLegendInCharts = True Then .Legend.LegendEntries(.Legend.LegendEntries.Count).Delete
            End If
        Else
            seriesNum = 2
            With .SeriesCollection(seriesNum)
                If Left(.Name, 2) = "%%" Then
                    If groupByMetric = False Then
                        Debug.Print "old rng: " & GetChartRange(chartObj.Chart, seriesNum, "values").Address
                        Debug.Print "new rng: " & GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, newMetricNum - previous2ndMetricNum).Address
                        .Values = GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, newMetricNum - previous2ndMetricNum)
                    Else
                        .Values = GetChartRange(chartObj.Chart, seriesNum, "values").Offset(, (newMetricNum - previous2ndMetricNum) * profileCount)
                    End If
                End If
            End With
        End If
        '  Next seriesNum
    End With

    Call storeValue("secondSeriesAddedToChart", True, ActiveSheet)

End Sub
Sub removeSecondarySeries(chartObj As Object)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim seriesNum As Integer
    With chartObj.Chart
        If .SeriesCollection.Count >= 2 Then .SeriesCollection(2).Delete
        '        For seriesNum = .SeriesCollection.Count To 1 Step -1
        '            If Left(.SeriesCollection(seriesNum).Name, 2) = "%%" Then .SeriesCollection(seriesNum).Delete
        '        Next seriesNum
        .Axes(xlValue, xlPrimary).AxisTitle.Font.ColorIndex = 2
    End With
End Sub
Sub resetSecondarySeries(chartObj As Object)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    '    Dim chartObj As ChartObject
    Dim metricNum As Integer
    Dim newMetricNameDisp As String

    '  For Each chartObj In ActiveSheet.ChartObjects
    If chartObj.Chart.ChartType <> xlXYScatter Then
        metricNum = fetchValue("C" & chartObj.Index & "2ndSeries", ActiveSheet)
        If metricNum > 0 Then
            newMetricNameDisp = fetchValue("metricDisp" & metricNum, ActiveSheet)
            Call removeSecondarySeries(chartObj)
            Call storeValue("C" & chartObj.Index & "2ndSeries", 0, ActiveSheet)
            Call setSecondDataSeries(chartObj, metricNum, 0)
            chartObj.Chart.SeriesCollection(1).ChartType = xlLineMarkers
            With chartObj.Chart.Axes(xlValue, xlSecondary)
                .HasTitle = True
                ' .Format.Line.DashStyle = 7   'msoLineLongDash
                .Format.Line.Visible = False
                .AxisTitle.Text = newMetricNameDisp
            End With
        End If
    End If
    ' Next
End Sub

Function GetChartRange(cht As Chart, series As Integer, ValOrX As String) As Range
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    '   cht: A Chart object
    '   series: Integer representing the Series
    '   ValOrX: String, either "values" or "xvalues"

    Dim sf As String
    Dim CommaCnt As Integer
    Dim Commas() As Integer
    Dim ListSep As String * 1
    Dim temp As String
    Dim i As Long
    Dim k As Integer

    Set GetChartRange = Nothing
    'On Error Resume Next

    '   Get the SERIES formula
    sf = cht.SeriesCollection(series).Formula

    '   Check for noncontiguous ranges by counting commas
    '   Also, store the character position of the commas
    For k = 1 To 3
        CommaCnt = 0
        Select Case k
        Case 1
            ListSep = Application.International(xlListSeparator)
        Case 2
            ListSep = ","
        Case 3
            ListSep = ";"
        End Select
        If IsArray(Commas) Then Erase Commas

        For i = 1 To Len(sf)
            If Mid(sf, i, 1) = ListSep Then
                CommaCnt = CommaCnt + 1
                ReDim Preserve Commas(CommaCnt)
                Commas(CommaCnt) = i
            End If
        Next i
        If CommaCnt > 0 Then Exit For
    Next k
    If CommaCnt > 3 Then Exit Function

    '   XValues or Values?
    Select Case UCase(ValOrX)
    Case "XVALUES"
        '           Text between 1st and 2nd commas in SERIES Formula
        temp = Mid(sf, Commas(1) + 1, Commas(2) - Commas(1) - 1)
        Set GetChartRange = Range(temp)
        'If debugMode = True Then Debug.Print Range(Temp).Address
    Case "VALUES"
        '           Text between the 2nd and 3rd commas in SERIES Formula
        temp = Mid(sf, Commas(2) + 1, Commas(3) - Commas(2) - 1)
        Set GetChartRange = Range(temp)
    End Select
End Function

Sub toggleChartTypeForOneChart()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim chartObj As ChartObject
    Dim chartNum As Integer
    sheetID = fetchValue("sheetID", ActiveSheet)
    chartNum = Right(Application.Caller, Len(Application.Caller) - InStrRev(Application.Caller, "_"))
    Set chartObj = ActiveSheet.ChartObjects(chartNum)
    Call changeChartType(False, chartObj)
End Sub

Sub changeChartType(Optional forAllCharts As Boolean = True, Optional chartObj As Object)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim newtype As XlChartType
    Dim oldType As XlChartType
    Dim chartNum As Integer
    Dim allow2ndSeries As Boolean

    Dim reformattingRequired As Boolean
    reformattingRequired = False

    Dim firstValidChartNum As Long

    sheetID = fetchValue("sheetID", ActiveSheet)

    If ActiveSheet.ChartObjects(1).Chart.ChartType = xlXYScatter Then
        firstValidChartNum = 2
    Else
        firstValidChartNum = 1
    End If

    If firstValidChartNum = 2 And ActiveSheet.ChartObjects.Count < 2 Then Exit Sub

    If forAllCharts = True Then
        oldType = Evaluate(fetchValue("chartType", ActiveSheet))
    Else
        oldType = chartObj.Chart.ChartType
        chartNum = chartObj.Index
        If oldType = -4111 Then oldType = chartObj.Chart.SeriesCollection(1).ChartType
    End If

    allow2ndSeries = False

    Select Case oldType
    Case xlLineMarkers
        newtype = xlColumnStacked
        reformattingRequired = True
    Case xlColumnStacked
        newtype = xlColumnClustered
    Case xlColumnClustered
        newtype = xlColumnStacked100
    Case xlColumnStacked100
        newtype = xlAreaStacked
    Case xlAreaStacked
        newtype = xlAreaStacked100
    Case xlAreaStacked100
        newtype = xl3DColumnStacked
    Case xl3DColumnStacked
        newtype = xl3DColumnClustered
    Case xl3DColumnClustered
        newtype = xl3DColumnStacked100
    Case xl3DColumnStacked100
        newtype = xl3DColumn
    Case xl3DColumn
        newtype = xl3DLine
    Case xl3DLine
        newtype = xlBarStacked
    Case xlBarStacked
        newtype = xlBarClustered
    Case xlBarClustered
        newtype = xlBarStacked100
    Case xlBarStacked100
        newtype = xlRadarMarkers
        reformattingRequired = True
    Case xlRadarMarkers
        newtype = xlLineMarkers
    Case Else
        newtype = xlLineMarkers
    End Select


    Select Case newtype
    Case xlLineMarkers, xlColumnStacked, xlColumnClustered, xlAreaStacked
        allow2ndSeries = True
    End Select

    Call storeValue("chartType", newtype, ActiveSheet)

    If fetchValue("secondSeriesAddedToChart", ActiveSheet) = True Then
        If forAllCharts = True Then
            For chartNum = 1 To ActiveSheet.ChartObjects.Count
                If ActiveSheet.ChartObjects(chartNum).Chart.ChartType <> xlXYScatter Then
                    Call removeSecondarySeries(ActiveSheet.ChartObjects(chartNum))
                    If allow2ndSeries = True Then
                        ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).Visible = True
                        ActiveSheet.Shapes(sheetID & "CASSB2_" & chartNum).Visible = True
                    Else
                        ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).Visible = False
                        ActiveSheet.Shapes(sheetID & "CASSB2_" & chartNum).Visible = False
                    End If
                    Call storeValue("C" & chartNum & "2ndSeries", 0, ActiveSheet)
                    ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).TextFrame.Characters.Text = "None"
                End If
            Next chartNum
        Else
            Call removeSecondarySeries(chartObj)
            If allow2ndSeries = True Then
                ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).Visible = True
                ActiveSheet.Shapes(sheetID & "CASSB2_" & chartNum).Visible = True
            Else
                ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).Visible = False
                ActiveSheet.Shapes(sheetID & "CASSB2_" & chartNum).Visible = False
            End If
            Call storeValue("C" & chartNum & "2ndSeries", 0, ActiveSheet)
            ActiveSheet.Shapes(sheetID & "CASSB1_" & chartNum).TextFrame.Characters.Text = "None"
        End If
    End If

    If forAllCharts = True Then
        For chartNum = 1 To ActiveSheet.ChartObjects.Count
            If ActiveSheet.ChartObjects(chartNum).Chart.ChartType <> xlXYScatter Then
                ActiveSheet.ChartObjects(chartNum).Chart.ChartType = newtype
                If reformattingRequired = True Then
                    Call formatChart(ActiveSheet.ChartObjects(chartNum))
                End If
            End If
        Next chartNum
        ActiveSheet.ChartObjects(1).Select
        Range("C1").Select
    Else
        chartObj.Chart.ChartType = newtype
        If reformattingRequired = True Then Call formatChart(chartObj)
    End If



End Sub



Sub setChartCategories(Optional forceNumberOfCategories As Long = 0)

    On Error Resume Next

    Dim sheetID As String
    sheetID = Cells(1, 1).value
    sheetID = findRangeName(Cells(1, 1))

    Dim firstChartToModify As Boolean
    firstChartToModify = True

    Application.ScreenUpdating = False
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

    Dim numberOfCategories As Variant
    If forceNumberOfCategories = 0 Then
        numberOfCategories = Application.Index(Range("chartCategoriesNumberList"), fetchValue("catSel", ActiveSheet))
    Else
        numberOfCategories = forceNumberOfCategories
    End If

    Dim firstRow As Long
    Dim lastRow As Long
    Dim newLastRow As Long

    firstRow = fetchValue("firstDataRow", ActiveSheet)
    lastRow = fetchValue("lastDataRow", ActiveSheet)

    Dim formulaStr As String
    Dim formulaStrAlku As String
    Dim formulaStrLoppu As String

    Dim num As Double
    Dim numalku As Long
    Dim numLoppu As Long
    Dim kirjain As Long

    Dim chartTypesArr As Variant


    Dim originalType As XlChartType

    If numberOfCategories > lastRow - firstRow + 1 Then numberOfCategories = lastRow - firstRow + 1

    If numberOfCategories = "All" Then numberOfCategories = lastRow - firstRow + 1

    newLastRow = firstRow + numberOfCategories - 1

    Dim chartNum As Long
    Dim seriesNum As Long

    For chartNum = 1 To ActiveSheet.ChartObjects.Count

        With ActiveSheet.ChartObjects(chartNum).Chart

            originalType = .ChartType
            ReDim chartTypesArr(1 To .SeriesCollection.Count)
            For seriesNum = 1 To .SeriesCollection.Count
                chartTypesArr(seriesNum) = .SeriesCollection(seriesNum).ChartType
            Next seriesNum
            .ChartType = xlArea


            For seriesNum = 1 To .SeriesCollection.Count

                formulaStr = .SeriesCollection(seriesNum).Formula

                ' If firstChartToModify = True Or originalType = -4169 Then  'first chart or scatterplot





                For kirjain = Len(formulaStr) To 1 Step -1
                    If Mid$(formulaStr, kirjain, 1) = "," Then numLoppu = kirjain - 1
                    If Mid$(formulaStr, kirjain, 1) = "$" Then
                        numalku = kirjain + 1
                        num = Mid$(formulaStr, numalku, numLoppu - numalku + 1)
                        Exit For
                    End If

                Next kirjain
                firstChartToModify = False
                '  End If

                formulaStrAlku = Left$(formulaStr, InStr(1, formulaStr, Chr$(34) & ","))
                formulaStrLoppu = Right$(formulaStr, Len(formulaStr) - InStr(1, formulaStr, Chr$(34) & ","))
                Dim colLetter As String
                Dim dataSeriesRng As Range
                If firstRow <> newLastRow Then
                    Set dataSeriesRng = SeriesRange(.SeriesCollection(seriesNum), "primary")
                    colLetter = ColumnLetter(dataSeriesRng.Cells(1, 1).Column)
                    formulaStrLoppu = Replace(formulaStrLoppu, "!$" & colLetter & "$" & num & ",", "!$" & colLetter & "$" & num & ":$" & colLetter & "$" & num & ",")
                    Set dataSeriesRng = SeriesRange(.SeriesCollection(seriesNum), "labels")
                    colLetter = ColumnLetter(dataSeriesRng.Cells(1, 1).Column)
                    formulaStrLoppu = Replace(formulaStrLoppu, "!$" & colLetter & "$" & num & ",", "!$" & colLetter & "$" & num & ":$" & colLetter & "$" & num & ",")
                    Set dataSeriesRng = SeriesRange(.SeriesCollection(seriesNum), "secondary")
                    colLetter = ColumnLetter(dataSeriesRng.Cells(1, 1).Column)
                    formulaStrLoppu = Replace(formulaStrLoppu, "!$" & colLetter & "$" & num & ",", "!$" & colLetter & "$" & num & ":$" & colLetter & "$" & num & ",")
                End If

                If num <> newLastRow Then
                    formulaStrLoppu = Replace(formulaStrLoppu, "$" & num & ",", "$" & newLastRow & ",")
                    .SeriesCollection(seriesNum).Formula = formulaStrAlku & formulaStrLoppu
                End If



            Next seriesNum


            .ChartType = originalType
            For seriesNum = 1 To .SeriesCollection.Count
                .SeriesCollection(seriesNum).ChartType = chartTypesArr(seriesNum)
            Next seriesNum


            If numberOfCategories > 50 Then
                .Axes(xlCategory).TickLabels.Font.Size = 8
            Else
                .Axes(xlCategory).TickLabels.Font.Size = 9
            End If


        End With



    Next chartNum

End Sub


Function SeriesRange(s As series, Optional seriesType As String = "primary") As Range
    Dim sf As String, fa() As String


    sf = s.Formula
    sf = Replace(sf, "=SERIES(", "")

    If sf = "" Then
        Set SeriesRange = Nothing
        Exit Function
    End If

    fa = Split(sf, ",")

    If seriesType = "labels" Then
        Set SeriesRange = Range(fa(1))
    ElseIf seriesType = "primary" Then
        Set SeriesRange = Range(fa(2))
    ElseIf seriesType = "secondary" Then
        Set SeriesRange = Range(fa(UBound(fa)))
    End If
End Function




Sub chartLegend(chartObj As ChartObject)

    On Error Resume Next

    If fetchValue("showLegendInCharts", ActiveSheet) = False Then Exit Sub

    Application.StatusBar = "Setting chart legend..."


    Dim oLegend As Legend
    Dim oPlotArea As PlotArea

    Dim legerrors As Long
    legerrors = 0

    Dim cWidth As Long
    Dim cHeight As Long

    Dim i As Long

    Dim maxLeft As Long
    Dim lenum As Long


    Set oLegend = chartObj.Chart.Legend
    Set oPlotArea = chartObj.Chart.PlotArea

    Dim w As Long
    Dim l As Long

    With chartObj

        cWidth = .Width
        cHeight = .Height

    End With



    With chartObj.Chart

        .HasLegend = True


        oLegend.Top = 0
        If .SeriesCollection.Count * 20 > cHeight Then
            oLegend.Height = cHeight
        Else
            oLegend.Height = .SeriesCollection.Count * 20
        End If



        Application.StatusBar = "Formatting chart... " & .ChartTitle.Text & " Legend size"

        For lenum = 1 To oLegend.LegendEntries.Count

            On Error GoTo legerrorh
            If oLegend.LegendEntries(lenum).Top > 10 Then
            End If

        Next lenum

        On Error Resume Next
        maxLeft = oLegend.LegendEntries(1).Left

        If legerrors = 0 Then
            For lenum = 1 To oLegend.LegendEntries.Count

                If oLegend.LegendEntries(lenum).Left > maxLeft + 5 Then

                    For i = 1 To 3

                        oLegend.Width = .Legend.Width - 25
                        If oLegend.LegendEntries(lenum).Left <= maxLeft + 1 Then Exit For
                    Next i

                End If

            Next lenum
        End If




        On Error Resume Next


        oLegend.Left = cWidth - oLegend.Width
        oPlotArea.Width = cWidth - oLegend.Width - 5 - .PlotArea.Left
        If .Axes.Count = 3 Then  'secondary axis
            oPlotArea.Width = cWidth - oLegend.Width - 5 - .PlotArea.Left - 15
        End If

    End With



    Application.StatusBar = False


    Exit Sub


legerrorh:
    legerrors = legerrors + 1

    If oLegend.Font.Size > 6 Then

        oLegend.Font.Size = oLegend.Font.Size - 1

    ElseIf oPlotArea.Width > 300 Then

        oPlotArea.Width = oPlotArea.Width - 25

        w = oPlotArea.Width
        l = oPlotArea.Left

        oLegend.Left = l + w + 5
        oLegend.Width = cWidth - (l + w + 5)

    End If



    If legerrors > 10 Then
        Resume Next
    Else
        Resume
    End If

End Sub






Sub formatChart(chartObj As ChartObject)

    On Error Resume Next
    'On Error GoTo 0
    Application.ScreenUpdating = False

    Dim titleStr As String
    Dim seriesNum As Integer
    Dim seriesColour As Long

    titleStr = chartObj.Chart.ChartTitle.Text

    For seriesNum = 1 To chartObj.Chart.SeriesCollection.Count

        With chartObj.Chart.SeriesCollection(seriesNum)

            Application.StatusBar = "Formatting chart... " & titleStr & " " & .Name

            '.MarkerStyle = xlMarkerStyleNone
            .MarkerStyle = 8
            .MarkerSize = 3
            .Format.Line.weight = 1
            .Format.Line.Visible = False


            If Left(.Name, 2) = "%%" Then
                .Delete
            Else
                seriesColour = getSeriesColour(seriesNum)
                .Format.Fill.ForeColor.RGB = seriesColour
                .Border.Color = seriesColour
                .Interior.Color = seriesColour
                .MarkerBackgroundColor = seriesColour
                .MarkerForegroundColor = seriesColour
            End If
            .Format.Fill.Transparency = Range("chartSeriesTransparency").value / 100
        End With
    Next seriesNum

    Application.StatusBar = False

End Sub

Public Function getSeriesColour(seriesNum As Integer) As Long

    If seriesNum > 12 Then
        getSeriesColour = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
    ElseIf excelVersion <= 11 Then
        Select Case seriesNum
        Case 1
            getSeriesColour = RGB(230, 0, 0)
        Case 2
            getSeriesColour = RGB(0, 112, 192)
        Case 3
            getSeriesColour = RGB(122, 188, 50)
        Case 4
            getSeriesColour = RGB(255, 192, 0)
        Case 5
            getSeriesColour = RGB(112, 48, 160)
        Case 6
            getSeriesColour = RGB(13, 13, 13)
        Case 7
            getSeriesColour = RGB(146, 208, 80)
        Case 8
            getSeriesColour = RGB(151, 72, 7)
        Case 9
            getSeriesColour = RGB(85, 142, 213)
        Case 10
            getSeriesColour = RGB(245, 117, 11)
        Case 11
            getSeriesColour = RGB(255, 51, 51)
        Case 12
            getSeriesColour = RGB(127, 127, 127)
        End Select
    Else
        getSeriesColour = Range("seriesColoursStart").Offset(seriesNum - 1).Interior.Color
    End If
End Function


Sub addScatterPlotLabels(chartObj As ChartObject, labelColumnNum As Long)

    On Error Resume Next


    Dim Counter As Long
    Dim formulaStr As String
    Dim numLoppu As Long
    Dim numalku As Long
    Dim kirjain As Long
    Dim xValRange As String

    With chartObj.Chart

        '.ChartType = xlArea

        'Store the formula for the first series in "xVals".
        formulaStr = .SeriesCollection(1).Formula



        numLoppu = 9999

        For kirjain = Len(formulaStr) To 1 Step -1
            If numLoppu <> 9999 And Mid$(formulaStr, kirjain, 1) = "," Then
                numalku = kirjain + 1
                xValRange = Mid$(formulaStr, numalku, numLoppu - numalku + 1)
                Exit For
            End If
            If Mid$(formulaStr, kirjain, 1) = "," Then numLoppu = kirjain - 1


        Next kirjain

        'Attach a label to each data point in the chart.
        For Counter = 1 To Range(xValRange).Cells.Count
            .SeriesCollection(1).Points(Counter).HasDataLabel = True
            .SeriesCollection(1).Points(Counter).DataLabel.Text = ActiveSheet.Cells(Range(xValRange).Cells(Counter, 1).row, labelColumnNum).value
            .SeriesCollection(1).Points(Counter).DataLabel.Text = "=" & ActiveSheet.Name & "!" & Range(ColumnLetter(labelColumnNum) & Range(xValRange).Cells(Counter, 1).row).Address
        Next Counter


        .ChartType = xlXYScatter

        .SeriesCollection(1).DataLabels.Font.Size = 8

    End With

End Sub


Sub setScatterPlotSeriesSD()
    Call setScatterPlotSeries
    'Call setChartCategories
End Sub


Sub setScatterPlotSeries(Optional labelColumnNum As Long, Optional secondLabelColumn As Long = 0)

'On Error Resume Next

    Application.ScreenUpdating = False

    Call checkOperatingSystem

    Dim sheetID As String
    ActiveSheet.Cells(1, 1).Select
    sheetID = Cells(1, 1).value
    sheetID = findRangeName(Cells(1, 1))

    Dim chartObj As ChartObject
    If ChartExists(sheetID & "scatterPlot", ActiveSheet) = False Then Exit Sub
    Set chartObj = ActiveSheet.ChartObjects(sheetID & "scatterPlot")

    If chartObj Is Nothing Then Exit Sub


    Dim seriesNum As Long

    Dim xSelection As Long
    Dim ySelection As Long

    Dim xLabel As String
    Dim yLabel As String


    Dim yCol As Long
    Dim xCol As Long
    Dim firstRow As Long
    Dim lastRow As Long


    Dim rivi As Long
    Dim yarr() As Variant
    Dim xarr() As Variant
    Dim labelarr() As Variant

    Dim xyCount As Long
    Dim xyNum As Long

    xyCount = 0

    If IsMissing(labelColumnNum) Or labelColumnNum = 0 Then
        labelColumnNum = fetchValue("rowLabelsCol", ActiveSheet)
        If IsNumeric(fetchValue("rowLabelsCol2", ActiveSheet)) Then
            secondLabelColumn = fetchValue("rowLabelsCol2", ActiveSheet)
        Else
            secondLabelColumn = 0
        End If
    End If

    firstRow = fetchValue("firstDataRow", ActiveSheet)
    lastRow = fetchValue("lastDataRow", ActiveSheet)


    With ActiveSheet

        ySelection = fetchValue("YvaluesSel", ActiveSheet)
        xSelection = fetchValue("XvaluesSel", ActiveSheet)

        xLabel = Range(sheetID & "_x1col").Offset(xSelection - 1, -1)
        yLabel = Range(sheetID & "_y1col").Offset(ySelection - 1, -1)

        yCol = .Cells(Range(sheetID & "_y1col").row + ySelection - 1, Range(sheetID & "_y1col").Column).value
        xCol = .Cells(Range(sheetID & "_x1col").row + xSelection - 1, Range(sheetID & "_x1col").Column).value


        For rivi = firstRow To lastRow
            If .Cells(rivi, yCol).value <> vbNullString And .Cells(rivi, yCol).value <> "NEW" And .Cells(rivi, xCol).value <> vbNullString And .Cells(rivi, xCol).value <> "NEW" Then
                xyCount = xyCount + 1
            End If
        Next rivi

        If xyCount = 0 Then xyCount = 1
        If xyCount > 256 Then xyCount = 256

        ReDim yarr(1 To xyCount)
        ReDim xarr(1 To xyCount)
        ReDim labelarr(1 To xyCount)
        xyNum = 0
        For rivi = firstRow To lastRow

            If Cells(rivi, yCol).value <> vbNullString And Cells(rivi, yCol).value <> "NEW" And Cells(rivi, xCol).value <> vbNullString And Cells(rivi, xCol).value <> "NEW" Then
                xyNum = xyNum + 1
                yarr(xyNum) = Round(Cells(rivi, yCol).value, 4)
                xarr(xyNum) = Round(Cells(rivi, xCol).value, 4)
                If secondLabelColumn > 0 Then
                    labelarr(xyNum) = Cells(rivi, labelColumnNum).value & " | " & Cells(rivi, secondLabelColumn).value
                Else
                    labelarr(xyNum) = Cells(rivi, labelColumnNum).value
                End If
                If xyNum >= 256 Then Exit For
            End If

        Next rivi

    End With

    With chartObj.Chart

        If excelVersion <= 11 Then

            .ChartType = xlArea

            yarr = Application.Transpose(yarr)
            xarr = Application.Transpose(xarr)

            With Cells(30, 9).Resize(xyCount)
                .value = yarr
                .Font.ColorIndex = 2
            End With
            With Cells(30, 10).Resize(xyCount)
                .value = xarr
                .Font.ColorIndex = 2
            End With

            .SeriesCollection(1).Values = Cells(30, 9).Resize(xyCount)
            .SeriesCollection(1).XValues = Cells(30, 10).Resize(xyCount)


            .ChartType = xlXYScatter

        Else

            .SeriesCollection(1).Values = yarr
            .SeriesCollection(1).XValues = xarr

        End If

        .Axes(xlValue).AxisTitle.Text = yLabel
        .Axes(xlCategory).AxisTitle.Text = xLabel

        .Axes(xlValue).TickLabels.NumberFormat = Cells(firstRow, yCol).NumberFormat
        .Axes(xlCategory).TickLabels.NumberFormat = Cells(firstRow, xCol).NumberFormat


        With .SeriesCollection(1)
            For xyNum = 1 To xyCount
                .Points(xyNum).HasDataLabel = True
                .Points(xyNum).DataLabel.Text = labelarr(xyNum)
            Next xyNum
        End With


        .PlotArea.Left = 15
        .HasTitle = True
        .ChartTitle.Text = UCase(yLabel & " vs. " & xLabel)
        .ChartTitle.Left = .PlotArea.InsideLeft

        With .SeriesCollection(1).DataLabels.Font
            .Size = 8
            .Color = RGB(89, 89, 89)
        End With

    End With


End Sub





