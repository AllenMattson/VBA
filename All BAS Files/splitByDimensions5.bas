Attribute VB_Name = "splitByDimensions5"
Option Private Module
Option Explicit

Sub fetchFigureSplitByDimensionsFormattingCharts()


    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim settingsSheetVisibility As Integer

    Dim kuvaaja As Object
    Dim lineChartsCreated As Boolean
    Dim rivi As Long
    Dim rivi2 As Long
    Dim col As Long
    Dim col2 As Long

    Dim seriesAdded As Boolean

    Dim columnCount As Long

    Dim profName As String
    Dim firstRowValueChanged As Boolean
    Dim seriesNum As Long
    Dim arvo As Variant
    Dim buttonObjPrev As Object
    Dim YvaluesDropdown As Object
    Dim XvaluesDropdown As Object

    Dim chartCategoriesDropdown As Object

    Dim storedChartType As Integer

    Dim chartCount As Integer
    Dim scatterPlotCreated As Boolean
    Dim chartsCreated As Boolean
    chartsCreated = False
    scatterPlotCreated = False

    Application.ScreenUpdating = False


    With dataSheet

        stParam1 = "8.31"

        If sendMode = True Then Call checkE(email, dataSource)

        If createCharts = True And updatingPreviouslyCreatedSheet = True And timeDimensionIncluded = False Then
            Call setScatterPlotSeries(dimensionsCombinedCol)
        End If


        lineChartsCreated = False

        Dim legendSizesArr() As Variant

        Dim isFirstChart As Boolean

        Dim chartNum As Long
        Dim chartsPerMetric As Long

        If createCharts = True And updatingPreviouslyCreatedSheet = False Then

            rivi = lastHeaderRow + 4

            isFirstChart = True


            If queryType = "SD" Then
                chartsPerMetric = profileCount * segmentCount
            Else
                If segmentCount > 1 Then
                    chartsPerMetric = profileCount
                Else
                    chartsPerMetric = 1
                End If
            End If
            chartCount = metricsCount * chartsPerMetric


            ReDim legendSizesArr(1 To chartCount, 1 To 2)



            For metricNum = 1 To metricsCount

                segmentNum = 0
                profNum = 0

                For chartNum = 1 To chartsPerMetric

                    progresspct = progresspct + (100 - progresspct) * 0.05


                    If queryType = "SD" Then



                        If segmentCount > 1 Then
                            If profNum = 0 Then profNum = 1
                            segmentNum = segmentNum + 1
                            If segmentNum > segmentCount Then
                                segmentNum = 1
                                profNum = profNum + 1
                            End If
                            segmentName = segmentArr(segmentNum, 2)
                        Else
                            segmentNum = 1
                            profNum = profNum + 1
                        End If

                        profID = profilesArr(profNum, 3)
                        profName = profilesArr(profNum, 2)

                        progresspct = progresspct + (100 - progresspct) * 0.05


                        If segmentCount > 1 Then
                            Call updateProgress(progresspct, "Creating charts... " & profName & " | " & metricsArr(metricNum, 1) & " | " & segmentName)
                        Else
                            Call updateProgress(progresspct, "Creating charts... " & profName & " | " & metricsArr(metricNum, 1))
                        End If

                    Else

                        If segmentCount > 1 Then
                            profNum = profNum + 1
                            profID = profilesArr(profNum, 3)
                            profName = profilesArr(profNum, 2)
                            Call updateProgress(progresspct, "Creating charts... " & profName & " | " & metricsArr(metricNum, 1))
                        Else
                            Call updateProgress(progresspct, "Creating charts... " & metricsArr(metricNum, 1))
                        End If

                    End If

                    Application.ScreenUpdating = False

                    If sumAllProfiles Then profName = vbNullString

                    seriesNum = 0


                    With Sheets("settings")

                        settingsSheetVisibility = .Visible
                        .Visible = xlSheetVisible
                        .Select
                        .Cells(1, 1).Select

                        .ChartObjects("chartDim").Duplicate
                        .Cells(1, 1).Select
                        Set kuvaaja = .ChartObjects(.ChartObjects.Count)
                        kuvaaja.Chart.Location where:=xlLocationAsObject, Name:=dataSheet.Name

                    End With


                    .Select
                    Set kuvaaja = .ChartObjects(.ChartObjects.Count)


                    Sheets("settings").Visible = settingsSheetVisibility

                    If queryType = "SD" Then
                        kuvaaja.Name = sheetID & "_C" & .ChartObjects.Count & "_N" & chartNum & "_" & metricsArr(metricNum, 2) & "_M" & metricNum & "_P" & profID & "_SG" & segmentNum & "_X" & profNum
                    Else
                        kuvaaja.Name = sheetID & "_C" & .ChartObjects.Count & "_N" & chartNum & "_" & metricsArr(metricNum, 2) & "_M" & metricNum & "_SG" & segmentNum & "_X1"
                    End If

                    kuvaaja.Top = .Cells(rivi, reportStartColumn + 1).Top
                    kuvaaja.Left = .Cells(rivi, reportStartColumn + 1).Left

                    kuvaaja.Placement = xlFreeFloating

                    With kuvaaja.Chart

                        .ChartType = xlArea

                        If queryType = "SD" Then
                            If profileCount > 1 Then
                                If segmentCount > 1 Then
                                    .ChartTitle.Text = profName & " | " & segmentName & " | " & metricsArr(metricNum, 1)
                                Else
                                    .ChartTitle.Text = profName & " | " & metricsArr(metricNum, 1)
                                End If
                            Else
                                If segmentCount > 1 Then
                                    .ChartTitle.Text = segmentName & " | " & metricsArr(metricNum, 1)
                                Else
                                    .ChartTitle.Text = metricsArr(metricNum, 1)
                                End If
                            End If
                        Else
                            If segmentCount > 1 And profileCount > 1 Then
                                .ChartTitle.Text = profName & " | " & metricsArr(metricNum, 1)
                            Else
                                .ChartTitle.Text = metricsArr(metricNum, 1)
                            End If
                        End If


                        With .ChartTitle
                            If Left(.Text, 3) = " | " Then .Text = Right(.Text, Len(.Text) - 3)
                            .Text = UCase(.Text)
                            .Font.Name = fontName
                        End With
                        .ChartTitle.Left = .PlotArea.InsideLeft



                        For col = firstMetricCol To vsarData

                            firstRowValueChanged = False

                            If columnInfoArr(col, 8) = metricsArr(metricNum, 2) Then
                                '       If dataSheet.Cells(metricNameRow, col).MergeArea.Cells(1, 1).value = metricsArr(metricNum, 1) Then
                                '  If queryType <> "SD" Or columnInfoArr(col, 12) = profID Then


                                If (queryType = "SD" And columnInfoArr(col, 12) = profID And columnInfoArr(col, 14) = segmentNum) Or (queryType <> "SD" And segmentCount > 1 And columnInfoArr(col, 12) = profID) Or (queryType <> "SD" And segmentCount = 1) Then

                                    If columnInfoArr(col, 6) = False Then  'not change columns

                                        If columnInfoArr(col, 10) = vbNullString Then    'check for data errors

                                            If dataSheet.Cells(1, col).EntireColumn.Hidden = False Then

                                                lineChartsCreated = True
                                                Call updateProgressIterationBoxes

                                                chartsCreated = True
                                                seriesNum = seriesNum + 1

                                                If seriesNum < 256 Then
                                                    If seriesNum > 1 Then .SeriesCollection.NewSeries

                                                    If dataSheet.Cells(resultStartRow, col).value = vbNullString Then
                                                        firstRowValueChanged = True
                                                        dataSheet.Cells(resultStartRow, col).value = 1
                                                    End If

                                                    .SeriesCollection(seriesNum).Values = dataSheet.Range(dataSheet.Cells(lastHeaderRow + 1, col), dataSheet.Cells(vriviChart, col))
                                                    .SeriesCollection(seriesNum).XValues = dataSheet.Range(dataSheet.Cells(lastHeaderRow + 1, dimensionsCombinedCol), dataSheet.Cells(vriviChart, dimensionsCombinedCol))

                                                    If debugMode = False Then On Error Resume Next

                                                    If queryType = "SD" Then
                                                        If excelVersion > 11 Then
                                                            .SeriesCollection(seriesNum).Name = "=" & Chr(39) & sheetName & Chr(39) & "!" & Range(ColumnLetter(col) & segmDimRow).MergeArea.Cells(1, 1).Address
                                                        Else
                                                            .SeriesCollection(seriesNum).Name = dataSheet.Cells(segmDimRow, col).MergeArea.Cells(1, 1)
                                                        End If
                                                    ElseIf segmentCount > 1 Then
                                                        If Left(.SeriesCollection(seriesNum).Name, 2) <> "%%" Then
                                                            If excelVersion > 11 Then
                                                                .SeriesCollection(seriesNum).Name = "=" & Chr(39) & sheetName & Chr(39) & "!" & Range(ColumnLetter(col) & segmentRow).MergeArea.Cells(1, 1).Address
                                                            Else
                                                                .SeriesCollection(seriesNum).Name = dataSheet.Cells(segmentRow, col).MergeArea.Cells(1, 1)
                                                            End If
                                                        End If
                                                    Else
                                                        If Left(.SeriesCollection(seriesNum).Name, 2) <> "%%" Then
                                                            If excelVersion > 11 Then
                                                                .SeriesCollection(seriesNum).Name = "=" & Chr(39) & sheetName & Chr(39) & "!" & Range(ColumnLetter(col) & profNameRow).MergeArea.Cells(1, 1).Address
                                                            Else
                                                                .SeriesCollection(seriesNum).Name = dataSheet.Cells(profNameRow, col).MergeArea.Cells(1, 1)
                                                            End If
                                                        End If
                                                    End If
                                                    Call updateProgressAdditionalMessage(.SeriesCollection(seriesNum).Name)
                                                    If firstRowValueChanged = True Then dataSheet.Cells(resultStartRow, col).value = vbNullString

                                                    Call storeValue(kuvaaja.Name & "_dataColumns", fetchValue(kuvaaja.Name & "_dataColumns", ActiveSheet) & "|" & col & "|", ActiveSheet)

                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                        Next col

                        If seriesNum > 256 And metricNum = 1 And chartNum = 1 Then
                            MsgBox "The query returned " & seriesNum - 1 & " data series. Due to Excel's limitations, the charts can only show the first 256 of these."
                        End If

                        .ChartType = xlLineMarkers

                        Call formatChart(kuvaaja)

                        With .Axes(xlCategory)
                            .TickLabels.Orientation = xlTickLabelOrientationAutomatic
                            With .AxisTitle
                                .Text = dimensionHeadersCombined
                                .Top = kuvaaja.Chart.PlotArea.Top + kuvaaja.Chart.PlotArea.Height + 5
                                .Left = kuvaaja.Chart.PlotArea.InsideLeft + kuvaaja.Chart.PlotArea.InsideWidth / 2
                            End With
                        End With

                        With .Axes(xlValue).AxisTitle
                            If InStr(1, metricsArr(metricNum, 1), "(") > 0 Then
                                arvo = Left(metricsArr(metricNum, 1), InStr(1, metricsArr(metricNum, 1), "(") - 2)
                            Else
                                arvo = metricsArr(metricNum, 1)
                            End If
                            .Text = arvo
                            .Font.ColorIndex = 2
                            .Top = kuvaaja.Chart.PlotArea.Top + kuvaaja.Chart.PlotArea.InsideHeight / 2
                        End With

                        .ChartTitle.Left = .PlotArea.InsideLeft


                        If seriesNum = 1 And queryType <> "SD" Then
                            .HasLegend = False
                            Call storeValue("showLegendInCharts", False, dataSheet)
                            .PlotArea.Width = .Parent.Width - .PlotArea.Left - 10
                        Else
                            Call storeValue("showLegendInCharts", True, dataSheet)
                            If Range("doChartLegendSizeOptimization").value = True Then
                                If metricNum = 1 Or dataSource = "FB" Then
                                    Call chartLegend(kuvaaja)
                                    legendSizesArr(chartNum, 1) = .Legend.Font.Size
                                    legendSizesArr(chartNum, 2) = .Legend.Left
                                Else
                                    With .Legend
                                        .Top = 0
                                        If kuvaaja.Chart.SeriesCollection.Count * 20 > kuvaaja.Height Then
                                            .Height = kuvaaja.Height
                                        Else
                                            .Height = kuvaaja.Chart.SeriesCollection.Count * 20
                                        End If
                                        .Font.Size = legendSizesArr(chartNum, 1)
                                        .Left = legendSizesArr(chartNum, 2)
                                    End With
                                    .PlotArea.Width = legendSizesArr(chartNum, 2) - 5 - .PlotArea.Left
                                End If
                            End If
                        End If

                        If vriviData - (resultStartRow) + 1 > 50 Then
                            .Axes(xlCategory).TickLabels.Font.Size = 8
                        Else
                            .Axes(xlCategory).TickLabels.Font.Size = 9
                        End If

                        Call storeValue("chartType", xlLineMarkers, ActiveSheet)
                        Call storeValue("C" & dataSheet.ChartObjects.Count & "2ndSeries", 0, ActiveSheet)

                        Set buttonObjPrev = Nothing

                        If metricsCount > 1 And profileCount = 1 And queryType = "D" And segmentCount <= 1 Then
                            Set buttonObj = dataSheet.Shapes.AddShape(5, 342, 15, 118, 29)
                            With buttonObj
                                .Name = sheetID & "CASSB1_" & dataSheet.ChartObjects.Count
                                .OnAction = "toggleSecondSeries"
                                .Adjustments(1) = 0.1
                                With .TextFrame
                                    .HorizontalAlignment = xlHAlignLeft
                                    .VerticalAlignment = xlVAlignBottom
                                    .MarginBottom = 0
                                    .Characters.Text = "None"
                                    .Characters.Font.Color = chartSeriesBlue
                                    .Characters.Font.Size = 7
                                    .Characters.Font.Name = fontName
                                End With
                                .Fill.ForeColor.RGB = buttonColour
                                .Line.ForeColor.RGB = buttonBorderColour
                                .Height = buttonHeight - 4
                                .Width = buttonWidth * 2 + buttonSpaceBetween
                                .Top = kuvaaja.Top - .Height - 5
                                .Left = kuvaaja.Left + kuvaaja.Width - .Width
                                .Placement = xlFreeFloating
                            End With

                            Set buttonObjPrev = buttonObj
                            Set buttonObj = dataSheet.Shapes.AddShape(152, 342, 15, 118, 29)
                            With buttonObj
                                .Name = sheetID & "CASSB2_" & dataSheet.ChartObjects.Count
                                .OnAction = "toggleSecondSeries"
                                .Adjustments(1) = 0.1
                                With .TextFrame
                                    .HorizontalAlignment = xlHAlignCenter
                                    .VerticalAlignment = xlVAlignCenter
                                    .Characters.Text = "2ND METRIC"
                                    .Characters.Font.Color = buttonFontColor
                                    .Characters.Font.Size = 9
                                    .Characters.Font.Name = fontName
                                    .MarginBottom = 0
                                End With
                                .Fill.ForeColor.RGB = buttonColour
                                .Line.ForeColor.RGB = buttonBorderColour
                                .Height = buttonObjPrev.Height / 2
                                .Width = buttonObjPrev.Width    '- 2
                                .Top = buttonObjPrev.Top
                                .Left = kuvaaja.Left + kuvaaja.Width - .Width
                                .Placement = xlFreeFloating
                            End With
                        End If


                        Set buttonObj = dataSheet.Shapes.AddShape(5, 342, 15, 118, 29)
                        With buttonObj
                            .Name = sheetID & "CTCTB_" & dataSheet.ChartObjects.Count
                            .OnAction = "toggleChartTypeForOneChart"
                            .Adjustments(1) = 0.1
                            With .TextFrame
                                .HorizontalAlignment = xlHAlignCenter
                                .VerticalAlignment = xlVAlignCenter
                                .Characters.Text = "CHANGE CHART TYPE"
                                .Characters.Font.Color = buttonFontColor
                                .Characters.Font.Size = 9
                                .Characters.Font.Name = fontName
                                .MarginBottom = 0
                            End With
                            .Fill.ForeColor.RGB = buttonColour
                            .Line.ForeColor.RGB = buttonBorderColour
                            .Height = buttonHeight - 4
                            .Width = buttonWidth * 2 + buttonSpaceBetween
                            .Top = kuvaaja.Top - .Height - 5
                            If buttonObjPrev Is Nothing Then
                                .Left = kuvaaja.Left + kuvaaja.Width - .Width
                            Else
                                .Left = buttonObjPrev.Left - .Width - buttonSpaceBetween
                            End If
                            .Placement = xlFreeFloating
                        End With



                    End With

                    rivi = kuvaaja.BottomRightCell.row + 5


                    If runningSheetRefresh = False Then
                        Application.ScreenUpdating = True
                        Application.ScreenUpdating = False
                    End If

                Next chartNum

            Next metricNum




            stParam1 = "8.32"

            If sendMode = True Then Call checkE(email, dataSource)

            If timeDimensionIncluded = False Then

                If (queryType = "SD" And visibleMetricColumnsCount > 1) Or profileCount >= 2 Or metricsCount >= 2 Or (metricsCount = 1 And doComparisons = 1) Then
                    rivi = rivi + 2


                    progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")*0.08")
                    Call updateProgress(progresspct, "Creating charts... Scatterplot")


                    Dim scatterChartUpperLeftCell As Range
                    Set scatterChartUpperLeftCell = .Cells(rivi, reportStartColumn + 1)





                    With Sheets("settings")

                        settingsSheetVisibility = .Visible
                        .Visible = xlSheetVisible
                        .Select
                        .Cells(1, 1).Select

                        .ChartObjects("chartScatter").Duplicate
                        Set kuvaaja = .ChartObjects(.ChartObjects.Count)
                        kuvaaja.Chart.Location where:=xlLocationAsObject, Name:=dataSheet.Name

                    End With


                    .Select
                    Set kuvaaja = .ChartObjects(.ChartObjects.Count)

                    kuvaaja.Top = scatterChartUpperLeftCell.Top
                    kuvaaja.Left = scatterChartUpperLeftCell.Left

                    kuvaaja.Placement = xlFreeFloating

                    Sheets("settings").Visible = settingsSheetVisibility

                    kuvaaja.Name = sheetID & "scatterPlot"

                    rivi = scatterChartUpperLeftCell.row
                    col = scatterChartUpperLeftCell.Column + 1


                    Call storeValue("YvaluesSel", 1, dataSheet)
                    Call storeValue("XvaluesSel", 1, dataSheet)

                    Call storeValue("firstDataCol", firstMetricCol, dataSheet)
                    Call storeValue("firstDataRow", resultStartRow, dataSheet)

                    If vriviData - resultStartRow > 255 Then
                        Call storeValue("lastDataRow", resultStartRow + 255, dataSheet)
                    Else
                        Call storeValue("lastDataRow", vriviData, dataSheet)
                    End If


                    .Cells(rivi + 3, col + 1).Name = sheetID & "_y1col"
                    .Cells(rivi + 3, col + 4).Name = sheetID & "_x1col"

                    .Range(.Cells(9, reportStartColumn + 2), Cells(25, reportStartColumn + 8)).Font.ColorIndex = 2   'use white font for config range

                    rivi2 = 0
                    columnCount = 0
                    .Cells(rivi + 3, col).Resize(UBound(columnInfoArr, 1), 5).Font.ColorIndex = 2
                    For col2 = 1 To UBound(columnInfoArr, 1)
                        arvo = ""
                        If columnInfoArr(col2, 1) <> vbNullString Then
                            If columnInfoArr(col2, 5) <> True Then
                                rivi2 = rivi2 + 1
                                columnCount = columnCount + 1
                                If profileCount > 1 Then arvo = columnInfoArr(col2, 3) & "|"
                                If columnInfoArr(col2, 6) = True Then
                                    arvo = arvo & "Change in " & columnInfoArr(col2, 1) & "|"
                                Else
                                    arvo = arvo & columnInfoArr(col2, 1) & "|"
                                End If

                                If segmentCount > 1 Then
                                    segmentNum = columnInfoArr(col2, 14)
                                    segmentName = segmentArr(segmentNum, 2)
                                    arvo = arvo & segmentName & "|"
                                End If

                                If queryType = "SD" Then arvo = arvo & Replace(columnInfoArr(col2, 4), " | ", "|")
                                If Right(arvo, 1) = "|" Then arvo = Left(arvo, Len(arvo) - 1)
                                .Cells(rivi + 3 + rivi2 - 1, col).value = arvo
                                .Cells(rivi + 3 + rivi2 - 1, col + 3).value = arvo
                                .Cells(rivi + 3 + rivi2 - 1, col + 1).value = col2
                                .Cells(rivi + 3 + rivi2 - 1, col + 4).value = col2
                            End If
                        End If
                    Next col2

                    '1 metric name
                    '2 invert conditional formatting
                    '3 profile name
                    '4 segmdimname
                    '5 hidden
                    '6 comparisons
                    '7 SD query Other category
                    '8 metric code
                    '9 metric submetric count
                    '10 data fetch error
                    '11 metricnum
                    '12 prof id

                    If columnCount > 0 Then

                        Set YvaluesDropdown = dataSheet.DropDowns.Add(192, 106.5, 140.25, 28.5)
                        With YvaluesDropdown
                            .Height = 20
                            .Width = 200
                            .Top = scatterChartUpperLeftCell.Top - 25
                            .Left = scatterChartUpperLeftCell.Left
                            .ListFillRange = dataSheet.Cells(rivi + 3, col).Resize(columnCount).Address
                            .LinkedCell = dataSheet.Range(fetchSettingAddress("YvaluesSel", dataSheet)).Address
                            .DropDownLines = columnCount
                            .OnAction = "setScatterPlotSeriesSD"
                            .Placement = xlFreeFloating
                            .Name = sheetID & "_YvaluesDropdown"
                        End With


                        Set XvaluesDropdown = dataSheet.DropDowns.Add(192, 106.5, 140.25, 28.5)
                        With XvaluesDropdown
                            .Height = 20
                            .Width = 200
                            .Top = scatterChartUpperLeftCell.Top - 25
                            .Left = YvaluesDropdown.Left + YvaluesDropdown.Width + 10
                            .ListFillRange = dataSheet.Cells(rivi + 3, col + 3).Resize(columnCount).Address
                            .LinkedCell = dataSheet.Range(fetchSettingAddress("XvaluesSel", dataSheet)).Address
                            .DropDownLines = columnCount
                            .OnAction = "setScatterPlotSeriesSD"
                            .Placement = xlFreeFloating
                            .Name = sheetID & "_XvaluesDropdown"
                        End With
                    End If

                    If doComparisons = 0 Or metricsCount = 1 Then
                        Call storeValue("yvaluesSel", 2, dataSheet)
                    Else
                        Call storeValue("yvaluesSel", 3, dataSheet)
                    End If
                    Call storeValue("xvaluesSel", 1, dataSheet)

                    Call setScatterPlotSeries(dimensionsCombinedCol)

                    scatterPlotCreated = True

                    chartsCreated = True


                    rivi = kuvaaja.BottomRightCell.row + 4

                End If



            End If



            Dim chartCategoriesDropdownUpperLeftCell As Range
            Dim chartCategoriesDropdownLowerRightCell As Range

            Set chartCategoriesDropdownUpperLeftCell = .Cells(.ChartObjects(1).TopLeftCell.row, .ChartObjects(1).BottomRightCell.Column - 1).Offset(-2)
            Set chartCategoriesDropdownLowerRightCell = chartCategoriesDropdownUpperLeftCell




            If timeDimensionIncluded = True Then
                Call storeValue("catSel", 1, dataSheet)
                vriviChart = vriviData
            Else

                If lineChartsCreated = True Then

                    Set buttonObj = dataSheet.Shapes.AddShape(5, 342, 15, 118, 29)
                    With buttonObj
                        .Name = sheetID & "chartCategoriesLabel"
                        .TextFrame.HorizontalAlignment = xlHAlignCenter
                        .TextFrame.VerticalAlignment = xlVAlignTop
                        ' .TextFrame.Characters.Text = "Data points in charts:"
                        .TextFrame.Characters.Font.ColorIndex = 1
                        .TextFrame.Characters.Font.Size = 9
                        .Fill.ForeColor.RGB = buttonColour
                        .Line.ForeColor.RGB = buttonBorderColour
                        .Height = buttonHeight - 4
                        .Width = buttonWidth * 2 + buttonSpaceBetween
                        .Top = ActiveSheet.Shapes(sheetID & "RemoveSheetButton").Top + ActiveSheet.Shapes(sheetID & "RemoveSheetButton").Height + buttonSpaceBetween
                        .Left = ActiveSheet.Shapes(sheetID & "RemoveSheetButton").Left + ActiveSheet.Shapes(sheetID & "RemoveSheetButton").Width - .Width - (ActiveSheet.Shapes(sheetID & "RemoveSheetButton").Width + buttonSpaceBetween) * 2

                        '   .Placement = xlFreeFloating
                    End With


                    Set buttonObj = dataSheet.Shapes.AddShape(152, 342, 15, 118, 29)
                    With buttonObj
                        .Name = sheetID & "chartCategoriesLabel2"
                        With .TextFrame
                            .HorizontalAlignment = xlHAlignCenter
                            .VerticalAlignment = xlVAlignCenter
                            .Characters.Text = "Data points in charts"

                            With .Characters.Font
                                .Size = 9
                                .Color = buttonFontColor
                                .Name = fontName
                            End With
                        End With
                        .Fill.ForeColor.RGB = buttonColour
                        .Line.ForeColor.RGB = buttonBorderColour
                        .Height = (buttonHeight - 4) / 2
                        .Width = buttonWidth * 2 + buttonSpaceBetween
                        .Top = ActiveSheet.Shapes(sheetID & "chartCategoriesLabel").Top
                        .Left = ActiveSheet.Shapes(sheetID & "chartCategoriesLabel").Left
                        '   .Placement = xlFreeFloating
                    End With


                    Set chartCategoriesDropdown = dataSheet.DropDowns.Add(192, 106.5, 140.25, 28.5)
                    With chartCategoriesDropdown
                        .Height = buttonObj.Height - 1
                        .Width = buttonObj.Width - 4
                        .Top = buttonObj.Top + buttonObj.Height - 1
                        .Left = buttonObj.Left + buttonObj.Width - .Width - 2
                        .ListFillRange = "vars!" & Range("chartCategoriesNumberList").Address
                        .LinkedCell = dataSheet.Range(fetchSettingAddress("catsel", dataSheet)).Address
                        .DropDownLines = Range("chartCategoriesNumberList").Rows.Count
                        .OnAction = "setChartCategories"
                        .Name = sheetID & "chartCategoriesDropdown"
                    End With


                    If vriviData - resultStartRow > 50 Then
                        Call storeValue("catSel", 4, dataSheet, sheetID & "_" & "catSel")
                        vriviChart = resultStartRow + 49
                    Else
                        Call storeValue("catSel", 1, dataSheet, sheetID & "_" & "catSel")
                        vriviChart = vriviData
                    End If

                End If

            End If

        ElseIf updatingPreviouslyCreatedSheet = True And createCharts = True Then
            seriesAdded = False

            chartCount = .ChartObjects.Count
            ReDim legendSizesArr(1 To chartCount * 10, 1 To 2)

            If vriviData - resultStartRow > 50 And timeDimensionIncluded = False Then
                vriviChart = resultStartRow + 49
            Else
                vriviChart = vriviData
            End If
            For Each kuvaaja In .ChartObjects
                If Left(kuvaaja.Name, Len(sheetID & "_C")) = sheetID & "_C" Then
                    metricNum = parseVarFromName(kuvaaja.Name, "M")
                    profID = parseVarFromName(kuvaaja.Name, "P")
                    profNum = parseVarFromName(kuvaaja.Name, "X")
                    chartNum = parseVarFromName(kuvaaja.Name, "N")
                    segmentNum = parseVarFromName(kuvaaja.Name, "SG")
                    '        chartNum = kuvaaja.Index

                    progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
                    Call updateProgress(progresspct, "Updating charts... " & chartNum)


                    With kuvaaja.Chart

                        If .SeriesCollection.Count > 0 Then

                            storedChartType = .SeriesCollection(1).ChartType

                            For seriesNum = 1 To .SeriesCollection.Count
                                If seriesNum > .SeriesCollection.Count Then Exit For
                                col = SeriesRange(.SeriesCollection(seriesNum)).Column
                                If queryType = "SD" Then
                                    If excelVersion > 11 Then
                                        .SeriesCollection(seriesNum).Name = "=" & Chr(39) & sheetName & Chr(39) & "!" & Range(ColumnLetter(col) & segmDimRow).MergeArea.Cells(1, 1).Address
                                    Else
                                        .SeriesCollection(seriesNum).Name = dataSheet.Cells(segmDimRow, col).MergeArea.Cells(1, 1)
                                    End If
                                ElseIf segmentCount > 1 Then
                                    If Left(.SeriesCollection(seriesNum).Name, 2) <> "%%" Then
                                        If excelVersion > 11 Then
                                            .SeriesCollection(seriesNum).Name = "=" & Chr(39) & sheetName & Chr(39) & "!" & Range(ColumnLetter(col) & segmentRow).MergeArea.Cells(1, 1).Address
                                        Else
                                            .SeriesCollection(seriesNum).Name = dataSheet.Cells(segmentRow, col).MergeArea.Cells(1, 1)
                                        End If
                                    End If
                                Else
                                    If excelVersion > 11 Then
                                        .SeriesCollection(seriesNum).Name = "=" & Chr(39) & sheetName & Chr(39) & "!" & Range(ColumnLetter(col) & profNameRow).MergeArea.Cells(1, 1).Address
                                    Else
                                        .SeriesCollection(seriesNum).Name = dataSheet.Cells(profNameRow, col).MergeArea.Cells(1, 1)
                                    End If
                                End If
                            Next seriesNum
                        End If

                        seriesNum = .SeriesCollection.Count

                        firstRowValueChanged = False
                        If seriesNum < 256 Then

                            For col = firstMetricCol To vsarData

                                If columnInfoArr(col, 8) = metricsArr(metricNum, 2) Then
                                    If (queryType = "SD" And columnInfoArr(col, 12) = profID And columnInfoArr(col, 14) = segmentNum) Or (queryType <> "SD" And segmentCount > 1 And columnInfoArr(col, 12) = profID) Or (queryType <> "SD" And segmentCount = 1) Then


                                        If columnInfoArr(col, 6) = False Then  'not change columns
                                            '                              If dataSheet.Cells(firstHeaderRow - 1, col).value <> "CHANGE" Then
                                            If columnInfoArr(col, 10) = vbNullString Then    'check for data errors
                                                '  If Left$(dataSheet.Cells(resultStartRow, col).value, 6) <> "Error:" Then

                                                If dataSheet.Cells(1, col).EntireColumn.Hidden = False Then
                                                    If InStr(1, fetchValue(kuvaaja.Name & "_dataColumns", ActiveSheet), "|" & col & "|") = 0 Then

                                                        seriesNum = seriesNum + 1

                                                        If seriesNum < 256 Then
                                                            If seriesNum > 1 Then .SeriesCollection.NewSeries

                                                            If dataSheet.Cells(resultStartRow, col).value = vbNullString Then
                                                                firstRowValueChanged = True
                                                                dataSheet.Cells(resultStartRow, col).value = 1
                                                            End If
                                                            seriesAdded = True
                                                            .ChartType = xlArea
                                                            .SeriesCollection(seriesNum).Values = dataSheet.Range(ColumnLetter(col) & lastHeaderRow + 1 & ":" & ColumnLetter(col) & vriviChart)
                                                            .SeriesCollection(seriesNum).XValues = dataSheet.Range(ColumnLetter(dimensionsCombinedCol) & lastHeaderRow + 1 & ":" & ColumnLetter(dimensionsCombinedCol) & vriviChart)

                                                            If debugMode = False Then On Error Resume Next
                                                            If queryType = "SD" Then
                                                                If excelVersion > 11 Then
                                                                    .SeriesCollection(seriesNum).Name = "=" & sheetName & "!" & Range(ColumnLetter(col) & segmDimRow).MergeArea.Cells(1, 1).Address
                                                                Else
                                                                    .SeriesCollection(seriesNum).Name = dataSheet.Cells(segmDimRow, col).MergeArea.Cells(1, 1)
                                                                End If
                                                            Else
                                                                If Left(.SeriesCollection(seriesNum).Name, 2) <> "%%" Then
                                                                    If excelVersion > 11 Then
                                                                        .SeriesCollection(seriesNum).Name = "=" & sheetName & "!" & Range(ColumnLetter(col) & profNameRow).MergeArea.Cells(1, 1).Address
                                                                    Else
                                                                        .SeriesCollection(seriesNum).Name = dataSheet.Cells(profNameRow, col).MergeArea.Cells(1, 1)
                                                                    End If
                                                                End If
                                                            End If
                                                            Call updateProgressAdditionalMessage(.SeriesCollection(seriesNum).Name)
                                                            If firstRowValueChanged = True Then dataSheet.Cells(resultStartRow, col).value = vbNullString

                                                            Call storeValue(kuvaaja.Name & "_dataColumns", fetchValue(kuvaaja.Name & "_dataColumns", ActiveSheet) & "|" & col & "|", ActiveSheet)

                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next col

                            If seriesAdded = True Then
                                .ChartType = xlLineMarkers

                                Call formatChart(kuvaaja)

                                If Range("doChartLegendSizeOptimization").value = True And .HasLegend = True Then
                                    If (metricNum = 1 Or (segmentCount > 1 And profNum = 1) Or (queryType = "SD" And segmentCount > 1)) Or dataSource = "FB" Then
                                        Call chartLegend(kuvaaja)
                                        legendSizesArr(profNum * metricNum * segmentNum, 1) = .Legend.Font.Size
                                        legendSizesArr(profNum * metricNum * segmentNum, 2) = .Legend.Left
                                        Debug.Print "Storing " & profNum & metricNum & segmentNum & " to " & profNum * metricNum * segmentNum
                                    Else
                                        With .Legend
                                            .Top = 0
                                            .Height = kuvaaja.Height
                                            .Font.Size = legendSizesArr(profNum * 1 * segmentNum, 1)
                                            .Left = legendSizesArr(profNum * 1 * segmentNum, 2)
                                        End With
                                        .PlotArea.Width = legendSizesArr(chartNum, 2) - 5 - .PlotArea.Left
                                    End If
                                End If
                            End If

                            If fetchValue("secondSeriesAddedToChart", dataSheet) = True Then
                                Call resetSecondarySeries(kuvaaja)
                                kuvaaja.Chart.SeriesCollection(1).ChartType = storedChartType
                                kuvaaja.Chart.Axes(xlValue, xlPrimary).AxisTitle.Font.ColorIndex = 1
                            End If

                        End If
                    End With
                End If
            Next

            Call setChartCategories

        End If


    End With

End Sub




