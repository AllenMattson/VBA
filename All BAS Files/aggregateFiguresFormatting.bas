Attribute VB_Name = "aggregateFiguresFormatting"
Option Private Module
Option Explicit


Sub fetchAggregateFiguresFormatting()

    On Error GoTo generalErrHandler
    If debugMode = True Then On Error GoTo 0
    Application.EnableCancelKey = xlErrorHandler

    Dim col As Long
    Dim dataSar As Long
    Dim groupSize As Integer
    Dim i As Long
    Dim rivi As Long
    Dim muutos As Variant
    Dim kuvaaja As ChartObject
    Dim YvaluesDropdown As Object
    Dim XvaluesDropdown As Object
    Dim scatterPlotCreated As Boolean
    Dim chartsCreated As Boolean
    Dim firstButtonLeft As Double
    Dim buttonNum As Integer
    Dim buttonObjPrev As Shape
    Dim doConditionalFormatting As Boolean

    Dim settingsSheetVisibility As Integer
    Dim warningText As String

    Dim asteriskCount As Integer
    asteriskCount = 0


    chartsCreated = False
    scatterPlotCreated = False

    With dataSheet

        Call updateProgress(progresspct, "Formatting...")
        For col = resultStartColumn To resultStartColumn + 2
            .Columns(ColumnLetter(col)).AutoFit
            If .Columns(ColumnLetter(col)).ColumnWidth > 20 Then .Columns(ColumnLetter(col)).ColumnWidth = 20
        Next col



        'don't wrap error messages
        .Cells(1, resultStartColumn + 3 + metricsCount + doComparisons * (metricNum - 1)).EntireColumn.WrapText = False


        stParam1 = "7.80"




        progresspct = progresspct + (100 - progresspct) * 0.05
        Call updateProgress(progresspct, "Sorting data...")


        Select Case sortType

        Case "alphabetic"
            If segmentCount > 1 Then
                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).sort key1:=.Cells(resultStartRow, resultStartColumn + 2), key2:=.Cells(resultStartRow, resultStartColumn + 3), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlAscending
            Else
                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).sort key1:=.Cells(resultStartRow, resultStartColumn + 2), key2:=.Cells(resultStartRow, resultStartColumn + 3), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlDescending
            End If
        Case "alphabetic desc"
            If segmentCount > 1 Then
                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).sort key1:=.Cells(resultStartRow, resultStartColumn + 2), key2:=.Cells(resultStartRow, resultStartColumn + 3), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlAscending
            Else
                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).sort key1:=.Cells(resultStartRow, resultStartColumn + 2), key2:=.Cells(resultStartRow, resultStartColumn + 3), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlDescending
            End If
        Case "metric desc"

            .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).sort key1:=.Cells(resultStartRow, firstMetricCol), key2:=.Cells(resultStartRow, resultStartColumn + 2), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlAscending

        Case "metric asc"

            .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).sort key1:=.Cells(resultStartRow, firstMetricCol), key2:=.Cells(resultStartRow, resultStartColumn + 2), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlAscending
        Case Else
            'no sort
        End Select


        Call storeValue("sortingCol", firstMetricCol, dataSheet)
        Call storeValue("sortType", sortType, dataSheet)
        Call storeValue("sortRange", .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, vsarData + 1)).Address, dataSheet)


        Range(.Cells(1, resultStartColumn), .Cells(vriviData, vsarData)).Name = sheetID & "_dataRange"

        stParam1 = "7.801"


        If doTotals Then

            .Cells(vriviData + 2, resultStartColumn).value = "Total"
            .Cells(vriviData + 3, resultStartColumn).value = "Average"


            With .Range(.Cells(vriviData + 2, resultStartColumn), .Cells(vriviData + 2 + 1, vsarData))
                ' .Font.Bold = True
                If excelVersion <= 11 Then
                    .Interior.ColorIndex = 50
                    .Font.ColorIndex = 2
                Else
                    .Interior.Color = Range("totalsColour").Interior.Color
                    .Font.Color = Range("totalsColour").Font.Color
                End If
            End With

            For metricNum = 1 To metricsCount
                For iterationNum = 1 To iterationsCount
                    dataSar = firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1) + (iterationNum - 1)
                    If (metricsArr(metricNum, 4) = 1 Or metricsArr(metricNum, 5) = "minus") And iterationNum = 1 And metricsArr(metricNum, 12) <> True Then
                        .Cells(vriviData + 2, dataSar).Formula = "=SUBTOTAL(109," & .Range(.Cells(resultStartRow, dataSar), .Cells(vriviData, dataSar)).Address & ")"
                    End If
                    .Cells(vriviData + 3, dataSar).Formula = "=SUBTOTAL(101," & .Range(.Cells(resultStartRow, dataSar), .Cells(vriviData, dataSar)).Address & ")"

                    If Not IsNumeric(.Cells(vriviData + 3, dataSar).value) Then .Cells(vriviData + 3, dataSar).value = vbNullString
                Next iterationNum
            Next metricNum

        End If







        If updatingPreviouslyCreatedSheet = True Then
            tempSheet.Range(Range(sheetID & "_dataRange").Address).Copy
            .Range(sheetID & "_dataRange").PasteSpecial (xlPasteFormats)
        End If
        Application.DisplayAlerts = False
        tempSheet.Delete
        Application.DisplayAlerts = True




        If Range("doColours").value = True Then
            If updatingPreviouslyCreatedSheet = False Then
                'mark row groups with colour
                Dim colour1 As Long
                Dim colour2 As Long
                colour1 = Range("rowColoursStart").Interior.Color
                colour2 = Range("rowColoursStart").Offset(1).Interior.Color
                groupSize = 3
                i = 0
                For rivi = resultStartRow To vriviData
                    i = i + 1
                    If excelVersion <= 11 Then
                        If i <= groupSize Then .Range(.Cells(rivi, resultStartColumn), .Cells(rivi, vsarData)).Interior.ColorIndex = 54
                    Else
                        If i <= groupSize Then
                            .Range(.Cells(rivi, resultStartColumn), .Cells(rivi, vsarData)).Interior.Color = colour1
                        Else
                            .Range(.Cells(rivi, resultStartColumn), .Cells(rivi, vsarData)).Interior.Color = colour2
                        End If
                    End If
                    If i = groupSize * 2 Then i = 0
                Next rivi
            End If


            'colour change columns
            If doComparisons = 1 Then
                For col = resultStartColumn + 4 To vsarData Step 2
                    For rivi = resultStartRow To vriviData
                        With .Cells(rivi, col)
                            muutos = .value
                            If muutos <> vbNullString Then
                                If muutos > 0.0049 Then
                                    .Interior.ColorIndex = 12
                                ElseIf muutos < -0.0049 Then
                                    .Interior.ColorIndex = 22
                                Else
                                    .Interior.ColorIndex = 2
                                End If
                            Else
                                .Interior.ColorIndex = 2
                            End If
                        End With
                    Next rivi
                Next col
            End If

        End If




        If Range("doAutofilter").value <> False Then
            If vriviData - resultStartRow > 5 Then
                Call updateProgress(progresspct, "Formatting... adding filters")
                .Range(.Cells(resultStartRow - 1, resultStartColumn), .Cells(resultStartRow - 1, vsarData)).AutoFilter
            End If
        End If





        'reduce profile ID font size
        .Range(ColumnLetter(resultStartColumn) & resultStartRow & ":" & ColumnLetter(resultStartColumn) & vriviData).Font.Size = 8

        stParam1 = "7.81"


        If runningSheetRefresh = False Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If


        Application.CutCopyMode = False


        If createCharts = True And updatingPreviouslyCreatedSheet = True Then
            If segmentCount > 1 And profileCount > 1 Then
                Call setScatterPlotSeries(resultStartColumn + 2, resultStartColumn + 3)
            ElseIf segmentCount > 1 Then
                Call setScatterPlotSeries(resultStartColumn + 3)
            Else
                Call setScatterPlotSeries(resultStartColumn + 2)
            End If
        End If


        If createCharts = True And (profileCount > 1 Or segmentCount > 1) And updatingPreviouslyCreatedSheet = False Then

            If metricsCount >= 2 Or (metricsCount = 1 And doComparisons = 1) Then


                progresspct = progresspct + (100 - progresspct) * 0.08
                Call updateProgress(progresspct, "Creating charts... Scatterplot")


                Dim scatterChartUpperLeftCell As Range
                Set scatterChartUpperLeftCell = .Cells(9, reportStartColumn + 2)


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
                kuvaaja.Top = scatterChartUpperLeftCell.Top

                stParam1 = "7.82"

                Sheets("settings").Visible = settingsSheetVisibility


                kuvaaja.Name = sheetID & "scatterPlot"
                kuvaaja.Placement = xlFreeFloating

                rivi = scatterChartUpperLeftCell.row
                col = scatterChartUpperLeftCell.Column


                Call storeValue("YvaluesSel", 1, dataSheet)
                Call storeValue("XvaluesSel", 1, dataSheet)

                Call storeValue("firstDataCol", firstMetricCol, dataSheet)
                Call storeValue("firstDataRow", resultStartRow, dataSheet)
                Call storeValue("lastDataRow", vriviData, dataSheet)


                .Cells(rivi + 3, col + 1).Name = sheetID & "_y1col"
                .Cells(rivi + 3, col + 4).Name = sheetID & "_x1col"

                .Range("C9:I25").Font.ColorIndex = 2  'use white font for config range

                stParam1 = "7.83"
                'storing metric info under chart


                For metricNum = 1 To metricsCount

                    With .Cells(rivi + 3 + metricNum - 1 + doComparisons * (metricNum - 1), col)
                        .value = metricsArr(metricNum, 1)
                        .Offset(, 3).value = metricsArr(metricNum, 1)

                        .Offset(, 1).value = firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1)
                        .Offset(, 4).value = firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1)

                        If doComparisons = 1 Then
                            stParam1 = "7.84"
                            .Offset(1).value = "Change in " & metricsArr(metricNum, 1)
                            .Offset(1, 3).value = "Change in " & metricsArr(metricNum, 1)

                            .Offset(1, 1).value = firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1) + 1
                            .Offset(1, 4).value = firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1) + 1

                        End If
                    End With
                Next metricNum

                stParam1 = "7.85"

                Dim selUpperLeftCell As Range
                Set selUpperLeftCell = .Cells(scatterChartUpperLeftCell.row + 11, scatterChartUpperLeftCell.Column - 2)

                Set YvaluesDropdown = dataSheet.DropDowns.Add(192, 106.5, 140.25, 28.5)
                With YvaluesDropdown
                    .Height = selUpperLeftCell.Height
                    .Width = (selUpperLeftCell.Width * 2) * 0.9
                    .Top = selUpperLeftCell.Top
                    .Left = selUpperLeftCell.Left + (selUpperLeftCell.Width * 2) * 0.05
                    .ListFillRange = ColumnLetter(col) & rivi + 3 & ":" & ColumnLetter(col) & rivi + 3 + metricsCount - 1 + doComparisons * metricsCount
                    .LinkedCell = dataSheet.Range(fetchSettingAddress("YvaluesSel", dataSheet)).Address
                    .DropDownLines = metricsCount + doComparisons * metricsCount
                    .OnAction = "setScatterPlotSeries"
                    .Placement = xlFreeFloating
                    .Name = sheetID & "_YvaluesDropdown"
                End With

                Set selUpperLeftCell = .Cells(scatterChartUpperLeftCell.row - 1, scatterChartUpperLeftCell.Column + 4)
                Set XvaluesDropdown = dataSheet.DropDowns.Add(192, 106.5, 140.25, 28.5)
                With XvaluesDropdown
                    .Height = selUpperLeftCell.Height
                    .Width = (selUpperLeftCell.Width * 2) * 0.9
                    .Top = selUpperLeftCell.Top - (selUpperLeftCell.Width * 2) * 0.05
                    .Left = selUpperLeftCell.Left
                    .ListFillRange = ColumnLetter(col + 3) & rivi + 3 & ":" & ColumnLetter(col + 3) & rivi + 3 + metricsCount - 1 + doComparisons * metricsCount
                    .LinkedCell = dataSheet.Range(fetchSettingAddress("xvaluesSel", dataSheet)).Address
                    .DropDownLines = metricsCount + doComparisons * metricsCount
                    .OnAction = "setScatterPlotSeries"
                    .Placement = xlFreeFloating
                    .Name = sheetID & "_XvaluesDropdown"
                End With

                stParam1 = "7.86"

                If doComparisons = 0 Or metricsCount = 1 Then
                    Call storeValue("yvaluesSel", 2, dataSheet)
                Else
                    Call storeValue("yvaluesSel", 3, dataSheet)
                End If
                Call storeValue("xvaluesSel", 1, dataSheet)

                stParam1 = "7.87"

                If segmentCount > 1 And profileCount > 1 Then
                    Call setScatterPlotSeries(resultStartColumn + 2, resultStartColumn + 3)
                    Call storeValue("rowLabelsCol", resultStartColumn + 2, dataSheet)
                    Call storeValue("rowLabelsCol2", resultStartColumn + 3, dataSheet)
                ElseIf segmentCount > 1 Then
                    Call setScatterPlotSeries(resultStartColumn + 3)
                    Call storeValue("rowLabelsCol", resultStartColumn + 3, dataSheet)
                Else
                    Call setScatterPlotSeries(resultStartColumn + 2)
                    Call storeValue("rowLabelsCol", resultStartColumn + 2, dataSheet)
                End If


                scatterPlotCreated = True

                chartsCreated = True

                rivi = kuvaaja.BottomRightCell.row + 4

            End If

            .Cells(1, 1).Copy
            Application.CutCopyMode = False
            If scatterPlotCreated = False Then rivi = 10

            For metricNum = 1 To metricsCount

                stParam1 = "7.88"

                progresspct = progresspct + (100 - progresspct) * 0.05
                Call updateProgress(progresspct, "Creating charts... " & metricsArr(metricNum, 1))


                With Sheets("settings")

                    .Visible = xlSheetVisible
                    .Select
                    .Cells(1, 1).Select

                    .ChartObjects("chartAgg").Duplicate
                    Set kuvaaja = .ChartObjects(.ChartObjects.Count)
                    kuvaaja.Chart.Location where:=xlLocationAsObject, Name:=dataSheet.Name

                End With

                stParam1 = "7.89"
                .Select
                Set kuvaaja = .ChartObjects(.ChartObjects.Count)


                kuvaaja.Top = .Cells(rivi, reportStartColumn + 2).Top
                kuvaaja.Left = .Cells(rivi, reportStartColumn + 2).Left

                kuvaaja.Placement = xlFreeFloating

                ' Sheets("settings").Visible = xlSheetHidden

                stParam1 = "7.90"
                kuvaaja.Name = sheetID & "Chart" & metricNum
                With kuvaaja.Chart
                    stParam1 = "7.900"
                    .ChartTitle.Text = UCase(metricsArr(metricNum, 1))
                    stParam1 = "7.901"
                    .SeriesCollection(1).Values = dataSheet.Range(dataSheet.Cells(resultStartRow, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1)), dataSheet.Cells(vriviData, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1)))
                    .SeriesCollection(1).Values = dataSheet.Range(dataSheet.Cells(resultStartRow, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1)), dataSheet.Cells(vriviData, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1)))
                    stParam1 = "7.902"

                    If profileCount > 1 And segmentCount > 1 Then
                        .SeriesCollection(1).XValues = dataSheet.Range(dataSheet.Cells(resultStartRow, resultStartColumn + 2), dataSheet.Cells(vriviData, resultStartColumn + 3))
                    ElseIf segmentCount > 1 Then
                        .SeriesCollection(1).XValues = dataSheet.Range(dataSheet.Cells(resultStartRow, resultStartColumn + 3), dataSheet.Cells(vriviData, resultStartColumn + 3))
                    Else
                        .SeriesCollection(1).XValues = dataSheet.Range(dataSheet.Cells(resultStartRow, resultStartColumn + 2), dataSheet.Cells(vriviData, resultStartColumn + 2))
                    End If

                    stParam1 = "7.903"
                    .SeriesCollection(1).Name = metricsArr(metricNum, 1)

                    stParam1 = "7.91"
                    If doComparisons = 1 Then
                        .SeriesCollection(2).Values = dataSheet.Range(dataSheet.Cells(resultStartRow, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1) + 1), dataSheet.Cells(vriviData, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1) + 1))

                        If profileCount > 1 And segmentCount > 1 Then
                            .SeriesCollection(2).XValues = dataSheet.Range(dataSheet.Cells(resultStartRow, resultStartColumn + 2), dataSheet.Cells(vriviData, resultStartColumn + 3))
                        ElseIf segmentCount > 1 Then
                            .SeriesCollection(2).XValues = dataSheet.Range(dataSheet.Cells(resultStartRow, resultStartColumn + 3), dataSheet.Cells(vriviData, resultStartColumn + 3))
                        Else
                            .SeriesCollection(2).XValues = dataSheet.Range(dataSheet.Cells(resultStartRow, resultStartColumn + 2), dataSheet.Cells(vriviData, resultStartColumn + 2))
                        End If

                        .SeriesCollection(2).Name = "Change in " & metricsArr(metricNum, 1)
                    Else
                        .SeriesCollection(2).Delete
                        .Legend.Delete

                    End If

                    stParam1 = "7.92"
                    .Axes(xlCategory).TickLabels.Orientation = xlTickLabelOrientationAutomatic
                    .ChartTitle.Left = .PlotArea.InsideLeft

                    chartsCreated = True

                End With
                rivi = rivi + 24
                stParam1 = "7.93"

                If runningSheetRefresh = False Then
                    Application.ScreenUpdating = True
                    Application.ScreenUpdating = False
                End If


            Next metricNum

        End If


        stParam1 = "7.94"


        If updatingPreviouslyCreatedSheet = False Then

            progresspct = progresspct + (100 - progresspct) * 0.05
            Call updateProgress(progresspct, "Inserting buttons...")

            buttonObj.Delete

            firstButtonLeft = Round(.Cells(1, reportStartColumn + 4).Left + buttonSpaceBetween)

            Dim createdButtonNum As Integer
            createdButtonNum = 1

            For buttonNum = 1 To 6

                Set buttonObj = dataSheet.Shapes.AddShape(5, 10, 10, 200, 40)  '5=msoShapeRoundedRectangle
                With buttonObj

                    .Adjustments(1) = 0.1

                    With .TextFrame
                        .HorizontalAlignment = xlHAlignCenter
                        .VerticalAlignment = xlVAlignCenter
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .MarginBottom = 0
                        .Characters.Font.Color = buttonFontColor
                        .Characters.Font.Size = 8
                        .Characters.Font.Name = "Calibri Ligth"
                    End With

                    .Fill.ForeColor.RGB = buttonColour
                    .Line.ForeColor.RGB = buttonBorderColour
                    .Height = buttonHeight
                    .Width = buttonWidth
                    .Top = buttonTop
                    .Left = firstButtonLeft + (createdButtonNum - 1) * (buttonWidth + buttonSpaceBetween)
                    Select Case buttonNum
                    Case 1
                        .OnAction = "refreshDataOnSelectedSheet"
                        .TextFrame.Characters.Text = "REFRESH"
                        .Name = sheetID & "RefreshButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 2
                        .OnAction = "createPPTofActiveSheet"
                        .TextFrame.Characters.Text = "CREATE PPT"
                        .Name = sheetID & "CreatePPTButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 3
                        .OnAction = "exportReportToExcel"
                        .TextFrame.Characters.Text = "EXPORT TO EXCEL"
                        .Name = sheetID & "ExportExcelButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 4
                        If excelVersion > 11 Then
                            .OnAction = "changeConditionalFormatting"
                            .TextFrame.Characters.Text = "TABLE FORMAT"
                            .Name = sheetID & "condFormButton"
                            createdButtonNum = createdButtonNum + 1
                        Else
                            .Delete
                        End If
                    Case 5
                        .OnAction = "selectActiveReportInQuerystorage"
                        .TextFrame.Characters.Text = "MODIFY QUERY"
                        .Name = sheetID & "ModifyQueryButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 6
                        .OnAction = "removeSheet"
                        .TextFrame.Characters.Text = "REMOVE SHEET"
                        .Fill.ForeColor.RGB = buttonColourRed
                        .Name = sheetID & "RemoveSheetButton"
                        createdButtonNum = createdButtonNum + 1
                    End Select

                End With
            Next buttonNum



            Set buttonObjPrev = buttonObj


            If profileCount > 1 Or segmentCount > 1 Then
                stParam1 = "7.945"

                'Sorting button

                Set buttonObj = dataSheet.Shapes.AddShape(5, 10, 10, 200, 40)
                With buttonObj
                    .Name = sheetID & "sortButton1"
                    .OnAction = "changeSort"
                    With .TextFrame
                        .HorizontalAlignment = xlHAlignCenter
                        .VerticalAlignment = xlVAlignBottom
                        .MarginBottom = 0
                        .MarginLeft = 0
                        .MarginRight = 0
                        With .Characters.Font
                            .Color = buttonFontColor
                            .Name = "Calibri Ligth"
                            .Size = 7
                        End With
                    End With
                    Select Case sortType
                    Case "alphabetic"
                        .TextFrame.Characters.Text = "Sorted alphabetically"
                    Case "alphabetic desc"
                        .TextFrame.Characters.Text = "Sorted alphabetically (desc)"
                    Case "metric desc"
                        .TextFrame.Characters.Text = "Sorted by 1st metric (desc)"
                    Case "metric asc"
                        .TextFrame.Characters.Text = "Sorted by 1st metric (asc)"
                    End Select
                    .Fill.ForeColor.RGB = buttonColour    ' RGB(255, 255, 255)
                    .Line.ForeColor.RGB = buttonBorderColour    ' buttonBorderColourLight
                    .Height = buttonHeight    '- 4
                    .Width = buttonWidth * 2 + buttonSpaceBetween

                    'place under other buttons
                    '.Top = buttonObjPrev.Top + buttonObjPrev.Height + buttonSpaceBetween
                    '.Left = buttonObjPrev.Left + buttonObjPrev.Width - .Width

                    'place on right
                    .Top = buttonObjPrev.Top
                    .Left = buttonObjPrev.Left + buttonObjPrev.Width + buttonSpaceBetween

                    '     .Placement = xlFreeFloating
                End With
                Set buttonObj = dataSheet.Shapes.AddShape(152, 342, 15, 118, 29)
                With buttonObj
                    .Name = sheetID & "sortButton2"
                    .OnAction = "changeSort"
                    With .TextFrame
                        .HorizontalAlignment = xlHAlignCenter
                        .VerticalAlignment = xlVAlignCenter
                        .Characters.Text = "CHANGE SORT"
                        .Characters.Font.ColorIndex = 1
                        .Characters.Font.Size = 9
                        .Characters.Font.Color = buttonFontColor
                    End With
                    .Fill.ForeColor.RGB = buttonColour    ' RGB(242, 242, 242)
                    .Line.ForeColor.RGB = buttonBorderColour
                    .Height = ActiveSheet.Shapes(sheetID & "sortButton1").Height / 2
                    .Width = ActiveSheet.Shapes(sheetID & "sortButton1").Width    '- 2

                    .Top = ActiveSheet.Shapes(sheetID & "sortButton1").Top    '+ ActiveSheet.Shapes(sheetID & "sortButton1").Height / 2
                    .Left = ActiveSheet.Shapes(sheetID & "sortButton1").Left    ' + 1


                    '     .Placement = xlFreeFloating
                End With




            End If

        End If




        stParam1 = "7.95"



        .Range("B1:B20").Font.Bold = False
        .Cells(2, 2).Font.Bold = True
        .Cells(3, 2).Font.Bold = False
        .Cells(5, 3).Font.Bold = False
        .Cells(5, 4).Font.Bold = False

        .Cells(1, reportStartColumn + 1).Resize(20, 1).Font.Bold = False
        .Cells(2, reportStartColumn + 1).Font.Bold = True
        .Cells(3, reportStartColumn + 1).Font.Bold = False
        .Cells(5, reportStartColumn + 2).Font.Bold = False
        .Cells(5, reportStartColumn + 3).Font.Bold = False

        aika1 = Timer - aika1
        If usingMacOSX = False Or forceOSXmode = True Then .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Fetching and processing data took " & Round(aika1, 1) & " s."

        If reportContainsSampledData = True Then
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = "This report contains sampled data (sampling done by Google)."
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
        End If

        If givemaxResultsPerQueryWarning = True Then
            If queryCount > 1 Then
                warningText = "At least one of the queries would have returned more rows than is the limit for this type of queries."
            Else
                warningText = "The query for this report would have returned more rows than is the limit for this type of queries."
            End If
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = warningText
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With

            warningText = vbNullString
            '  If maxResultsPerQuery < 1000000 Then warningText = warningText & "To get more complete results, increase this limit and rerun the query."
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = warningText
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
        End If

        If showNoteStr <> vbNullString Then
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = showNoteStr
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
        End If



        stParam1 = "7.96"


        'cond formatting
        If Range("conditionalFormattingType").value <> "none" Then doConditionalFormatting = True



        'cond formatting
        '  If Range("conditionalFormattingType").value <> "none" Then

        Dim condFormType As String
        Dim doColours As Boolean
        condFormType = Range("conditionalFormattingType").value
        doColours = Range("doColours").value = True
        Dim invertColoursCols As String
        Dim midPointAtZeroCols As String
        invertColoursCols = "|"
        midPointAtZeroCols = "|"

        progresspct = progresspct + (100 - progresspct) * 0.05
        Call updateProgress(progresspct, "Formatting... applying conditional formatting")

        For col = resultStartColumn + 3 To vsarData
            If doComparisons = 0 Or (doComparisons = 1 And (col - resultStartColumn - 3) Mod 2 = 0) Then
                If doConditionalFormatting Then Call applyConditionalFormatting(.Range(ColumnLetter(col) & resultStartRow & ":" & ColumnLetter(col) & vriviData), condFormType, CBool(columnInfoArr(col, 2)))
            Else
                Select Case comparisonValueType
                Case "perc", "abs"
                    Call applyConditionalFormatting(.Range(.Cells(resultStartRow, col), .Cells(vriviData, col)), "colouring", CBool(columnInfoArr(col, 2)), True)
                    midPointAtZeroCols = midPointAtZeroCols & col & "|"
                Case "val"
                    Call applyConditionalFormatting(.Range(.Cells(resultStartRow, col), .Cells(vriviData, col)), "colouring", CBool(columnInfoArr(col, 2)))
                End Select

            End If
            If CBool(columnInfoArr(col, 2)) Then invertColoursCols = invertColoursCols & col & "|"
        Next col

        Call storeValue("condFormType", condFormType, dataSheet)
        Call storeValue("invertColoursCols", invertColoursCols, dataSheet)
        Call storeValue("midPointAtZeroCols", midPointAtZeroCols, dataSheet)
        '   End If

        Call storeValue("firstMetricCol", resultStartColumn + 3, dataSheet)
        Call storeValue("lastMetricCol", vsarData, dataSheet)
        Call storeValue("firstDataRow", resultStartRow, dataSheet)
        Call storeValue("lastDataRow", vriviData, dataSheet)







        If doComparisons = 1 Then
            With .Cells(vriviData + 2 + IIf(doTotals, 3, 0), resultStartColumn + 1)
                .Font.Size = 9
                Select Case comparisonValueType
                Case "perc"
                    If comparisonType = "previous" Then
                        If timeDimensionIncluded = False And segmDimIsTime = False Then
                            .value = "* change from previous period of same length (" & startDate2 & "-" & endDate2 & "), as percentage"
                        Else
                            .value = "* change from previous " & mostGranularTimeDimension & " (%)"
                        End If
                    ElseIf comparisonType = "yearly" Then
                        .value = "* change from the same period a year earlier (%)"
                    Else
                        .value = "* change from " & startDate2 & "-" & endDate2 & " (%)"
                    End If
                Case "abs"
                    If comparisonType = "previous" Then
                        If timeDimensionIncluded = False And segmDimIsTime = False Then
                            .value = "* change from previous period of same length (" & startDate2 & "-" & endDate2 & ")"
                        Else
                            .value = "* change from previous " & mostGranularTimeDimension
                        End If
                    ElseIf comparisonType = "yearly" Then
                        .value = "* change from the same period a year earlier"
                    Else
                        .value = "* change from " & startDate2 & "-" & endDate2
                    End If
                Case "val"
                    If comparisonType = "previous" Then
                        If timeDimensionIncluded = False And segmDimIsTime = False Then
                            .value = "* comparison value from previous period of same length (" & startDate2 & "-" & endDate2 & ")"
                        Else
                            .value = "* comparison value from previous " & mostGranularTimeDimension
                        End If
                    ElseIf comparisonType = "yearly" Then
                        .value = "* comparison value from the same period a year earlier"
                    Else
                        .value = "* comparison value from " & startDate2 & "-" & endDate2
                    End If
                End Select
            End With
            asteriskCount = 1
        End If


        If sumAllProfiles Then
            .Cells(resultStartRow, resultStartColumn).value = "Summed results for " & UBound(profilesArr) & " " & referToProfilesAs
            With .Cells(resultStartRow, resultStartColumn).Font
                .Bold = True
                .Size = 11
            End With
            With .Cells(vriviData + (asteriskCount * 2) + 2 + IIf(doTotals, 3, 0), resultStartColumn + 1)
                .Font.Size = 9
                .value = "Results contain data from these " & UBound(profilesArr) & " " & referToProfilesAs & ":"
                .Offset(1, 1) = capitalizeFirstLetter(referToAccountsAsSing)
                .Offset(1, 2) = capitalizeFirstLetter(referToProfilesAsSing)
                .Offset(1, 3) = capitalizeFirstLetter(referToProfilesAsSing) & " ID"
                With .Offset(profNum, 1).Resize(UBound(profilesArr) + 1, 3)
                    .Font.Size = 9
                    .NumberFormat = ""
                End With
                For profNum = 1 To UBound(profilesArr)
                    .Offset(profNum + 1, 1) = profilesArr(profNum, 1)
                    .Offset(profNum + 1, 2) = profilesArr(profNum, 2)
                    .Offset(profNum + 1, 3) = profilesArr(profNum, 3)
                Next profNum
            End With
            asteriskCount = asteriskCount + 1
        End If




        stParam1 = "7.97"

        If updatingPreviouslyCreatedSheet = False Then
            .Rows(resultStartRow).Select
            ActiveWindow.FreezePanes = True
        End If


        If vsarData < resultStartColumn + 17 Then
            col = vsarData
        Else
            col = resultStartColumn + 17
        End If

        If vriviData < resultStartRow + 20 Then
            rivi = resultStartRow + 20
        Else
            rivi = resultStartRow + 40
        End If

        .Range(.Cells(resultStartRow - 1, resultStartColumn), .Cells(vriviData, vsarData)).Name = sheetID & "_PPTrange"

        progresspct = 99
        Call updateProgress(progresspct, vbNullString)
        DoEvents


        stParam1 = "7.98"

        Application.Calculation = calculationSetting
        If Application.Calculation <> xlAutomatic Then .Calculate
        .Select

        If runningSheetRefresh = False Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If

        Application.EnableEvents = True

    End With



    Call hideProgressBox
    Call removeTempsheet
    Application.StatusBar = False



    Exit Sub

generalErrHandler:

    stParam2 = "REPORTERROR " & Err.Number & "|" & Err.Description & "|" & Application.StatusBar

    Debug.Print "REPORT ERROR: " & stParam1 & " " & stParam2
    'Call checkE(email, dataSource, True)

    If Err.Number = 18 Then
        Call hideProgressBox
        Call removeTempsheet
        End
    End If

    Resume Next
End Sub

