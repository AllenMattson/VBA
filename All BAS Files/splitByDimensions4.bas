Attribute VB_Name = "splitByDimensions4"
Option Private Module
Option Explicit


Sub fetchFigureSplitByDimensionsFormatting()


    Dim rivi As Long
    Dim sar As Long
    Dim col As Long

    Dim i As Long

    Dim dataSar As Long
    Dim dataRivi As Long
    Dim div As Variant
    Dim num As Variant
    Dim muutos As Variant
    Dim arvo As Variant
    Dim arvo1 As Variant

    Dim dataRng As Range
    Dim comparisonDataRng As Range
    Dim minValue As Double
    Dim maxValue As Double
    Dim comparisonMinValue As Double
    Dim comparisonMaxValue As Double
    Dim isComparisonCol As Boolean

    Dim buttonNum As Integer
    Dim buttonObjPrev As Object
    Dim dimensionNum As Long
    Dim firstButtonLeft As Double

    Dim arr2 As Variant

    Dim segmDimCategoriesStr As String
    Dim valueCount As Long

    Dim hideColumn As Boolean

    Dim tempStr As String

    Dim doConditionalFormatting As Boolean
    Dim doManualColourFormatting As Boolean


    Dim segmDimNum As Long
    Dim profName As String

    Dim colour1 As Long
    Dim colour2 As Long
    Dim colour3 As Long
    Dim colour4 As Long

    Dim warningText As String

    Dim asteriskCount As Integer
    asteriskCount = 0

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    With dataSheet




        stParam1 = "8.18"

        progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
        Call updateProgress(progresspct, "Calculating...")

        'convert negative SDO values to zeroes
        For col = firstMetricCol To vsarData
            If columnInfoArr(col, 7) Then
                arr = .Range(.Cells(1, col), .Cells(vriviData, col)).value
                For rivi = resultStartRow To vikarivi(.Cells(1, col))
                    If arr(rivi, 1) < 0 Then arr(rivi, 1) = 0
                Next rivi
                .Range(.Cells(1, col), .Cells(vriviData, col)).value = arr
            End If
        Next col


        Dim arr3 As Variant    'results
        Dim arr4 As Variant    'change comparison column

        For col = firstMetricCol To vsarData

            If col >= sortingCol Then Exit For

            arr = .Range(.Cells(1, col), .Cells(vriviData, col)).value
            arr2 = tempSheet.Range(tempSheet.Cells(1, col), tempSheet.Cells(vriviData, col)).value
            ReDim arr3(1 To UBound(arr), 1 To 1)
            ReDim arr4(1 To UBound(arr), 1 To 1)

            If columnInfoArr(col, 6) Then arr4 = .Range(.Cells(1, col - 1), .Cells(vriviData, col - 1)).value

            '      If arr(firstHeaderRow - 1, 1) = "CHANGE" Then arr4 = .Range(.Cells(1, col - 1), .Cells(vriviData, col - 1)).value

            For rivi = 1 To resultStartRow - 1
                arr3(rivi, 1) = .Cells(rivi, col).value
            Next rivi

            For rivi = resultStartRow To vikarivi(.Cells(1, col))
                If columnInfoArr(col, 9) = 1 Then
                    arr3(rivi, 1) = arr(rivi, 1)
                Else
                    div = arr2(rivi, 1)
                    Select Case metricsArr(columnInfoArr(col, 11), 5)
                    Case "minus"    'lost impressions
                        If arr(rivi, 1) = vbNullString Then arr(rivi, 1) = 0
                        If div = vbNullString Then div = 0
                        arr3(rivi, 1) = arr(rivi, 1) - div
                    Case "div&minus&minusone", "div&minus&plus&minusone"   'viral amplification %
                        If div > 0 Then
                            arr3(rivi, 1) = (arr(rivi, 1) / div) - 1
                        Else
                            arr3(rivi, 1) = vbNullString
                        End If
                    Case Else
                        If div <> 0 Then
                            arr3(rivi, 1) = arr(rivi, 1) / div
                        Else
                            arr3(rivi, 1) = vbNullString
                        End If
                    End Select
                End If
                If columnInfoArr(col, 6) Then  'comp column
                    '   If arr(firstHeaderRow - 1, 1) = "CHANGE" Then
                    arvo = arr3(rivi, 1)
                    arvo1 = arr4(rivi, 1)
                    muutos = vbNullString
                    Select Case comparisonValueType
                    Case "perc"
                        If arvo1 <> vbNullString Then
                            If arvo <> 0 And arvo <> vbNullString Then
                                muutos = arvo1 / arvo - 1
                            End If
                        ElseIf arvo = 0 Or arvo = vbNullString Then    'both zeroes/blanks
                            muutos = vbNullString
                        Else    'newer value zero
                            muutos = -1
                        End If
                        arr3(rivi, 1) = muutos
                    Case "abs"
                        If arvo = vbNullString Then arvo = 0
                        If arvo1 = vbNullString Then arvo1 = 0
                        arr3(rivi, 1) = arvo1 - arvo
                    Case "val"
                        arr3(rivi, 1) = arvo
                    End Select
                End If
            Next rivi

            .Range(.Cells(1, col), .Cells(vriviData, col)).value = arr3
        Next col





        stParam1 = "8.19"


        If Not updatingPreviouslyCreatedSheet And Not rawDataReport Then
            'header row formatting
            With .Range(.Cells(profIDRow, firstMetricCol), .Cells(profIDRow, vsarData))
                .NumberFormat = "@"
                .Font.Size = 8
            End With
            .Range(.Cells(accountNameRow, firstMetricCol), .Cells(accountNameRow, vsarData)).NumberFormat = "@"
            .Range(.Cells(profNameRow, firstMetricCol), .Cells(profNameRow, vsarData)).NumberFormat = "@"
            If queryType = "SD" Then
                With .Range(.Cells(segmDimRow, firstMetricCol), .Cells(segmDimRow, vsarData))
                    .NumberFormat = "@"
                    .Font.Size = 9
                End With
                If segmDimCount > 1 Then
                    For segmDimNum = 1 To segmDimCount
                        With .Range(.Cells(segmDimRow + segmDimNum, firstMetricCol), .Cells(segmDimRow + segmDimNum, vsarData))
                            .NumberFormat = "@"
                            .Font.Size = 9
                        End With
                    Next segmDimNum
                End If
            End If
        End If



        stParam1 = "8.20"


        'calculate sorting column
        '  If sortType <> "alphabetic" Then
        If metricsArr(1, 4) > 1 And Not rawDataReport Then
            Select Case metricsArr(1, 5)
            Case "div", "div*86400", "1000*div", "div1000000", "div1000"
                For dataRivi = resultStartRow To vriviData
                    div = .Cells(dataRivi, sortingCol + 1).value
                    If div > 0 Then
                        .Cells(dataRivi, sortingCol).value = .Cells(dataRivi, sortingCol).value / div
                    Else
                        .Cells(dataRivi, sortingCol).value = vbNullString
                    End If
                Next dataRivi
            Case "div*86400-exits"  'avg time on page
                For dataRivi = resultStartRow To vriviData
                    num = .Cells(dataRivi, sortingCol).value
                    div = .Cells(dataRivi, sortingCol + 1).value - .Cells(dataRivi, sortingCol + 2).value
                    If div <> 0 Then
                        .Cells(dataRivi, sortingCol).value = num / div
                    Else
                        .Cells(dataRivi, sortingCol).value = vbNullString
                    End If
                Next dataRivi
            End Select
        End If
        '   End If


        If sendMode = True Then Call checkE(email, dataSource)


        If updatingPreviouslyCreatedSheet = False And Not rawDataReport Then
            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Formatting... column borders")

            For col = firstMetricCol To vsarData
                With .Range(.Cells(lastHeaderRow, col), .Cells(vriviData, col)).Borders(xlEdgeLeft)
                    If columnInfoArr(col, 15) = "L" Then
                        .ColorIndex = 16
                        .weight = xlThin
                        .LineStyle = xlContinuous
                    ElseIf queryType = "SD" Then
                        If dataSheet.Cells(segmDimRow, col).value <> vbNullString Then
                            .ColorIndex = 16
                            .weight = xlHairline
                        End If
                    End If
                End With
            Next col

        End If


        stParam1 = "8.21"




        progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
        Call updateProgress(progresspct, "Inserting headers...")

        If Not sumAllProfiles And Not rawDataReport Then
            If dataSource = "GA" Then
                .Cells(profIDRow, dimensionsCombinedCol).value = "Profile ID"
                .Cells(accountNameRow, dimensionsCombinedCol).value = "Account"
                .Cells(profNameRow, dimensionsCombinedCol).value = "Profile"
            ElseIf dataSource = "AW" Then
                .Cells(profIDRow, dimensionsCombinedCol).value = "Account ID"
                .Cells(accountNameRow, dimensionsCombinedCol).value = "MCC"
                .Cells(profNameRow, dimensionsCombinedCol).value = "Account"
            ElseIf dataSource = "AC" Then
                .Cells(profIDRow, dimensionsCombinedCol).value = "Account ID"
                .Cells(accountNameRow, dimensionsCombinedCol).value = "Account"
                .Cells(profNameRow, dimensionsCombinedCol).value = "Sub-account"
                '            ElseIf dataSource = "FB" Then
                '                .Cells(profIDRow, dimensionsCombinedCol).value = "ID"
                '                '.Cells(accountNameRow, dimensionsCombinedCol).value = "Account"
                '                .Cells(profNameRow, dimensionsCombinedCol).value = "Sub-account"
            End If
        End If


        If Not rawDataReport Then
            For dimensionNum = 1 To dimensionsCount
                If dimensionNum = 1 Then
                    dimensionHeadersCombined = dimensionsArr(dimensionNum, 1)
                Else
                    dimensionHeadersCombined = dimensionHeadersCombined & " | " & dimensionsArr(dimensionNum, 1)
                End If
            Next dimensionNum


            If queryType = "SD" Then
                .Cells(metricNameRow, dimensionsCombinedCol).value = "Metric"
                If segmentCount > 1 Then .Cells(segmentRow, dimensionsCombinedCol).value = "Segment"
            ElseIf segmentCount > 1 Then
                .Cells(metricNameRow, dimensionsCombinedCol).value = "Metric"
                .Cells(segmentRow, dimensionsCombinedCol).value = "Segment"
            Else
                If groupByMetric Then
                    .Cells(metricNameRow, dimensionsCombinedCol).value = "Metric"
                Else
                    .Cells(metricNameRow, dimensionsCombinedCol).value = dimensionHeadersCombined
                End If
            End If


            'SEGM DIM HEADERS
            If queryType = "SD" Then
                .Cells(segmDimRow, dimensionsCombinedCol).value = segmDimNameCombDisp
                If segmDimCount > 1 Then
                    For segmDimNum = 1 To segmDimCount
                        .Cells(segmDimRow + segmDimNum, dimensionsCombinedCol).value = Range("segmDimNameDisp" & segmDimNum).value
                    Next segmDimNum
                    .Cells(segmDimRow + segmDimCount, dimensionsCombinedCol).value = .Cells(segmDimRow + segmDimCount, dimensionsCombinedCol).value & " " & ChrW(8594)    'add arrow
                Else
                    .Cells(segmDimRow, dimensionsCombinedCol).value = segmDimNameDisp & " " & ChrW(8594)  'add arrow
                End If
            End If
        End If


        For dimensionNum = 1 To dimensionsCount
            .Cells(lastHeaderRow, resultStartColumn + dimensionNum - 1).value = dimensionsArr(dimensionNum, 1)
        Next dimensionNum



        If sendMode = True Then Call checkE(email, dataSource)


        progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
        Call updateProgress(progresspct, "Formatting...")



        If updatingPreviouslyCreatedSheet = False And Not rawDataReport Then
            'column widths
            For col = resultStartColumn To resultStartColumn + dimensionsCount
                With .Cells(1, col).EntireColumn
                    .AutoFit
                    If .ColumnWidth > 20 Then .ColumnWidth = 20
                    .HorizontalAlignment = xlLeft
                End With
            Next col
        End If



        stParam1 = "8.22"




        If updatingPreviouslyCreatedSheet = False Then

            firstButtonLeft = Round(.Cells(1, reportStartColumn + 4).Left + buttonSpaceBetween)

            buttonObj.Delete

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Inserting buttons...")

            Dim createdButtonNum As Integer
            createdButtonNum = 1

            For buttonNum = 1 To 7

                '   Set buttonObj = dataSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
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
                        If createCharts = True Then
                            .OnAction = "changeChartType"
                            .TextFrame.Characters.Text = "CHART TYPE"
                            .Name = sheetID & "ChartTypeButton"
                            createdButtonNum = createdButtonNum + 1
                        Else
                            .Delete
                        End If
                    Case 5
                        If excelVersion > 11 And Not rawDataReport Then
                            .OnAction = "changeConditionalFormatting"
                            .TextFrame.Characters.Text = "TABLE FORMAT"
                            .Name = sheetID & "condFormButton"
                            createdButtonNum = createdButtonNum + 1
                        Else
                            .Delete
                        End If
                    Case 6
                        .OnAction = "selectActiveReportInQuerystorage"
                        .TextFrame.Characters.Text = "MODIFY QUERY"
                        .Name = sheetID & "ModifyQueryButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 7
                        .OnAction = "removeSheet"
                        .TextFrame.Characters.Text = "REMOVE SHEET"
                        .Fill.ForeColor.RGB = buttonColourRed
                        .Name = sheetID & "RemoveSheetButton"
                        createdButtonNum = createdButtonNum + 1
                    End Select


                End With

            Next buttonNum
            Set buttonObjPrev = buttonObj



            stParam1 = "8.225"


            'Sorting button
            If Not rawDataReport Then

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
                    .Height = buttonHeight - 4
                    .Width = buttonWidth * 2 + buttonSpaceBetween
                    .Top = buttonObjPrev.Top + buttonObjPrev.Height + buttonSpaceBetween
                    .Left = buttonObjPrev.Left + buttonObjPrev.Width - .Width
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
                        With .Characters.Font

                            .Size = 9
                            .Color = buttonFontColor
                            .Name = fontName
                        End With
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



            If sendMode = True Then Call checkE(email, dataSource)


            stParam1 = "8.23"

            Call storeValue("firstDataRow", resultStartRow, dataSheet)

            Call storeValue("lastDataRow", vriviData, dataSheet)

            If vriviData - resultStartRow > 50 And timeDimensionIncluded = False Then
                Call storeValue("catSel", 4, dataSheet, sheetID & "_" & "catSel")
                vriviChart = resultStartRow + 49
            Else
                Call storeValue("catSel", 1, dataSheet, sheetID & "_" & "catSel")
                vriviChart = vriviData
            End If

        Else    'updating previously created sheet

            Call storeValue("lastDataRow", vriviData, dataSheet)
            If createCharts = True Then
                If fetchValue("catSel", dataSheet) = 1 Then Call setChartCategories
            End If

        End If



        stParam1 = "8.24"
        If sendMode = True Then Call checkE(email, dataSource)

        If Not rawDataReport Then
            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Sorting data...")

            Select Case sortType

            Case "alphabetic"

                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, sortingCol)).sort key1:=.Cells(firstHeaderRow + 5, dimensionsCombinedCol), order1:=Excel.XlSortOrder.xlAscending

            Case "alphabetic desc"

                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, sortingCol)).sort key1:=.Cells(firstHeaderRow + 5, dimensionsCombinedCol), order1:=Excel.XlSortOrder.xlDescending

            Case "metric desc"

                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, sortingCol)).sort key1:=.Cells(firstHeaderRow + 5, sortingCol), key2:=.Cells(firstHeaderRow + 5, dimensionsCombinedCol), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlAscending

            Case "metric asc"

                .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, sortingCol)).sort key1:=.Cells(firstHeaderRow + 5, sortingCol), key2:=.Cells(firstHeaderRow + 5, dimensionsCombinedCol), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlAscending
            Case Else
                'no sort
            End Select
        End If



        If runningSheetRefresh = False Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If


        progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.1")
        Call updateProgress(progresspct, "Formatting...")


        stParam1 = "8.241"
        'hide sorting columns
        If Not rawDataReport Then
            For i = 1 To metricsArr(1, 4)
                .Cells(1, sortingCol + i - 1).EntireColumn.Hidden = True
            Next i
        End If


        Call storeValue("firstCol", resultStartColumn, dataSheet)
        If rawDataReport Then
            Call storeValue("lastCol", sortingCol - 1, dataSheet)
        Else
            Call storeValue("lastCol", sortingCol + metricsArr(1, 4) - 1, dataSheet)
            Call storeValue("sortingCol", sortingCol, dataSheet)
            Call storeValue("sortType", sortType, dataSheet)
        End If

        If dimensionsCount = 1 Then
            Call storeValue("sortRange", .Range(.Cells(resultStartRow, resultStartColumn + 1), .Cells(vriviData, sortingCol + metricsArr(1, 4) - 1)).Address, dataSheet)
        Else
            Call storeValue("sortRange", .Range(.Cells(resultStartRow, resultStartColumn), .Cells(vriviData, sortingCol + metricsArr(1, 4) - 1)).Address, dataSheet)
        End If

        If dimensionCountMetricIncluded = True And queryType = "SD" Then
            stParam1 = "8.242"
            'hide Other column
            For col = resultStartColumn To vsarData
                If columnInfoArr(col, 7) = True Then
                    If columnInfoArr(col, 8) = "numberofkeywords" Or columnInfoArr(col, 8) = "numberofreferringsites" Or columnInfoArr(col, 8) = "numberofreferringdomains" Or columnInfoArr(col, 8) = "numberoflandingpages" Or columnInfoArr(col, 8) = "numberofpages" Or columnInfoArr(col, 8) = "numberofcountries" Or columnInfoArr(col, 8) = "numberofcities" Then
                        .Cells(1, col).EntireColumn.Hidden = True
                    End If
                End If
            Next col
        End If


        stParam1 = "8.25"
        If sendMode = True Then Call checkE(email, dataSource)


        '      vsarData = vikasar(.Cells(firstHeaderRow, 1))


        'date formatting
        progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
        Call updateProgress(progresspct, "Formatting... date formatting")
        For dimensionNum = 1 To dimensionsCount
            .Cells(1, resultStartColumn + dimensionNum - 1).EntireColumn.NumberFormat = ""
            If dimensionsCount = 1 Then .Cells(1, resultStartColumn + dimensionsCount).EntireColumn.NumberFormat = ""
            If LCase(dimensionsArr(dimensionNum, 2)) = "date" Then
                For rivi = resultStartRow To vriviData
                    With .Cells(rivi, resultStartColumn + dimensionNum - 1)
                        arvo = CStr(.value)
                        arvo = DateSerial(Left(arvo, 4), Mid(arvo, 5, 2), Right(arvo, 2))
                        .value = arvo
                        If dimensionsCount = 1 Then dataSheet.Cells(rivi, resultStartColumn + dimensionsCount).value = arvo
                    End With
                Next rivi
                .Cells(1, resultStartColumn + dimensionNum - 1).EntireColumn.NumberFormatLocal = Range("numformatdate").NumberFormatLocal
                If dimensionsCount = 1 Then .Cells(1, resultStartColumn + dimensionsCount).EntireColumn.NumberFormatLocal = Range("numformatdate").NumberFormatLocal
            End If
        Next dimensionNum


        'if just one dimension in the query no need to repeat values in two columns
        Dim dimColRemoved As Boolean
        dimColRemoved = False
        If rawDataReport Then
            'shift all dimension columns to right by one
            For dimensionNum = dimensionsCount To 1 Step -1
                With .Range(.Cells(1, resultStartColumn + dimensionNum - 1), .Cells(vriviData, resultStartColumn + dimensionNum - 1))
                    .Copy .Offset(, 1)
                    If dimensionNum = 1 Then .ClearContents
                End With
            Next dimensionNum
        ElseIf dimensionsCount = 1 Then
            With .Cells(1, resultStartColumn).EntireColumn
                .ClearContents
                If excelVersion <= 11 Then
                    .Interior.ColorIndex = 2
                Else
                    .Interior.Color = Range("sheetBackgroundColour").Interior.Color
                End If
            End With
            .Cells(lastHeaderRow + 1, resultStartColumn).value = dimensionsArr(1, 1) & " " & ChrW(8594)    'add dimension label and arrow
            .Cells(1, resultStartColumn).EntireColumn.AutoFit
            dimColRemoved = True
            Call storeValue("firstCol", resultStartColumn + 1, dataSheet)
        Else
            ' .Cells(1, dimensionsCombinedCol).EntireColumn.Hidden = True
        End If







        stParam1 = "8.26"
        If sendMode = True Then Call checkE(email, dataSource)

        If Range("doColours").value = True And Not rawDataReport Then

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Formatting... row colours")

            'mark row groups with colour
            Dim groupSize As Long
            Dim firstColToFormat As Long
            If dimColRemoved = True Then
                firstColToFormat = resultStartColumn + 1
            Else
                firstColToFormat = resultStartColumn
            End If
            groupSize = 3
            i = 0


            colour1 = Range("rowColoursStart").Interior.Color
            colour2 = Range("rowColoursStart").Offset(1).Interior.Color

            For rivi = resultStartRow To vriviData

                i = i + 1
                If excelVersion <= 11 Then
                    If i > groupSize Then .Range(.Cells(rivi, firstColToFormat), .Cells(rivi, vsarData)).Interior.ColorIndex = 54
                Else
                    If i <= groupSize Then
                        .Range(.Cells(rivi, firstColToFormat), .Cells(rivi, vsarData)).Interior.Color = colour1
                    Else
                        .Range(.Cells(rivi, firstColToFormat), .Cells(rivi, vsarData)).Interior.Color = colour2
                    End If
                End If
                If i = groupSize Or i = groupSize * 2 Then
                    With .Range(.Cells(rivi, firstColToFormat), .Cells(rivi, vsarData)).Borders(xlEdgeBottom)
                        .ColorIndex = 16
                        .weight = xlHairline
                    End With
                End If
                If i = groupSize * 2 Then i = 0

            Next rivi


            If sendMode = True Then Call checkE(email, dataSource)

            doManualColourFormatting = False
            If excelVersion <= 11 Then doManualColourFormatting = True

            'colour change columns
            If doComparisons = 1 Then

                progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
                Call updateProgress(progresspct, "Formatting... colouring change percentages")

                Dim addChangeLabelNew As Boolean

                addChangeLabelNew = True
                For col = firstMetricCol + 1 To vsarData Step 2
                    If nonTimeDimensionIncluded = False Then
                        addChangeLabelNew = False
                    ElseIf queryType = "SD" Then
                        If .Cells(segmDimRow, col - 1).value = "Other" And segmDimIsTime = True Then addChangeLabelNew = False
                    End If
                    For rivi = resultStartRow To vikarivi(.Cells(1, col - 1))
                        muutos = .Cells(rivi, col).value
                        If muutos = vbNullString Then
                            If .Cells(rivi, col - 1).value <> vbNullString And .Cells(rivi, col - 1).value <> 0 And addChangeLabelNew = True Then
                                With .Cells(rivi, col)
                                    .value = "NEW"
                                    .HorizontalAlignment = xlRight
                                End With
                                '.Cells(rivi, col).Interior.ColorIndex = 18
                            End If
                        ElseIf doManualColourFormatting = True Then
                            If muutos > 0.0049 Then
                                .Cells(rivi, col).Interior.ColorIndex = 18
                            ElseIf muutos < -0.0049 Then
                                .Cells(rivi, col).Interior.ColorIndex = 19
                            End If
                        End If
                    Next rivi
                Next col
            End If
        End If


        If updatingPreviouslyCreatedSheet = True And Not rawDataReport Then
            rivi = tempSheet.Range(sheetID & "_tempDataRangeFormats").Rows.Count
            col = Range(sheetID & "_tempDataRangeFormats").Columns.Count
            If vriviData < rivi Then rivi = vriviData - lastHeaderRow
            If vsarData - resultStartColumn + 1 < col Then col = vsarData - firstMetricCol + 1
            If rivi > 0 Then
                tempSheet.Range(sheetID & "_tempDataRangeFormats").Copy
                .Range(sheetID & "_dataRange").Resize(rivi, col).PasteSpecial (xlPasteFormats)
            End If
            ThisWorkbook.Names(sheetID & "_tempDataRangeFormats").Delete
        End If


        Application.DisplayAlerts = False
        tempSheet.Delete
        Application.DisplayAlerts = True



        progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
        Call updateProgress(progresspct, "Formatting... error labels")

        'transfer no data markings etc to upmost data row
        If Not rawDataReport Or vriviData <= 1 Then
            For col = firstMetricCol To vsarData
                If columnInfoArr(col, 10) <> vbNullString And Not columnInfoArr(col, 6) Then
                    .Cells(resultStartRow, col).value = columnInfoArr(col, 10)   ' .Cells(firstHeaderRow - 1, col).value
                    If queryType = "SD" Then .Cells(segmDimRow, col).value = columnInfoArr(col, 10)   ' .Cells(firstHeaderRow - 1, col).value
                    .Cells(resultStartRow, col).Font.Italic = True
                    .Cells(firstHeaderRow - 1, col).ClearContents
                End If
            Next col
        End If



        stParam1 = "8.27"

        If sendMode = True Then Call checkE(email, dataSource)




        If queryType = "SD" Then

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Formatting... removing empty columns")

            visibleMetricColumnsCount = vsarData - firstMetricCol + 1

            'remove empty columns
            For col = vsarData To firstMetricCol Step -1

                Set dataRng = .Range(.Cells(resultStartRow, col), .Cells(vriviData, col))
                minValue = Application.Min(dataRng)
                maxValue = Application.max(dataRng)
                segmDimCategoriesStr = columnInfoArr(col, 4)

                hideColumn = False
                If doComparisons = 0 Then
                    isComparisonCol = False
                Else
                    isComparisonCol = columnInfoArr(col, 6)
                    If isComparisonCol = True Then
                        Set comparisonDataRng = .Range(.Cells(resultStartRow, col - 1), .Cells(vriviData, col))
                    Else
                        Set comparisonDataRng = .Range(.Cells(resultStartRow, col), .Cells(vriviData, col + 1))
                    End If
                    comparisonMinValue = Application.Min(comparisonDataRng)
                    comparisonMaxValue = Application.max(comparisonDataRng)
                End If


                'If columnInfoArr(col, 8) <> columnInfoArr(col - 1, 8) Then
                'leave first col for each metric visible
                'ElseIf columnInfoArr(col, 10) <> vbNullString Then
                'leave data fetch errors visible
                'Else
                If doComparisons = 0 And segmDimCategoriesStr = vbNullString Then       'HIDE DATA COLUMN WHEN NO SD LABEL
                    hideColumn = True
                ElseIf doComparisons = 0 And Abs(minValue) < (1 / 100000) And Abs(maxValue) < (1 / 100000) Then
                    hideColumn = True
                ElseIf doComparisons = 1 And comparisonMinValue = 0 And comparisonMaxValue = 0 Then
                    hideColumn = True
                End If
                If hideColumn = True And Left(.Cells(resultStartRow, col).value, 6) <> "Error:" Then
                    Cells(1, col).EntireColumn.Hidden = True
                    columnInfoArr(col, 5) = True
                    visibleMetricColumnsCount = visibleMetricColumnsCount - 1
                End If
            Next col

            'mark no data found
            For col = vsarData To firstMetricCol Step -1
                If columnInfoArr(col, 5) = False Then  'not hidden
                    If columnInfoArr(col, 6) = False Then  'not comparisons
                        Set dataRng = .Range(.Cells(resultStartRow, col), .Cells(vriviData, col))
                        valueCount = Application.Count(dataRng)
                        If valueCount = 0 Then
                            With .Cells(lastHeaderRow + 1, col)
                                If .value = "" Then .value = "No data found"
                                .Font.Italic = True
                            End With
                        End If
                    End If
                End If
            Next col
        End If
        '    vsarData = vikasar(.Cells(firstHeaderRow, 1))



        stParam1 = "8.28"


        If sendMode = True Then Call checkE(email, dataSource)


        If Range("doColours").value = True And updatingPreviouslyCreatedSheet = False And Not rawDataReport Then

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Formatting... colouring metric headers")

            'mark changes between metrics with colours
            Dim rowToCheck As Long
            Dim rowToColour As Long

            i = 0

            colour1 = Range("headerColoursStart").Interior.Color
            colour2 = Range("headerColoursStart").Offset(1).Interior.Color
            colour3 = Range("headerColoursStart").Offset(2).Interior.Color
            colour4 = Range("headerColoursStart").Offset(3).Interior.Color

            rowToCheck = metricNameRow
            rowToColour = metricNameRow

            For col = firstMetricCol To vsarData Step 1 + doComparisons
                If col > firstMetricCol Then

                    If .Cells(metricNameRow, col).value <> .Cells(metricNameRow, col - 1 - doComparisons).value Then
                        If groupByMetric = True Then
                            i = i + 1
                        ElseIf .Cells(profIDRow, col).value <> .Cells(profIDRow, col - 1 - doComparisons).value Then
                            i = 0
                        Else
                            i = i + 1
                        End If
                    End If

                End If

                If excelVersion <= 11 Then
                    With .Cells(metricNameRow, col).Interior
                        Select Case i Mod 4
                        Case 0
                            .ColorIndex = 13
                        Case 1
                            .ColorIndex = 14
                        Case 2
                            .ColorIndex = 13
                        Case 3
                            .ColorIndex = 14
                        End Select
                    End With
                Else
                    With .Cells(metricNameRow, col).Interior
                        Select Case i Mod 4
                        Case 0
                            .Color = colour1
                        Case 1
                            .Color = colour2
                        Case 2
                            .Color = colour3
                        Case 3
                            .Color = colour4
                        End Select
                    End With
                End If


                If doComparisons = 1 Then
                    With .Cells(metricNameRow, col + 1).Interior
                        Select Case i Mod 4
                        Case 0
                            .ColorIndex = 13
                        Case 1
                            .ColorIndex = 14
                        Case 2
                            .ColorIndex = 13
                        Case 3
                            .ColorIndex = 14
                        End Select
                    End With
                End If
            Next col

        End If





        stParam1 = "8.29"
        If sendMode = True Then Call checkE(email, dataSource)



        If Range("doCellMerging").value = True And Not rawDataReport Then

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Formatting... merging header cells")

            Dim firstCol As Long
            firstCol = firstMetricCol
            'merging metric headers
            Application.DisplayAlerts = False
            If groupByMetric = True Or doComparisons = 1 Or queryType = "SD" Or segmentCount > 1 Then
                For col = firstMetricCol + 1 To vsarData Step 1
                    If .Cells(metricNameRow, col).value <> .Cells(metricNameRow, col - 1).value Then
                        If col - 1 > firstCol Then .Range(.Cells(metricNameRow, firstCol), .Cells(metricNameRow, col - 1)).Merge
                        firstCol = col
                    End If
                Next col
                If vsarData > firstCol Then .Range(.Cells(metricNameRow, firstCol), .Cells(metricNameRow, vsarData)).Merge
            End If



            If Not sumAllProfiles Then
                Call updateProgress(progresspct, "Formatting... merging account headers")

                'merge account names
                firstCol = firstMetricCol
                i = 0
                .Cells(accountNameRow, firstMetricCol).Interior.ColorIndex = 13
                For col = firstMetricCol + 1 To vsarData
                    If .Cells(accountNameRow, col).value <> .Cells(accountNameRow, col - 1).value Then
                        If col - 1 > firstCol Then .Range(.Cells(accountNameRow, firstCol), .Cells(accountNameRow, col - 1)).Merge
                        firstCol = col
                        i = i + 1
                    End If
                    If dataSource <> "FB" Then    'for FB, this row shows just Page or App
                        Select Case i Mod 2
                        Case 0
                            .Cells(accountNameRow, col).Interior.ColorIndex = 13
                        Case 1
                            .Cells(accountNameRow, col).Interior.ColorIndex = 14
                        End Select
                    End If
                Next col
                If vsarData > firstCol Then .Range(.Cells(accountNameRow, firstCol), .Cells(accountNameRow, vsarData)).Merge


                Call updateProgress(progresspct, "Formatting... merging profile headers")


                'merge profile names and ids

                firstCol = firstMetricCol
                .Cells(profNameRow, firstMetricCol).Interior.ColorIndex = 13
                i = 0
                For col = firstMetricCol + 1 To vsarData
                    If CStr(.Cells(profIDRow, col).value) <> CStr(.Cells(profIDRow, col - 1).value) Then
                        If col - 1 > firstCol Then
                            .Range(.Cells(profIDRow, firstCol), .Cells(profIDRow, col - 1)).Merge
                            If .Cells(profIDRow, col).value <> "*" Then .Range(.Cells(profNameRow, firstCol), .Cells(profNameRow, col - 1)).Merge
                        End If
                        firstCol = col
                        i = i + 1
                    End If

                    Select Case i Mod 2
                    Case 0
                        .Cells(profNameRow, col).Interior.ColorIndex = 13
                    Case 1
                        .Cells(profNameRow, col).Interior.ColorIndex = 14
                    End Select
                    If .Cells(accountNameRow, col).value <> .Cells(accountNameRow, col - 1).value And .Cells(accountNameRow, col).value <> vbNullString Then
                        .Cells(profNameRow, col).Interior.ColorIndex = .Cells(accountNameRow, col).Interior.ColorIndex
                        If .Cells(profNameRow, col).Interior.ColorIndex = 13 Then
                            i = 2
                        Else
                            i = 3
                        End If
                    End If
                Next col
                If vsarData > firstCol Then .Range(.Cells(profIDRow, firstCol), .Cells(profIDRow, vsarData)).Merge
                If vsarData > firstCol Then .Range(.Cells(profNameRow, firstCol), .Cells(profNameRow, vsarData)).Merge
            End If

            If segmentCount > 1 And segmentRow < lastHeaderRow Then
                Call updateProgress(progresspct, "Formatting... merging segment headers")
                firstCol = firstMetricCol
                i = 0
                .Cells(segmentRow, firstMetricCol).Interior.ColorIndex = 13
                For col = firstMetricCol + 1 To vsarData
                    If .Cells(segmentRow, col).value <> .Cells(segmentRow, col - 1).value Then
                        If col - 1 > firstCol Then .Range(.Cells(segmentRow, firstCol), .Cells(segmentRow, col - 1)).Merge
                        firstCol = col
                        i = i + 1
                    End If
                    Select Case i Mod 2
                    Case 0
                        .Cells(segmentRow, col).Interior.ColorIndex = 13
                    Case 1
                        .Cells(segmentRow, col).Interior.ColorIndex = 14
                    End Select
                Next col
                If vsarData > firstCol Then .Range(.Cells(segmentRow, firstCol), .Cells(segmentRow, vsarData)).Merge
            End If

        End If


        stParam1 = "8.30"
        If sendMode = True Then Call checkE(email, dataSource)


        'cond formatting
        If Range("conditionalFormattingType").value <> "none" Then doConditionalFormatting = True


        If Not rawDataReport Then

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Formatting... applying conditional formatting")

            Dim condFormType As String
            Dim doColours As Boolean
            condFormType = Range("conditionalFormattingType").value
            doColours = Range("doColours").value = True
            Dim invertColoursCols As String
            Dim midPointAtZeroCols As String
            invertColoursCols = "|"
            midPointAtZeroCols = "|"

            For col = firstMetricCol To vsarData
                If dataSheet.Cells(1, col).EntireColumn.Hidden = False Then
                    If Not columnInfoArr(col, 6) Then
                        If doConditionalFormatting Then Call applyConditionalFormatting(.Range(.Cells(resultStartRow, col), .Cells(vriviData, col)), condFormType, CBool(columnInfoArr(col, 2)))
                        If CBool(columnInfoArr(col, 2)) Then invertColoursCols = invertColoursCols & col & "|"
                    Else
                        Select Case comparisonValueType
                        Case "perc", "abs"
                            If doManualColourFormatting = False And doColours Then Call applyConditionalFormatting(.Range(.Cells(resultStartRow, col), .Cells(vriviData, col)), "colouring", CBool(columnInfoArr(col, 2)), True)
                            midPointAtZeroCols = midPointAtZeroCols & col & "|"
                        Case "val"
                            If doManualColourFormatting = False And doColours Then Call applyConditionalFormatting(.Range(.Cells(resultStartRow, col), .Cells(vriviData, col)), "colouring", CBool(columnInfoArr(col, 2)))
                        End Select
                        If CBool(columnInfoArr(col, 2)) Then invertColoursCols = invertColoursCols & col & "|"
                    End If
                End If
            Next col
            Call storeValue("condFormType", condFormType, dataSheet)
            Call storeValue("invertColoursCols", invertColoursCols, dataSheet)
            Call storeValue("midPointAtZeroCols", midPointAtZeroCols, dataSheet)
        End If



        Call storeValue("firstMetricCol", firstMetricCol, dataSheet)
        Call storeValue("lastMetricCol", vsarData, dataSheet)





        Range(.Cells(1, resultStartColumn), .Cells(vriviData, vsarData)).Name = sheetID & "_dataRange"

        stParam1 = "8.301"
        If doTotals And Not rawDataReport Then
            Call updateProgress(progresspct, "Calculating totals and averages...")

            .Cells(vriviData + 2, resultStartColumn).value = "Total"
            .Cells(vriviData + 3, resultStartColumn).value = "Average"


            With .Range(.Cells(vriviData + 2, resultStartColumn), .Cells(vriviData + 2 + 1, vsarData))

                '  .Font.Bold = True
                If excelVersion <= 11 Then
                    .Interior.ColorIndex = 40
                    '.Font.ColorIndex = 2
                Else
                    .Interior.Color = Range("totalsColour").Interior.Color
                    .Font.Color = Range("totalsColour").Font.Color
                End If
                .Name = sheetID & "_totals"
            End With



            For dataSar = firstMetricCol To vsarData

                metricNum = columnInfoArr(dataSar, 11)
                If columnInfoArr(dataSar, 6) = True Then
                    iterationNum = 2
                Else
                    iterationNum = 1
                End If

                If (metricsArr(metricNum, 4) = 1 Or metricsArr(metricNum, 5) = "minus") And iterationNum = 1 And metricsArr(metricNum, 12) <> True Then
                    .Cells(vriviData + 2, dataSar).Formula = "=SUBTOTAL(109," & .Range(.Cells(resultStartRow, dataSar), .Cells(vriviData, dataSar)).Address & ")"
                End If
                .Cells(vriviData + 3, dataSar).value = Application.Average(.Range(.Cells(resultStartRow, dataSar), .Cells(vriviData, dataSar)))
                .Cells(vriviData + 3, dataSar).Formula = "=SUBTOTAL(101," & .Range(.Cells(resultStartRow, dataSar), .Cells(vriviData, dataSar)).Address & ")"
                If Not IsNumeric(.Cells(vriviData + 3, dataSar).value) Then .Cells(vriviData + 3, dataSar).value = vbNullString

            Next dataSar

        End If



        stParam1 = "8.305"

        If Range("doAutofilter").value <> False And Not rawDataReport Then
            If vriviData - resultStartRow > 5 Then
                Call updateProgress(progresspct, "Formatting... adding filters")
                If dimensionsCount = 1 Then
                    col = resultStartColumn + 1
                Else
                    col = resultStartColumn
                End If
                If queryType = "SD" Then
                    .Range(.Cells(resultStartRow - 1, col), .Cells(segmDimRow, vsarData)).AutoFilter
                ElseIf segmentCount > 1 Then
                    .Range(.Cells(resultStartRow - 1, col), .Cells(segmentRow, vsarData)).AutoFilter
                ElseIf groupByMetric = True And profileCount > 1 Then
                    .Range(.Cells(resultStartRow - 1, col), .Cells(profNameRow, vsarData)).AutoFilter
                Else
                    .Range(.Cells(resultStartRow - 1, col), .Cells(metricNameRow, vsarData)).AutoFilter
                End If
            End If
        End If





        If createCharts And Not rawDataReport Then Call fetchFigureSplitByDimensionsFormattingCharts



        stParam1 = "8.33"



        '  Call updateProgress(progresspct, "Formatting... removing change labels")

        'remove change markings
        '        For col = firstMetricCol To vsarData
        '            If .Cells(firstHeaderRow - 1, col).value = "CHANGE" Then .Cells(firstHeaderRow - 1, col).ClearContents
        '        Next col




        If debugMode = False Then On Error Resume Next


        .Cells(1, reportStartColumn + 1).Resize(20, 1).Font.Bold = False
        '   .Cells(2, reportStartColumn + 1).Font.Bold = True
        .Cells(3, reportStartColumn + 1).Font.Bold = False

        If usingMacOSX = False Or forceOSXmode = True Then
            aika1 = Timer - aika1
            .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Fetching and processing data took " & Round(aika1, 1) & " s."
        End If

        If dataSource = "GW" Then
            If dateRangeRestriction <> "" Then
                With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                    If Left(dateRangeRestriction, 4) = "last" Then
                        warningText = "This report type is restricted to data from the last " & Replace(dateRangeRestriction, "last", "") & " days"
                    ElseIf dateRangeRestriction = "latest" Then
                        warningText = "This report type always shows the latest available data. The selected date range is ignored."
                    Else
                        warningText = dateRangeRestriction
                    End If
                    .value = warningText
                    .Font.Size = 8
                    .Font.ColorIndex = 16
                End With
            End If
        End If


        If reportContainsSampledData = True Then
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = "This report contains sampled data (sampling done by Google)."
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
        End If

        If givemaxResultsPerQueryWarning = True Then
            If queryCount > 1 Then
                warningText = "At least one of the queries would have returned more rows than the limit set on the " & configsheet.Name & " sheet (" & maxResults1 & ")."
            Else
                warningText = "The query for this report would have returned more rows than the limit set on the " & configsheet.Name & " sheet (" & maxResults1 & ")."
            End If
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = warningText
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
            warningText = vbNullString
            If maxResults < 1000000 Then warningText = warningText & "To get more complete results, increase this limit and rerun the query."
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = warningText
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
        End If


        If giveUniqueSumWarning = True Then
            warningText = "For this report, Supermetrics Data Grabber had to calculate sums of some figures that contained unique counts."
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = warningText
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
            warningText = "The resulting deduplicated values may be somewhat higher than they should be."
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = warningText
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
            If dateDimensionIncluded = False Then
                warningText = "Splitting the results by date should return the accurate figures."
                With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                    .value = warningText
                    .Font.Size = 8
                    .Font.ColorIndex = 16
                End With
            End If
        End If

        If showNoteStr <> vbNullString Then
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                .value = showNoteStr
                .Font.Size = 8
                .Font.ColorIndex = 16
            End With
        End If


        With .Range(.Cells(3, reportStartColumn + 1), Cells(10, reportStartColumn + 1))
            .Font.Size = 8
            .Font.ColorIndex = 56
        End With

        If doComparisons = 1 And Not rawDataReport Then
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

        Dim tempCell As Range

        If sumAllProfiles Then
            If Not rawDataReport Then
                With .Cells(accountNameRow, firstMetricCol)
                    .value = "Summed results for " & UBound(profilesArr) & " " & referToProfilesAs
                    '.Font.Bold = True
                End With
            End If
            If rawDataReport Then
                Set tempCell = .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
            Else
                Set tempCell = .Cells(vriviData + 2 + IIf(doTotals, 3, 0) + 2 * asteriskCount, resultStartColumn + 1)
            End If
            With tempCell
                .Font.Size = 9
                .value = "Results contain data from these " & UBound(profilesArr) & " " & referToProfilesAs & ":"
                .Offset(1, 1) = capitalizeFirstLetter(referToAccountsAsSing)
                .Offset(1, 2) = capitalizeFirstLetter(referToProfilesAsSing)
                .Offset(1, 3) = capitalizeFirstLetter(referToProfilesAsSing) & " ID"
                With .Offset(1, 1).Resize(UBound(profilesArr) + 1, 3)
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








        'inform if the number of rows is too large
        If vriviData >= rowLimit Then
            If rowLimit <= 65536 Then
                MsgBox "The number of rows returned by the query exceeded the row limit of " & rowLimit & ". The macro will return as many rows as possible, but some important information may be left out. For better results, lower the " & Chr$(34) & "Limit for result rows per profile" & Chr$(34) & " setting or remove some dimensions or profiles from the query." & vbCrLf & vbCrLf & "(Note that the Excel 2007/2010 version of this tool supports 1048576 rows.)"
            Else
                MsgBox "The number of rows returned by the query exceeded the row limit of " & rowLimit & ". The macro will return as many rows as possible, but some important information may be left out. For better results, lower the " & Chr$(34) & "Limit for result rows per profile" & Chr$(34) & " setting or remove some dimensions or profiles from the query."
            End If
        End If






        If updatingPreviouslyCreatedSheet = False And Not rawDataReport Then
            .Rows(resultStartRow).Select
            ActiveWindow.FreezePanes = True
        End If


        stParam1 = "8.34"

        If sendMode = True Then Call checkE(email, dataSource)



        If vsarData < resultStartColumn + 25 Then
            col = vsarData
        Else
            col = resultStartColumn + 25
        End If


        If vriviData < resultStartRow + 20 Then
            rivi = resultStartRow + 20
        Else
            rivi = resultStartRow + 40
        End If

        .Range(.Cells(firstHeaderRow, resultStartColumn), .Cells(vriviData + 3, vsarData)).Name = sheetID & "_PPTrange"

        progresspct = 99
        Call updateProgress(progresspct, vbNullString)



        Application.Calculation = calculationSetting
        .Calculate
        .Select
        .Cells(resultStartRow, resultStartColumn).Select

        If runningSheetRefresh = False Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If

    End With


    stParam1 = "8.35"


    stParam1 = "8.36"


    If IsArray(columnInfoArr) Then Erase columnInfoArr

    stParam1 = "8.37"

    Call hideProgressBox

    Call removeTempsheet

    Application.StatusBar = False

    stParam1 = "8.38"

    Exit Sub

generalErrHandler:


    stParam2 = "REPORTFORMATERROR"
    stParam2 = "REPORTFORMATERROR " & Err.Number & "|" & Err.Description & "|" & Application.StatusBar
    Debug.Print "REPORTFORMATERROR: " & stParam1 & " " & stParam2
    'Call checkE(email, dataSource, True)


    If Err.Number = 18 Then
        Call hideProgressBox
        Call removeTempsheet
        End
    End If


    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If debugMode = False Then Resume Next
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




End Sub




