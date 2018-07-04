Attribute VB_Name = "splitByDimensions2"
Option Private Module
Option Explicit


Sub fetchFiguresSplitByDimensionsStartQueriesAndFormatSheet()

    Dim sar As Long
    Dim buttonNum As Integer
    Dim firstButtonLeft As Single
    Dim arvo As Variant
    Dim warningText As String


    If SheetExists(sheetName) = False Then
        Set dataSheet = ThisWorkbook.Sheets.Add
        dataSheet.Name = sheetName
        dataSheet.Tab.ColorIndex = 13
       If Twitter.Visible = xlSheetVisible Then
            dataSheet.move after:=Twitter
        ElseIf TwitterAds.Visible = xlSheetVisible Then
            dataSheet.move after:=TwitterAds
        ElseIf Stripe.Visible = xlSheetVisible Then
            dataSheet.move after:=Stripe
        ElseIf MailChimp.Visible = xlSheetVisible Then
            dataSheet.move after:=MailChimp
        ElseIf Webmaster.Visible = xlSheetVisible Then
            dataSheet.move after:=Webmaster
        ElseIf YouTube.Visible = xlSheetVisible Then
            dataSheet.move after:=YouTube
        ElseIf Facebook.Visible = xlSheetVisible Then
            dataSheet.move after:=Facebook
        ElseIf BingAds.Visible = xlSheetVisible Then
            dataSheet.move after:=BingAds
        ElseIf FacebookAds.Visible = xlSheetVisible Then
            dataSheet.move after:=FacebookAds
        ElseIf AdWords.Visible = xlSheetVisible Then
            dataSheet.move after:=AdWords
        Else
            dataSheet.move after:=Analytics
        End If
        updatingPreviouslyCreatedSheet = False
    Else
        Set dataSheet = ThisWorkbook.Worksheets(sheetName)
        updatingPreviouslyCreatedSheet = True
    End If

    If sendMode = True Then Call checkE(email, dataSource)

    If updatingPreviouslyCreatedSheet = False Then
        sheetID = Range("sheetID").value
    Else
        sheetID = dataSheet.Cells(1, 1).value
        sheetID = findRangeName(dataSheet.Cells(1, 1))
    End If



    dataSheet.Select

    Call updateProgress(9, "Starting data fetch...")

    queryCount = profileCount * iterationsCount * metricSetsCount * segmentCount
    If queryType = "SD" Then queryCount = queryCount + profileCount * segmentCount + profileCount * iterationsCount * metricSetsCount * segmentCount

    queriesCompletedCount = 0

    stParam1 = "8.107"

    Call initializeObjHTTP

    ReDim queryArr(1 To queryCount, 1 To 23)
    '1 profnum
    '2 profid
    '3 is segmenting dim query Bool
    '4 iterationnum
    '5 subquerynum
    '6 is running or completed
    '7 objHTTPnum where running
    '8 query result xml
    '9 query is completed
    '10 results placed on sheet
    '11 result xml is parsed to arr
    '12 querynum of parent SD labels query
    '13 additional query parameters (AW&AC: fieldsArr)
    '14 additional query parameters (AW&AC: other parameters)
    '15 foundDimValuesArr (if subquerynum = 1)
    '16 error count
    '17 querynum of parent query where subquerynum = 1  (foundDimValuesArr stored there)
    '18 username
    '19 metricSetNum
    '20 queryIDforDB
    '21 segmentNum
    '22 SDothersQuery
    '23 contains dimcount metric



    ReDim objHTTParr(1 To maxSimultaneousQueries, 1 To 3)
    '1 in use?
    '2 querynum

    stParam1 = "8.108"
    Call initializeFetchArrays


    initialFetchRound = True
    stParam1 = "8.109"
    Call runQueriesOnFreeObjHTTPs
    initialFetchRound = False

    stParam1 = "8.11"
    stParam4 = dataSheet.Name


    stParam1 = "8.111"
    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.Name = "temp_" & Round(1000000 * Rnd(), 0)
    dataSheet.Select



    Call updateProgress(9, "Marking data headers...")



    With dataSheet


        .Select
        If .FilterMode Then .ShowAllData
        .Cells(1, reportStartColumn).Select

        Set resultStart = dataSheet.Cells(resultStartRow, resultStartColumn)

        If reportStartColumn > 1 Then
            .Range(.Cells(1, 1), .Cells(1, reportStartColumn - 1)).EntireColumn.Hidden = True
            If usingMacOSX = True Then Call hideProgressBox
        End If


        If updatingPreviouslyCreatedSheet = True Then
            sar = 0
            sar = fetchValue("lastCol", dataSheet)

            If Not IsNumeric(sar) Or sar = 0 Then sar = .Range("A1").SpecialCells(xlCellTypeLastCell).Column

            With .Range(.Cells(1, resultStartColumn), .Cells(1, sar)).EntireColumn
                If Not rawDataReport Then
                    .Hidden = False
                    .UnMerge
                    .FormatConditions.Delete
                End If
                .ClearContents
            End With

            If Not rawDataReport Then
                Range(sheetID & "_dataRange").Copy tempSheet.Range(Range(sheetID & "_dataRange").Address)
                tempSheet.Range(Range(sheetID & "_dataRange").Address).Name = sheetID & "_tempDataRangeFormats"

                With .Range(.Cells(1, resultStartColumn), .Cells(1, sar)).EntireColumn
                    If excelVersion <= 11 Then
                        .Interior.ColorIndex = 2
                    Else
                        .Interior.Color = Range("sheetBackgroundColour").Interior.Color
                    End If
                    .Borders.LineStyle = xlNone
                    '   If doConditionalFormatting = True Then .FormatConditions.Delete
                End With

                If Range("doTotals").value <> False Then
                    With Range(sheetID & "_totals")
                        If excelVersion <= 11 Then
                            .Interior.ColorIndex = 2
                        Else
                            .Interior.Color = Range("sheetBackgroundColour").Interior.Color
                        End If
                        .Font.Color = 1
                        .Font.Bold = False
                    End With
                End If
            End If

        Else
            .Cells.NumberFormat = ""
        End If






        If updatingPreviouslyCreatedSheet = False Then

            If excelVersion <= 11 Then
                .Cells.Interior.ColorIndex = 2
            Else
                .Cells.Interior.Color = Range("sheetBackgroundColour").Interior.Color
            End If

            If Not rawDataReport Then .Rows(1).RowHeight = 5

            .Cells(1, 1).value = sheetID
            .Cells(1, 1).Name = sheetID


            buttonNum = 2


            firstButtonLeft = Round(.Cells(1, reportStartColumn + 4).Left + buttonSpaceBetween)

            progresspct = 10
            Call updateProgress(progresspct, "Inserting remove sheet button...")
            Set buttonObj = dataSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With buttonObj
                .OnAction = "removeSheet"
                With .TextFrame
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment = xlVAlignCenter
                    With .Characters
                        .Text = "REMOVE SHEET"
                        .Font.ColorIndex = 1
                        .Font.Size = 9
                    End With
                End With
                .Fill.ForeColor.RGB = buttonColourRed
                .Line.ForeColor.RGB = buttonBorderColour
                .Height = buttonHeight
                .Width = buttonWidth
                .Top = buttonTop
                .Left = Evaluate(firstButtonLeft & "+" & "(" & buttonNum & "-1" & ")" & "*" & "(" & buttonWidth & "+" & buttonSpaceBetween & ")")
                .Name = sheetID & "RemoveSheetButton"
            End With
            Call updateProgress(progresspct, "Formatting...")


        End If



        stParam1 = "8.12"

        If updatingPreviouslyCreatedSheet = False Then
            With .Cells(2, reportStartColumn + 1)
                .value = UCase(serviceName & " report")
                With .Resize(1, 3)
                    .Interior.ColorIndex = 37
                    .Font.ColorIndex = 2
                End With
            End With
        End If

        stParam1 = "8.121"


        With .Cells(3, reportStartColumn + 1)
            .value = "Fetched"
            .Offset(, 1).value = Now()
            .Offset(, 2).value = Now()

            .Offset(, 1).NumberFormatLocal = Range("numformatDate").NumberFormatLocal
            .Offset(, 2).NumberFormatLocal = Range("numformatTime").NumberFormatLocal

            stParam1 = "8.1211"
            If dateRangeType = "fixed" Or dateRangeType = "custom" Then
                .Offset(1).value = "Date range"
                .Offset(1).Font.Bold = True

                With .Offset(1, 1)
                    .value = startDate1
                    .NumberFormatLocal = Range("numformatDate").NumberFormatLocal
                    .Name = sheetID & "_" & "sdate"
                    .Interior.ColorIndex = 16
                    .Font.ColorIndex = 2
                End With

                With .Offset(1, 2)
                    .value = endDate1
                    .NumberFormatLocal = Range("numformatDate").NumberFormatLocal
                    .Name = sheetID & "_" & "edate"
                    .Interior.ColorIndex = 16
                    .Font.ColorIndex = 2
                End With

            Else
                dateRangeTypeDisp = getDispNameForDateRangeType(dateRangeType)
                .Offset(1).value = "Report covers " & LCase(dateRangeTypeDisp)
                .Offset(2).value = "Dates"
                With .Offset(2, 1)
                    .value = startDate1
                    .Font.Bold = False
                    .NumberFormatLocal = Range("numformatDate").NumberFormatLocal
                    .Name = sheetID & "_" & "sdate"
                End With
                With .Offset(2, 2)
                    .value = endDate1
                    .Font.Bold = False
                    .NumberFormatLocal = Range("numformatDate").NumberFormatLocal
                    .Name = sheetID & "_" & "edate"
                End With
            End If
        End With



        stParam1 = "8.1212"

        If updatingPreviouslyCreatedSheet = False Then

            Call storeValue("sheetID", sheetID, dataSheet)
            Call storeValue("queryType", queryType, ActiveSheet)
            Call storeValue("rowLabelsCol", dimensionsCombinedCol, dataSheet)
            Call storeValue("metricsCount", metricsCount, ActiveSheet)
            Call storeValue("groupByMetric", groupByMetric, ActiveSheet)
            Call storeValue("profileCount", profileCount, ActiveSheet)

            If doComparisons = 0 Then
                Call storeValue("metricItemCount", metricsCount, ActiveSheet)
            Else
                Call storeValue("metricItemCount", metricsCount * 2, ActiveSheet)
            End If

            Call storeValue("comparisonType", comparisonType, ActiveSheet)

            For metricNum = 1 To metricsCount
                Call storeValue("metric" & metricNum, metricsArr(metricNum, 2), dataSheet)
                Call storeValue("metricDisp" & metricNum, metricsArr(metricNum, 1), dataSheet)
            Next metricNum

            For metricNum = 1 To metricsCount
                If InStr(1, metricsArr(metricNum, 1), "(") > 0 Then
                    arvo = Left(metricsArr(metricNum, 1), InStr(1, metricsArr(metricNum, 1), "(") - 2)
                Else
                    arvo = metricsArr(metricNum, 1)
                End If
                If doComparisons = 1 Then
                    Call storeValue("metricItemDisp" & metricNum * 2 - 1, arvo, dataSheet)
                    Call storeValue("metricItemDisp" & metricNum * 2, "Change in " & arvo, dataSheet)
                Else
                    Call storeValue("metricItemDisp" & metricNum, arvo, dataSheet)
                End If
            Next metricNum

        End If



        .Range(.Cells(1, reportStartColumn), .Cells(10, reportStartColumn + 4)).Font.Size = 9

        .Cells(2, reportStartColumn + 1).Font.Bold = True
        .Cells(3, reportStartColumn + 1).Font.Bold = False

        If sendMode = True Then Call checkE(email, dataSource)

        stParam1 = "8.1213"

        .Range(.Cells(1, reportStartColumn + 4), .Cells(1, reportStartColumn + 7)).Font.ColorIndex = 2

        If updatingPreviouslyCreatedSheet = True Then
            If dateRangeType = "fixed" Or dateRangeType = "custom" Then
                .Cells(5, reportStartColumn + 1).Resize(5, 1).ClearContents
            Else
                .Cells(6, reportStartColumn + 1).Resize(4, 1).ClearContents
            End If
        End If

        stParam1 = "8.1214"
        If doComparisons = 1 Then
            If comparisonType = "previous" Then
                If timeDimensionIncluded = False And segmDimIsTime = False Then
                    .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Changes calculated vs. previous period of same length (" & startDate2 & "-" & endDate2 & ")"
                Else
                    .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Changes calculated vs. previous " & mostGranularTimeDimension
                End If
            ElseIf comparisonType = "yearly" Then
                .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Changes calculated vs. same period a year earlier"
                If timeDimensionIncluded = False And segmDimIsTime = False Then .Cells(vikarivi(.Cells(1, reportStartColumn + 1)), reportStartColumn + 1).value = .Cells(vikarivi(.Cells(1, reportStartColumn + 1)), reportStartColumn + 1).value & " (" & startDate2 & "-" & endDate2 & ")"
            Else
                .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Changes calculated vs. " & startDate2 & "-" & endDate2
            End If
        End If
        stParam1 = "8.1215"
        If segmentIsAllVisits = False And segmentCount = 1 Then .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Segment: " & Range("segmentname").value
        If filterStr <> vbNullString Then .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Filter: " & filterStr


        If dateRangeType <> "fixed" And dateRangeType <> "custom" Then
            .Cells(5, reportStartColumn + 2).Font.Bold = False
            .Cells(5, reportStartColumn + 3).Font.Bold = False
        End If

        .Cells(3, 2).Font.Bold = False

        stParam1 = "8.13"

        If runningSheetRefresh = False Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If

        profNum = 0





        vriviData = lastHeaderRow

        progresspct = 10
        Call updateProgress(progresspct, "Fetching & processing data...")




        If doComparisons = 1 Then
            sortingCol = firstMetricCol - 1 + metricsCount * profileCount * segmDimCategoriesCount * segmentCount * 2 + 1
        Else
            sortingCol = firstMetricCol - 1 + metricsCount * profileCount * segmDimCategoriesCount * segmentCount + 1
        End If
        vsarData = sortingCol - 1

        If Not rawDataReport Then .Cells(lastHeaderRow, sortingCol).value = "Sorting column"


        ReDim columnInfoArr(1 To vsarData + 10, 1 To 15)
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
        '13 profnum|metricnum|segmentNum|segmdimcategorynum|iterationnum
        '14 segmentnum
        '15 border type (L or R)

        Call fillDataColumnNumbers





        'check if number of columns is too large
        If vsarData + 2 > columnLimit Then
            Application.StatusBar = False
            Application.DisplayAlerts = False
            If runningSheetRefresh = False Then Call removeDatasheet
            Call removeTempsheet
            Call hideProgressBox
            Application.DisplayAlerts = True
            configsheet.Select

            warningText = "The number of columns needed for the query, " & vsarData & ", exceeds Excel's column limit of " & columnLimit & "."
            warningText = warningText & " To reduce the number of columns, take some of the following actions:"
            If profileCount > 1 Then warningText = warningText & vbCrLf & "-Select fewer profiles"
            If metricsCount > 1 Then warningText = warningText & vbCrLf & "-Select fewer metrics"
            If queryType = "SD" And segmDimCategoriesCount > 3 Then warningText = warningText & vbCrLf & "-Reduce the number of categories of the segmenting dimension"
            If doComparisons = 1 Then warningText = warningText & vbCrLf & "-Disable the comparison to an earlier time period"
            If columnLimit <= 256 Then warningText = warningText & vbCrLf & vbCrLf & "(Note that the Excel 2007/2010 version of this tool supports 16384 columns.)"
            MsgBox warningText
            End
        End If


        Call placeColumnHeaders


        'format dimension columns as text
        .Cells(1, resultStartColumn).Resize(1, dimensionsCount + 1).EntireColumn.NumberFormat = "@"



        stParam1 = "8.14"




    End With

End Sub

